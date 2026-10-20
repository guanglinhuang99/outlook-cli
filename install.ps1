#requires -Version 5.1
<#
User-local installer. Run from the extracted source package, NOT inside the ZIP.
No admin rights, global Node installation or machine policy changes required.
#>
[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'OutlookLocal'),
    [ValidateRange(1024,65535)][int]$Port = 17379,
    [switch]$NoStart,
    [switch]$AutoStart,
    [switch]$SkipLocalReader,
    [string]$NodeArchive
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$utf8 = [Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$nodeVersion = '24.22.0'
$nodeHashes = @{
    x64 = 'f9683a19877e12a9b167305c8db71040d812876359de21f0d2e97c3f321f80c3'
    arm64 = '3a5c209a5bc59e65b3bc0fa13ed5e166a26bd3d55332847f498b76b037408d61'
}
$previousState = $env:OUTLOOK_LOCAL_STATE_DIR
$previousPort = $env:OUTLOOK_LOCAL_PORT
$previousPath = $env:PATH
$locationPushed = $false
$mutex = $null
$ownsMutex = $false

function Write-Utf8($Path, $Text) { [IO.File]::WriteAllText($Path, $Text, $utf8) }
function Run-Node([string[]]$Arguments) {
    & $script:node @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Node command failed (exit $LASTEXITCODE). See output above." }
}
function Is-Within($Parent, $Child) {
    return $Child.StartsWith($Parent.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)
}
function Download-File($Url, $Destination) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing -TimeoutSec 300
}

try {
    if (-not [Environment]::Is64BitOperatingSystem) { throw '64-bit Windows is required.' }
    if ($null -eq [Type]::GetTypeFromProgID('Outlook.Application')) {
        Write-Warning 'Classic Outlook is not registered. COM will be unavailable; local reading requires existing supported OST/PST files.'
    }
    $source = [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')
    $target = [IO.Path]::GetFullPath($InstallDir).TrimEnd('\')
    if ($target.StartsWith('\\') -or $target -eq [IO.Path]::GetPathRoot($target).TrimEnd('\')) { throw 'Choose a local subdirectory, not a network share or drive root.' }
    if ($target -eq $source -or (Is-Within $source $target) -or (Is-Within $target $source)) {
        throw 'The install directory must be separate from the extracted source package.'
    }
    $required = @('package.json', 'package-lock.json', 'README.md', 'install.ps1', 'install.cmd', 'bin\outlook.mjs', 'src\service.mjs', 'powershell\worker.ps1', 'scripts\mcp-setup.txt', 'scripts\check-mcp.mjs')
    foreach ($relative in $required) {
        if (-not (Test-Path -LiteralPath (Join-Path $source $relative) -PathType Leaf)) { throw "Missing package file: $relative. Extract the complete installation ZIP first." }
    }
    $manifest = Get-Content -LiteralPath (Join-Path $source 'package.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($manifest.name -ne 'outlook-local-service') { throw 'This is not an Outlook Local installation package.' }
    $markerPath = Join-Path $target 'installation.json'
    $existing = $null
    if (Test-Path -LiteralPath $target) {
        if ((Get-Item -LiteralPath $target).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Install directory must not be a junction or symbolic link.' }
        if (Test-Path -LiteralPath $markerPath) {
            $existing = Get-Content -LiteralPath $markerPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($existing.product -ne 'outlook-local-service' -or $existing.install_dir -ne $target) { throw 'Directory belongs to another installation.' }
            if (-not $PSBoundParameters.ContainsKey('Port') -and $existing.port) {
                $Port = [int]$existing.port
                if ($Port -lt 1024 -or $Port -gt 65535) { throw 'Stored port is invalid.' }
            }
        } elseif (@(Get-ChildItem -LiteralPath $target -Force).Count) {
            throw 'Destination is not empty and is not a managed Outlook Local installation. Choose a new folder.'
        }
    }
    $digest = [Security.Cryptography.SHA256]::Create()
    try { $mutexName = 'Local\OutlookLocalInstall-' + [BitConverter]::ToString($digest.ComputeHash([Text.Encoding]::UTF8.GetBytes($target.ToLowerInvariant()))).Replace('-','') }
    finally { $digest.Dispose() }
    $mutex = [Threading.Mutex]::new($false, $mutexName)
    try { $ownsMutex = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $ownsMutex = $true }
    if (-not $ownsMutex) { throw 'Another installer is already working on this directory.' }
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    $env:OUTLOOK_LOCAL_STATE_DIR = Join-Path $target '.runtime\service'
    $env:OUTLOOK_LOCAL_PORT = [string]$Port
    # Mark partial installs, so a failed download/npm step can be retried safely.
    $marker = [ordered]@{ product = 'outlook-local-service'; install_dir = $target; version = $manifest.version; status = 'installing'; port = $Port; installed_at = [DateTimeOffset]::Now.ToString('o') }
    if ($existing -and $existing.node_path -and (Test-Path -LiteralPath $existing.node_path) -and (Test-Path -LiteralPath (Join-Path $target 'bin\outlook.mjs'))) {
        if (-not (Is-Within $target ([IO.Path]::GetFullPath($existing.node_path)))) { throw 'Stored Node path is outside this installation.' }
        $script:node = $existing.node_path
        $savedPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        $oldStatus = & $script:node (Join-Path $target 'bin\outlook.mjs') status 2>$null
        $oldExit = $LASTEXITCODE
        $ErrorActionPreference = $savedPreference
        if ($oldExit -eq 0) { Write-Host 'Stopping the previous service...'; Run-Node @((Join-Path $target 'bin\outlook.mjs'), 'stop') }
    }
    if (-not $NoStart) {
        $probe = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $Port)
        try { $probe.Start() } catch { throw "Port $Port is already in use. Stop the other service or use -Port with a free port." } finally { $probe.Stop() }
    }
    Write-Utf8 $markerPath ($marker | ConvertTo-Json)
    Write-Host '[1/5] Installing application files...'
    foreach ($file in @('package.json', 'package-lock.json', 'pnpm-lock.yaml', 'README.md', 'install.ps1', 'install.cmd')) {
        if (-not (Test-Path -LiteralPath (Join-Path $source $file))) { continue }
        Copy-Item -LiteralPath (Join-Path $source $file) -Destination (Join-Path $target $file) -Force
    }
    foreach ($directory in @('bin','src','powershell','scripts','test','third-party')) {
        if (-not (Test-Path -LiteralPath (Join-Path $source $directory))) { continue }
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $source $directory) -File -Recurse) {
            if ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Source package contains a symbolic link.' }
            $relative = $file.FullName.Substring($source.Length + 1)
            $destination = Join-Path $target $relative
            New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
            Copy-Item -LiteralPath $file.FullName -Destination $destination -Force
        }
    }
    Write-Host '[2/5] Preparing a private Node.js runtime...'
    $architecture = $env:PROCESSOR_ARCHITECTURE
    if ($env:PROCESSOR_ARCHITEW6432) { $architecture = $env:PROCESSOR_ARCHITEW6432 }
    $arch = if ($architecture -eq 'ARM64') { 'arm64' } else { 'x64' }
    $nodeFolder = "node-v$nodeVersion-win-$arch"
    $runtime = Join-Path $target 'runtime'
    $nodeHome = Join-Path $runtime $nodeFolder
    $script:node = Join-Path $nodeHome 'node.exe'
    $npm = Join-Path $nodeHome 'node_modules\npm\bin\npm-cli.js'
    New-Item -ItemType Directory -Path $runtime -Force | Out-Null
    if (-not ((Test-Path -LiteralPath $script:node) -and (Test-Path -LiteralPath $npm))) {
        if ($NodeArchive) { $archive = (Resolve-Path -LiteralPath $NodeArchive).Path }
        else {
            $archive = Join-Path $runtime "$nodeFolder.zip"
            if (-not (Test-Path -LiteralPath $archive)) {
                Write-Host "Downloading Node.js v$nodeVersion ($arch) from nodejs.org..."
                Download-File "https://nodejs.org/dist/v$nodeVersion/$nodeFolder.zip" $archive
            }
        }
        if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant() -ne $nodeHashes[$arch]) {
            throw "Node archive checksum mismatch. Do not use it. Remove the invalid ZIP and retry: $archive"
        }
        Expand-Archive -LiteralPath $archive -DestinationPath $runtime -Force
    }
    $actual = & $script:node --version
    if ($LASTEXITCODE -ne 0 -or $actual -ne "v$nodeVersion") { throw 'Private Node runtime verification failed.' }
    $env:PATH = $nodeHome + ';' + $previousPath
    $marker.node_path = $script:node
    Write-Utf8 $markerPath ($marker | ConvertTo-Json)
    Write-Host '[3/5] Installing locked production dependencies...'
    Push-Location -LiteralPath $target
    $locationPushed = $true
    # npm ci verifies package-lock integrity. No third-party lifecycle scripts run.
    Run-Node @($npm, 'ci', '--omit=dev', '--ignore-scripts', '--no-audit', '--no-fund', '--cache', (Join-Path $target '.runtime\npm-cache'))
    if (-not $SkipLocalReader) {
        & (Join-Path $target 'scripts\setup-reader.ps1') -InstallDir $target
    } else { Write-Warning 'Local-file fallback installation skipped; only COM is guaranteed to be available.' }
    Write-Host '[4/5] Creating launchers and MCP configuration...'
    $launcher = @"
@echo off
setlocal DisableDelayedExpansion
set "OUTLOOK_LOCAL_STATE_DIR=%~dp0.runtime\service"
set "OUTLOOK_LOCAL_PORT=$Port"
"%~dp0runtime\$nodeFolder\node.exe" "%~dp0bin\outlook.mjs" %*
"@
    Write-Utf8 (Join-Path $target 'outlook.cmd') $launcher
    foreach ($item in @(@('Start-Outlook.cmd', 'start'), @('Stop-Outlook.cmd', 'stop'))) {
        Write-Utf8 (Join-Path $target $item[0]) ("@echo off`r`ncall `"%~dp0outlook.cmd`" " + $item[1] + "`r`npause`r`n")
    }
    $cli = Join-Path $target 'bin\outlook.mjs'
    $mcp = & $script:node $cli mcp-config
    if ($LASTEXITCODE -ne 0) { throw 'MCP configuration generation failed.' }
    Write-Utf8 (Join-Path $target 'mcp-config.json') ($mcp -join "`n")
    $mcpGuide = Get-Content -LiteralPath (Join-Path $target 'scripts\mcp-setup.txt') -Raw -Encoding UTF8
    Write-Utf8 (Join-Path $target 'MCP-SETUP.txt') ($mcpGuide.Replace('__INSTALL_DIR__', $target).Replace('__NODE_PATH__', $script:node))
    $showMcp = @'
@echo off
setlocal DisableDelayedExpansion
start "" notepad.exe "%~dp0MCP-SETUP.txt"
start "" notepad.exe "%~dp0mcp-config.json"
'@
    Write-Utf8 (Join-Path $target 'Setup-MCP.cmd') $showMcp
    $testMcp = @"
@echo off
setlocal DisableDelayedExpansion
"%~dp0runtime\$nodeFolder\node.exe" "%~dp0scripts\check-mcp.mjs" "%~dp0mcp-config.json" --health
set "MCP_CHECK_EXIT=%ERRORLEVEL%"
echo.
if not "%MCP_CHECK_EXIT%"=="0" echo Check failed. Read the error above and see MCP-SETUP.txt.
pause
exit /b %MCP_CHECK_EXIT%
"@
    Write-Utf8 (Join-Path $target 'Test-MCP.cmd') $testMcp
    Write-Host 'Verifying MCP stdio initialization and tools/list...'
    Run-Node @((Join-Path $target 'scripts\check-mcp.mjs'), (Join-Path $target 'mcp-config.json'))
    $background = @'
$env:OUTLOOK_LOCAL_STATE_DIR = Join-Path $PSScriptRoot '.runtime\service'
$env:OUTLOOK_LOCAL_PORT = '__PORT__'
& (Join-Path $PSScriptRoot 'runtime\__NODE_FOLDER__\node.exe') (Join-Path $PSScriptRoot 'bin\outlook.mjs') start | Out-Null
'@
    Write-Utf8 (Join-Path $target 'Start-Background.ps1') ($background.Replace('__PORT__', [string]$Port).Replace('__NODE_FOLDER__', $nodeFolder))
    if ($AutoStart) {
        # A hidden shortcut in the current user's Startup folder, no elevation.
        $startup = [Environment]::GetFolderPath('Startup')
        $shell = New-Object -ComObject WScript.Shell
        try {
            $shortcut = $shell.CreateShortcut((Join-Path $startup 'Outlook Local Service.lnk'))
            $shortcut.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
            $shortcut.Arguments = '-NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + (Join-Path $target 'Start-Background.ps1') + '"'
            $shortcut.WorkingDirectory = $target
            $shortcut.WindowStyle = 7
            $shortcut.Description = 'Start the local read-only Outlook service after sign-in'
            $shortcut.Save()
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) }
    }
    $marker.status = 'installed'
    Write-Utf8 $markerPath ($marker | ConvertTo-Json)
    Write-Host '[5/5] Checking the service...'
    if (-not $NoStart) {
        Run-Node @($cli, 'start')
        # Health returns no mail bodies. Installation is still valid if Outlook is closed.
        $savedPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        $health = & $script:node $cli health 2>&1
        $healthExit = $LASTEXITCODE
        $ErrorActionPreference = $savedPreference
        if ($healthExit -eq 0) { Write-Host 'Mailbox backend ready (COM or local file). Run outlook.cmd health to see which backend is active.' }
        else { Write-Warning 'Service started, but no mailbox backend is ready. Open classic Outlook or configure local-stores.json, then run outlook.cmd health.' }
    } else { Write-Host 'Startup skipped (-NoStart).' }
    Write-Host ''
    Write-Host 'Installation completed.' -ForegroundColor Green
    Write-Host "Install directory: $target"
    Write-Host "CLI:               $(Join-Path $target 'outlook.cmd')"
    Write-Host "MCP config:        $(Join-Path $target 'mcp-config.json')"
    Write-Host "MCP setup guide:   $(Join-Path $target 'MCP-SETUP.txt')"
    Write-Host "MCP setup helper:  $(Join-Path $target 'Setup-MCP.cmd')"
    Write-Host "MCP connection test: $(Join-Path $target 'Test-MCP.cmd')"
    Write-Host "MCP HTTP:          http://127.0.0.1:$Port/mcp"
    Write-Host 'MCP server installed and protocol check passed. No separate MCP package installation is needed.'
    Write-Host 'Next: double-click Setup-MCP.cmd for the Chinese guide and local configuration.'
    Write-Host 'Import mcp-config.json or merge its outlook-local entry into your agent settings; preserve other servers.'
    Write-Host 'Reconnect your agent, then run Test-MCP.cmd. No AI account/key is required.'
    if (-not $AutoStart) { Write-Host 'Sign-in auto-start was not changed. Use Start-Outlook.cmd after reboot.' }
} catch {
    Write-Host ("Installation failed: " + $_.Exception.Message) -ForegroundColor Red
    Write-Host 'Fix the reported issue and re-run the installer. HTTPS certificate checks are never disabled.'
    exit 1
} finally {
    if ($locationPushed) { Pop-Location }
    $env:OUTLOOK_LOCAL_STATE_DIR = $previousState
    $env:OUTLOOK_LOCAL_PORT = $previousPort
    $env:PATH = $previousPath
    if ($ownsMutex) { $mutex.ReleaseMutex() }
    if ($mutex) { $mutex.Dispose() }
}
