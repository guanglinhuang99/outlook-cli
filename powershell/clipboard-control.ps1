#requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('start','stop','status')][string]$Action = 'status',
    [string]$StateDir,
    [switch]$Quiet
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $StateDir) {
    $serviceDir = if ($env:OUTLOOK_LOCAL_STATE_DIR) { $env:OUTLOOK_LOCAL_STATE_DIR } else { Join-Path $projectRoot '.runtime/service' }
    if (-not [IO.Path]::IsPathRooted($serviceDir)) { throw 'OUTLOOK_LOCAL_STATE_DIR must be an absolute path.' }
    $StateDir = Join-Path $serviceDir 'clipboard'
}
$StateDir = [IO.Path]::GetFullPath($StateDir)
$stateFile = Join-Path $StateDir 'clipboard.json'
function Read-BridgeState {
    if (-not [IO.File]::Exists($stateFile)) {
        return [pscustomobject]@{ running = $false; export_dir = (Join-Path $StateDir 'exports') }
    }
    # Permit the worker's atomic File.Replace while reading a stable snapshot.
    for ($readAttempt = 0; $readAttempt -lt 10; $readAttempt++) {
        $inputStream = $null
        $reader = $null
        try {
            $share = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
            $inputStream = [IO.File]::Open($stateFile, [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
            $reader = [IO.StreamReader]::new($inputStream, [Text.Encoding]::UTF8)
            $current = $reader.ReadToEnd() | ConvertFrom-Json
            break
        } catch [IO.IOException] {
            if ($readAttempt -eq 9) { throw }
            Start-Sleep -Milliseconds 30
        } finally {
            if ($null -ne $reader) { $reader.Dispose() }
            elseif ($null -ne $inputStream) { $inputStream.Dispose() }
        }
    }
    $alive = $null -ne (Get-Process -Id $current.pid -ErrorAction SilentlyContinue)
    $current.running = ($current.running -eq $true -and $alive)
    # PowerShell 7 may deserialize ISO dates to DateTime; 5.1 leaves strings.
    $updatedAt = if ($current.updated_at -is [DateTime]) { $current.updated_at } else { [DateTime]::Parse($current.updated_at, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind) }
    $responsive = ([DateTime]::UtcNow - $updatedAt.ToUniversalTime()).TotalSeconds -lt 15
    $current | Add-Member -NotePropertyName responsive -NotePropertyValue $responsive -Force
    return $current
}
function Show-State($value) { if (-not $Quiet) { $value | ConvertTo-Json -Depth 5 } }

$current = Read-BridgeState
if ($Action -eq 'status') { Show-State $current; return }
if ($Action -eq 'stop') {
    if (-not $current.running) { Show-State $current; return }
    if ($current.instance -notmatch '^[a-f0-9-]{36}$') { throw 'Invalid bridge instance identifier.' }
    [IO.File]::WriteAllText((Join-Path $StateDir ($current.instance + '.stop')), '')
    for ($attempt = 0; $attempt -lt 100; $attempt++) {
        Start-Sleep -Milliseconds 150
        $next = Read-BridgeState
        if ($next.instance -eq $current.instance -and -not $next.running) { Show-State $next; return }
    }
    throw 'Stop requested. The bridge is still finishing an Outlook operation; please check status again.'
}
if ($current.running) { Show-State $current; return }

# Keep the standalone default compatible with the Node CLI's managed runtime.
# Only initialize the marker in a new/empty service directory.
if ($serviceDir) {
    $serviceDir = [IO.Path]::GetFullPath($serviceDir)
    [void][IO.Directory]::CreateDirectory($serviceDir)
    $marker = Join-Path $serviceDir '.outlook-local-runtime'
    if (-not [IO.File]::Exists($marker) -and [IO.Directory]::GetFileSystemEntries($serviceDir).Length -eq 0) {
        [IO.File]::WriteAllText($marker, "outlook-local-runtime-v1`n", [Text.UTF8Encoding]::new($false))
    }
}
[void][IO.Directory]::CreateDirectory($StateDir)
$instance = [Guid]::NewGuid().ToString()
$worker = Join-Path $PSScriptRoot 'clipboard.ps1'
$executable = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
$log = Join-Path $StateDir 'clipboard.log'
$argumentsLine = '-NoLogo -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "{0}" -StateDir "{1}" -Instance "{2}"' -f $worker, $StateDir, $instance
$child = Start-Process -FilePath $executable -ArgumentList $argumentsLine -WindowStyle Hidden -PassThru -RedirectStandardError $log -RedirectStandardOutput (Join-Path $StateDir 'clipboard-output.log')
for ($attempt = 0; $attempt -lt 100; $attempt++) {
    Start-Sleep -Milliseconds 150
    $next = Read-BridgeState
    if ($next.instance -eq $instance -and $next.running) { Show-State $next; return }
    $child.Refresh()
    if ($child.HasExited) { break }
}
# Prevent an unusually slow launch from unexpectedly becoming active later.
[IO.File]::WriteAllText((Join-Path $StateDir ($instance + '.stop')), '')
throw "Bridge did not start. Check $log. If another copy is running, stop it from its original folder first."
