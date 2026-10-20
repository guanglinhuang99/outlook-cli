#requires -Version 5.1
[CmdletBinding()]
param([string]$InstallDir, [string]$Archive)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
if (-not $InstallDir) { $InstallDir = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')) }
$metadata = Get-Content -LiteralPath (Join-Path $InstallDir 'third-party\ost-mcp.json') -Encoding UTF8 -Raw | ConvertFrom-Json
$readerDir = Join-Path $InstallDir 'runtime\ost-mcp'
$exe = Join-Path $readerDir 'ost-mcp.exe'
New-Item -ItemType Directory -Path $readerDir -Force | Out-Null
if (-not ((Test-Path -LiteralPath $exe) -and (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash.ToLowerInvariant() -eq $metadata.exe_sha256)) {
    if (-not $Archive) {
        $Archive = Join-Path $readerDir 'ost-mcp.zip'
        if (-not (Test-Path -LiteralPath $Archive)) {
            Write-Host 'Downloading pinned, read-only OST/PST parser from GitHub...'
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $metadata.url -UseBasicParsing -OutFile $Archive -TimeoutSec 300
        }
    }
    if ((Get-FileHash -LiteralPath $Archive -Algorithm SHA256).Hash.ToLowerInvariant() -ne $metadata.sha256) { throw 'OST/PST reader ZIP checksum mismatch. Remove the invalid ZIP and retry.' }
    Expand-Archive -LiteralPath $Archive -DestinationPath $readerDir -Force
}
if ((Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash.ToLowerInvariant() -ne $metadata.exe_sha256) { throw 'OST/PST reader executable checksum mismatch.' }
$version = & $exe --version
if ($LASTEXITCODE -ne 0 -or $version -ne ('ost-mcp ' + $metadata.version)) { throw 'Cannot run x64 OST/PST parser. ARM64 Windows requires x64 emulation; use -SkipLocalReader for COM-only installation.' }
Write-Host ('Local OST/PST reader ready: ' + $version)
