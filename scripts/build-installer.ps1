#requires -Version 5.1
[CmdletBinding()]
param([string]$OutputDir, [string]$NodeArchive, [string]$ReaderArchive)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $OutputDir) { $OutputDir = Join-Path $root 'dist' }
$output = [IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $output -Force | Out-Null
$package = Get-Content -LiteralPath (Join-Path $root 'package.json') -Raw | ConvertFrom-Json
if ($package.version -notmatch '^\d+\.\d+\.\d+$') { throw 'Invalid package version' }
$archivePath = Join-Path $output "outlook-local-$($package.version)-setup.zip"
# Exact allowlist: never include runtime state, tokens, dependencies or mail files.
$files = @('package.json','package-lock.json','pnpm-lock.yaml','README.md','INSTALL-WINDOWS.md','install.ps1','install.cmd','outlook.cmd')
foreach ($dir in @('bin','src','powershell','scripts','test','third-party','docs')) {
    $files += @(Get-ChildItem -LiteralPath (Join-Path $root $dir) -File -Recurse | ForEach-Object { $_.FullName.Substring($root.Length + 1) })
}
# Optional redistributable archives only; never copy an installed runtime or state.
$payload = @()
if ($NodeArchive) {
    $nodeFile = Get-Item -LiteralPath $NodeArchive
    if ($nodeFile.PSIsContainer -or $nodeFile.Name -notmatch '^node-v\d+\.\d+\.\d+-win-(x64|arm64)\.zip$') { throw 'Expected an official Windows Node ZIP filename.' }
    $payload += @{ Path = $nodeFile.FullName; Name = $nodeFile.Name }
}
if ($ReaderArchive) {
    $readerFile = Get-Item -LiteralPath $ReaderArchive
    $metadata = Get-Content -LiteralPath (Join-Path $root 'third-party\ost-mcp.json') -Raw | ConvertFrom-Json
    if ($readerFile.PSIsContainer -or (Get-FileHash -LiteralPath $readerFile.FullName -Algorithm SHA256).Hash.ToLowerInvariant() -ne $metadata.sha256) { throw 'Bundled reader checksum mismatch.' }
    $payload += @{ Path = $readerFile.FullName; Name = 'ost-mcp.zip' }
}
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$stream = [IO.File]::Open($archivePath, [IO.FileMode]::Create, [IO.FileAccess]::Write)
$zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($relative in $files) {
        $full = Join-Path $root $relative
        if ((Get-Item -LiteralPath $full).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Cannot package a symbolic link.' }
        $name = 'OutlookLocal-Setup/' + $relative.Replace('\','/')
        [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $full, $name, [IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
    foreach ($item in $payload) {
        if ((Get-Item -LiteralPath $item.Path).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Cannot package a symbolic link.' }
        [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $item.Path, ('OutlookLocal-Setup/payload/' + $item.Name), [IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
} finally { $zip.Dispose(); $stream.Dispose() }
$hash = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText($archivePath + '.sha256', "$hash  $([IO.Path]::GetFileName($archivePath))`n", [Text.UTF8Encoding]::new($false))
[pscustomobject]@{ Path = $archivePath; Bytes = (Get-Item -LiteralPath $archivePath).Length; Files = $files.Count + $payload.Count; SHA256 = $hash } | ConvertTo-Json
