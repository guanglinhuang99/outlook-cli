#requires -Version 5.1
[CmdletBinding()]
param([string]$OutputDir)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $OutputDir) { $OutputDir = Join-Path $root 'dist' }
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
[void][IO.Directory]::CreateDirectory($OutputDir)
$archive = Join-Path $OutputDir 'Outlook-MailCopy-standalone.zip'
# Exact source allowlist: no runtime, copied mail, logs, credentials, or Node.
$files = @(
    'Start-MailCopy.cmd', 'Stop-MailCopy.cmd', 'Status-MailCopy.cmd',
    'powershell/clipboard-control.ps1', 'powershell/clipboard.ps1',
    'powershell/ClipboardBridge.cs', 'docs/mailcopy-standalone.txt'
)
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$stream = [IO.File]::Open($archive, [IO.FileMode]::Create, [IO.FileAccess]::Write)
$zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($relative in $files) {
        $source = Join-Path $root $relative
        if ((Get-Item -LiteralPath $source).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Cannot package a symbolic link.' }
        $target = if ($relative -eq 'docs/mailcopy-standalone.txt') { 'README.txt' } else { $relative }
        [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $source, ('Outlook-MailCopy/' + $target), [IO.Compression.CompressionLevel]::Optimal)
    }
} finally { $zip.Dispose(); $stream.Dispose() }
[pscustomobject]@{ path = $archive; bytes = (Get-Item -LiteralPath $archive).Length; files = $files.Count; sha256 = (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant() } | ConvertTo-Json
