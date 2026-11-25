# Synthetic COM data only. No real clipboard writes or mailbox access.
param([Parameter(Mandatory=$true)][string]$OutputDir)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -Path @((Join-Path $PSScriptRoot '../../powershell/ClipboardBridge.cs'), (Join-Path $PSScriptRoot 'ClipboardTests.cs')) -ReferencedAssemblies System.Windows.Forms
[OutlookLocalClipboard.Tests]::Run($OutputDir)
'clipboard self-test passed'
