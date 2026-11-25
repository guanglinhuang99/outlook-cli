#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$StateDir, [Parameter(Mandatory=$true)][string]$Instance)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$stateFile = Join-Path $StateDir 'clipboard.json'
$stopFile = Join-Path $StateDir ($Instance + '.stop')
$exportDir = Join-Path $StateDir 'exports'
$mutex = $null
$bridge = $null
$owned = $false
function Publish-State {
    $state.updated_at = [DateTime]::UtcNow.ToString('o')
    $temporary = Join-Path $StateDir ($Instance + '.json.tmp')
    [IO.File]::WriteAllText($temporary, ($state | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
    for ($writeAttempt = 0; $writeAttempt -lt 20; $writeAttempt++) {
        try {
            if ([IO.File]::Exists($stateFile)) { [IO.File]::Replace($temporary, $stateFile, [NullString]::Value) }
            else { [IO.File]::Move($temporary, $stateFile) }
            return
        } catch [IO.IOException] {
            if ($writeAttempt -eq 19) { throw }
            Start-Sleep -Milliseconds 30
        }
    }
}
try {
    # One bridge per interactive logon session, even from different checkouts.
    $created = $false
    $mutex = [Threading.Mutex]::new($true, 'Local\OutlookLocalClipboardBridge', [ref]$created)
    if (-not $created) { throw 'A clipboard bridge is already running in this Windows session.' }
    $owned = $true
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -Path (Join-Path $PSScriptRoot 'ClipboardBridge.cs') -ReferencedAssemblies System.Windows.Forms
    [void][IO.Directory]::CreateDirectory($exportDir)
    $bridge = [OutlookLocalClipboard.Bridge]::new($exportDir)
    $state = [ordered]@{
        instance = $Instance; pid = $PID; running = $true
        started_at = [DateTime]::UtcNow.ToString('o'); updated_at = ''
        export_dir = $exportDir; exported_count = 0; last_export_at = $null
        last_files = @(); last_error = $null
    }
    Publish-State
    $heartbeat = [Diagnostics.Stopwatch]::StartNew()
    while (-not [IO.File]::Exists($stopFile)) {
        try {
            $files = $bridge.Poll()
            if ($null -ne $files -and $files.Count -gt 0) {
                $state.last_files = @($files)
                $state.exported_count += $files.Count
                $state.last_export_at = [DateTime]::UtcNow.ToString('o')
                $state.last_error = $null
                Publish-State
            }
        } catch {
            $state.last_error = $_.Exception.GetBaseException().Message
            Publish-State
        }
        if ($heartbeat.ElapsedMilliseconds -ge 1000) { Publish-State; $heartbeat.Restart() }
        Start-Sleep -Milliseconds 150
    }
} finally {
    if ($null -ne $bridge) { $bridge.Dispose() }
    if ($owned) {
        if ($null -ne $state) { $state.running = $false; Publish-State }
        $mutex.ReleaseMutex()
    }
    if ($null -ne $mutex) { $mutex.Dispose() }
    if ([IO.File]::Exists($stopFile)) { [IO.File]::Delete($stopFile) }
}
