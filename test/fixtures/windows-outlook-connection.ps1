# Exercise real connection logic with synthetic Outlook objects; never launch Outlook.
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../../powershell/worker.ps1') -Raw
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$connect = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Connect-Outlook' }, $true)
Invoke-Expression $connect.Extent.Text
function Assert-True($Value, $Message) { if (-not $Value) { throw $Message } }
function Track($Value) { return ,$Value }
function Get-RunningOutlook {
    $script:gets++
    if (-not $script:running) { throw 'not running' }
    return $script:fakeApp
}
function Start-Outlook {
    $script:starts++
    if ($script:activationFails) { throw 'activation failed' }
    return $script:fakeApp
}
function Release-OutlookApplication { $script:releases++; $script:application = $null }
function Write-ConnectionPhase($Phase) { $script:phases.Add($Phase) }
function Reset-Scenario {
    $script:AutoStart = '1'
    $script:application = $null
    $script:session = $null
    $script:running = $false
    $script:activationFails = $false
    $script:initializationFails = $false
    $script:namespaceFails = $false
    $script:gets = 0; $script:starts = 0; $script:inboxReads = 0; $script:releases = 0
    $script:phases = [Collections.Generic.List[string]]::new()
    $script:fakeSession = [pscustomobject]@{}
    $script:fakeSession | Add-Member ScriptMethod GetDefaultFolder {
        param($Kind)
        Assert-True ($Kind -eq 6) 'must initialize with the default inbox'
        $script:inboxReads++
        if ($script:initializationFails) { throw 'profile unavailable' }
        return [pscustomobject]@{}
    }
    $script:fakeApp = [pscustomobject]@{ Version = '16.0' }
    $script:fakeApp | Add-Member ScriptMethod GetNamespace {
        param($Name)
        Assert-True ($Name -eq 'MAPI') 'wrong namespace'
        if ($script:namespaceFails) { throw 'existing session is busy' }
        return $script:fakeSession
    }
}
function Assert-Failure($Code) {
    $message = ''
    try { Connect-Outlook } catch { $message = $_.Exception.Message }
    Assert-True ($message.StartsWith($Code + ':')) "expected $Code, got $message"
}

Reset-Scenario
$script:running = $true
Connect-Outlook
Assert-True ($script:starts -eq 0 -and $script:connection -eq 'existing') 'running Outlook must be reused'
Assert-True ($script:phases.Count -eq 0) 'existing session is not startup'

Reset-Scenario
Connect-Outlook
Assert-True ($script:starts -eq 1 -and $script:connection -eq 'started') 'closed Outlook must activate by default'
Assert-True ($script:inboxReads -eq 1) 'default profile was not initialized'
Assert-True (($script:phases -join ',') -eq 'starting,ready') 'startup phases incorrect'
Connect-Outlook
Assert-True ($script:starts -eq 1 -and $script:gets -eq 1 -and $script:connection -eq 'reused') 'application must stay alive between calls'
# Simulate a stale root after the user exits Outlook.
$script:application = [pscustomobject]@{ Version = '16.0' }
$script:application | Add-Member ScriptMethod GetNamespace { throw 'disconnected' }
Connect-Outlook
Assert-True ($script:starts -eq 2 -and $script:releases -eq 1) 'disconnected application did not reconnect'

Reset-Scenario
$script:AutoStart = '0'
Assert-Failure 'OUTLOOK_UNAVAILABLE'
Assert-True ($script:starts -eq 0) 'disabled startup still activated Outlook'

Reset-Scenario
$script:activationFails = $true
Assert-Failure 'OUTLOOK_START_FAILED'
Assert-True ($null -eq $script:application) 'failed activation retained application'

Reset-Scenario
$script:initializationFails = $true
Assert-Failure 'OUTLOOK_START_FAILED'
Assert-True ($null -eq $script:application -and ($script:phases -join ',') -eq 'starting') 'profile failure marked ready'

Reset-Scenario
$script:running = $true
$script:namespaceFails = $true
$message = ''
try { Connect-Outlook } catch { $message = $_.Exception.Message }
Assert-True ($message -like '*existing session is busy*' -and $message -notlike 'OUTLOOK_START_FAILED:*') 'existing session error became a startup failure'
Assert-True ($script:starts -eq 0) 'busy existing session launched Outlook'
@{ pass = $true; scenarios = 8 } | ConvertTo-Json -Compress
