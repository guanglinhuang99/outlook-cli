# Synthetic PowerShell harness for the internal worker helpers.
# It deliberately never loads Outlook or creates a COM object.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw $Message }
}

function Track($Value) { return ,$Value }
function Make-Id($Kind, $Entry, $Store) { return "$Kind|$Store|$Entry" }
function Get-Folder($Id) {
    $key = [string]$Id
    if ($script:folderMap.ContainsKey($key)) { return ,$script:folderMap[$key] }
    return $null
}
function Get-Mail($Id) { return ,$script:testMail }

function New-Items($Values) {
    $list = [Collections.ArrayList]::new()
    foreach ($value in @($Values)) { [void]$list.Add($value) }
    $items = [pscustomobject]@{ Values = $list; Clause = $null; SortBy = $null }
    Add-Member -InputObject $items -MemberType ScriptProperty -Name Count -Value { $this.Values.Count }
    Add-Member -InputObject $items -MemberType ScriptMethod -Name Restrict -Value {
        param($Clause)
        $this.Clause = [string]$Clause
        return ,$this
    }
    Add-Member -InputObject $items -MemberType ScriptMethod -Name Sort -Value {
        param($Field, $Descending)
        $this.SortBy = "$Field|$Descending"
    }
    Add-Member -InputObject $items -MemberType ScriptMethod -Name Item -Value {
        param($Index)
        return ,$this.Values[$Index - 1]
    }
    return ,$items
}

function New-Store($StoreId, $Roots) {
    $store = [pscustomobject]@{ StoreID = $StoreId; Roots = $Roots }
    Add-Member -InputObject $store -MemberType ScriptMethod -Name GetDefaultFolder -Value {
        param($Kind)
        $key = [string][int]$Kind
        if (-not $this.Roots.ContainsKey($key)) { throw "missing root $key" }
        return ,$this.Roots[$key]
    }
    return ,$store
}

function New-Folder($Name, $EntryId, $Store, $Parent) {
    $folder = [pscustomobject]@{
        Name = $Name
        EntryID = $EntryId
        StoreID = $Store.StoreID
        Store = $Store
        Parent = $Parent
        Items = $null
        Folders = $null
    }
    $folder.Folders = New-Items @()
    return ,$folder
}

function New-Accessor {
    $accessor = [pscustomobject]@{}
    Add-Member -InputObject $accessor -MemberType ScriptMethod -Name GetProperty -Value {
        param($Tag)
        return ,$script:key
    }
    return ,$accessor
}

function New-Mail($Parent) {
    $mail = [pscustomobject]@{
        Class = 43
        EntryID = 'message-entry'
        StoreID = $Parent.StoreID
        Subject = 'synthetic candidate'
        SenderName = 'Alice'
        SenderEmailAddress = 'alice@example.test'
        SenderEmailType = 'SMTP'
        ReceivedTime = [DateTime]::Parse('2025-01-02T12:00:00Z').ToLocalTime()
        UnRead = $true
        Size = 4
        DownloadState = 1
        LastModificationTime = [DateTime]::Parse('2025-01-02T12:00:00Z')
        Recipients = New-Items @()
        Attachments = New-Items @()
        PropertyAccessor = New-Accessor
    }
    Add-Member -InputObject $mail -MemberType ScriptProperty -Name Parent -Value {
        return ,$script:mailParent
    } -SecondValue { throw 'read-only' }
    Add-Member -InputObject $mail -MemberType ScriptProperty -Name Body -Value {
        $script:bodyReads++
        switch ($script:bodyMode) {
            'null' { return $null }
            'change-key' { $script:key = [byte[]](9, 8, 7); return 'body' }
            'move' { $script:mailParent = $script:sentRoot; return 'body' }
            'delete' { $script:mailParent = $script:deletedRoot; return 'body' }
            'truncated' { return ('x' * 16000001) }
            default { return 'body' }
        }
    } -SecondValue { throw 'read-only' }
    return ,$mail
}

# Load only the helper/function definitions. The worker's JSON loop and its
# Outlook.Application lookup are intentionally excluded from this harness.
$workerPath = Join-Path $PSScriptRoot '../../powershell/worker.ps1'
$workerSource = Get-Content -LiteralPath $workerPath -Raw
$workerStart = $workerSource.IndexOf('function Folder-Key')
$workerEnd = $workerSource.IndexOf('function Run-Operation')
if ($workerStart -lt 0 -or $workerEnd -le $workerStart) { throw 'worker helper boundaries not found' }

$rootsA = @{}
$storeA = New-Store 'store-a' $rootsA
$inboxRoot = New-Folder 'Inbox' 'inbox-root' $storeA $null
$sentRoot = New-Folder 'Sent' 'sent-root' $storeA $null
$deletedRoot = New-Folder 'Deleted Items' 'deleted-root' $storeA $null
$inboxChild = New-Folder 'Child' 'inbox-child' $storeA $inboxRoot
$rootsA['6'] = $inboxRoot
$rootsA['5'] = $sentRoot
$rootsA['3'] = $deletedRoot

# A second store proves the scope resolves defaults from the requested
# folder's Store rather than from the session's first/default store.
$rootsB = @{}
$storeB = New-Store 'store-b' $rootsB
$inboxB = New-Folder 'Inbox' 'inbox-b' $storeB $null
$deletedB = New-Folder 'Deleted Items' 'deleted-b' $storeB $null
$outsideB = New-Folder 'Archive' 'archive-b' $storeB $null
$rootsB['6'] = $inboxB
$rootsB['3'] = $deletedB

$script:folderMap = @{
    inbox = $inboxRoot
    child = $inboxChild
    sent = $sentRoot
    deleted = $deletedRoot
    'store-b-inbox' = $inboxB
    'store-b-deleted' = $deletedB
    'store-b-outside' = $outsideB
}
$script:session = [pscustomobject]@{}
Add-Member -InputObject $script:session -MemberType ScriptMethod -Name GetDefaultFolder -Value {
    param($Kind)
    $key = [string][int]$Kind
    return ,$rootsA[$key]
}

# These replacements keep the harness focused on adapter sequencing while
# leaving scope, guard, and full-text read logic from worker.ps1 intact.
Invoke-Expression $workerSource.Substring($workerStart, $workerEnd - $workerStart)
function Sender-Info($Mail) {
    return @{ name = 'Alice'; address = 'alice@example.test'; address_type = 'SMTP'; smtp_address = 'alice@example.test' }
}
function Person-Match($Mail, $Person, $Role, $SenderInfo = $null) {
    $sender = if ($null -eq $SenderInfo) { Sender-Info $Mail } else { $SenderInfo }
    $matched = $Role -in @('any', 'sender') -and ([string]$Person).ToLowerInvariant() -eq 'alice'
    return @{ matched = $matched; matched_roles = if ($matched) { @('sender') } else { @() }; incomplete = $false; sender = $sender }
}
function Mail-Metadata($Mail, $StoreId, $SenderInfo = $null, $Attachments = $null) {
    return @{ id = 'message|synthetic'; subject = [string]$Mail.Subject; sender = $SenderInfo; received_at = '2025-01-02T12:00:00.000Z'; unread = [bool]$Mail.UnRead; size_bytes = [int]$Mail.Size; attachment_count = 0; download_state = [int]$Mail.DownloadState }
}

$script:mailParent = $inboxRoot
$script:testMail = New-Mail $inboxRoot
$inboxRoot.Items = New-Items @($script:testMail)
$script:key = [byte[]](1, 2, 3)
$script:bodyMode = 'ready'
$script:bodyReads = 0

# Candidate filtering must use received/unread restrictions and person
# matching before any Body access. The query is intentionally irrelevant here.
$scope = Resolve-FullText-Scope 'inbox'
$candidateParams = [pscustomobject]@{
    query = 'query that candidates must ignore'
    unread = $true
    after = '2025-01-01T00:00:00Z'
    before = '2025-01-03T00:00:00Z'
    person = 'alice'
    person_role = 'sender'
    scan_limit = 50
    limit = 50
    offset = 0
}
$candidates = FullText-Candidates $scope $candidateParams
Assert-True ($script:bodyReads -eq 0) 'candidate scan read Body'
Assert-True ($candidates.messages.Count -eq 1) 'candidate received/person filtering did not match the synthetic message'
Assert-True ($inboxRoot.Items.Clause -match '\[UnRead\] = True') 'candidate did not apply unread restriction'
Assert-True ($inboxRoot.Items.Clause -match '\[ReceivedTime\]') 'candidate did not apply received-time restriction'
Assert-True ($candidates.messages[0]._fulltext_offset -eq 1) 'candidate offset was not after the scanned item'

$validationParams = [pscustomobject]@{ folder = 'inbox'; validation = $candidates.validation }
$script:bodyReads = 0
$validated = FullText-Validate $validationParams
Assert-True ($validated.valid -eq $true -and $script:bodyReads -eq 0) 'warm validation did not pass without reading Body'
$validationError = $null
$script:key = [byte[]](9, 8, 7)
$script:bodyReads = 0
try { [void](FullText-Validate $validationParams) } catch { $validationError = $_.Exception.Message }
Assert-True ($validationError -match '^FULLTEXT_SOURCE_CHANGED:') 'warm validation missed a change-key edit'
Assert-True ($script:bodyReads -eq 0) 'warm validation read Body after a change-key edit'
$validationError = $null
$script:key = [byte[]](1, 2, 3)
$script:mailParent = $sentRoot
$script:bodyReads = 0
try { [void](FullText-Validate $validationParams) } catch { $validationError = $_.Exception.Message }
Assert-True ($validationError -match '^FULLTEXT_SOURCE_CHANGED:') 'warm validation missed a move'
Assert-True ($script:bodyReads -eq 0) 'warm validation read Body after a move'
$validationError = $null
$script:mailParent = $deletedRoot
$script:bodyReads = 0
try { [void](FullText-Validate $validationParams) } catch { $validationError = $_.Exception.Message }
Assert-True ($validationError -match '^FULLTEXT_SOURCE_CHANGED:') 'warm validation missed deletion'
Assert-True ($script:bodyReads -eq 0) 'warm validation read Body after deletion'
$script:mailParent = $inboxRoot
$script:key = [byte[]](1, 2, 3)

# The requested store owns its roots, and an Inbox operation remains valid
# when that store has no Sent root. Deleted and unrelated roots fail closed.
$scopeB = Resolve-FullText-Scope 'store-b-inbox'
Assert-True ($scopeB.folder -eq $inboxB) 'scope did not resolve the requested store Inbox'
$scopeError = $null
try { [void](Resolve-FullText-Scope 'store-b-deleted') } catch { $scopeError = $_.Exception.Message }
Assert-True ($scopeError -match '^FULLTEXT_FOLDER_SCOPE:') 'Deleted scope did not use the stable scope error'
$scopeError = $null
try { [void](Resolve-FullText-Scope 'store-b-outside') } catch { $scopeError = $_.Exception.Message }
Assert-True ($scopeError -match '^FULLTEXT_FOLDER_SCOPE:') 'outside scope did not fail closed'

$readParams = [pscustomobject]@{ id = 'message'; folder = 'inbox'; signature = '010203|1' }
$script:bodyMode = 'ready'
$script:bodyReads = 0
$script:key = [byte[]](1, 2, 3)
$script:mailParent = $inboxRoot
$ready = FullText-Read $readParams
Assert-True ($ready.status -eq 'ready' -and $ready.text -eq 'body') 'ready body read failed'

$script:testMail.DownloadState = 0
$script:bodyReads = 0
$headerOnly = FullText-Read ([pscustomobject]@{ id = 'message'; folder = 'inbox'; signature = '010203|0' })
Assert-True ($headerOnly.status -eq 'missing' -and $script:bodyReads -eq 0) "header-only item was read as ready (status=$($headerOnly.status), state=$($script:testMail.DownloadState), reads=$script:bodyReads)"
$script:testMail.DownloadState = 1

$script:bodyMode = 'null'
$script:bodyReads = 0
$nullBody = FullText-Read $readParams
Assert-True ($nullBody.status -eq 'missing' -and $script:bodyReads -eq 1) 'null Body was claimed ready'

$script:bodyMode = 'truncated'
$truncated = FullText-Read $readParams
Assert-True ($truncated.status -eq 'truncated' -and $null -eq $truncated.text) 'truncated body returned text'

$script:bodyMode = 'change-key'
$script:key = [byte[]](1, 2, 3)
$changedKey = FullText-Read $readParams
Assert-True ($changedKey.status -eq 'changed') 'change-key update was not detected'

$script:bodyMode = 'move'
$script:key = [byte[]](1, 2, 3)
$script:mailParent = $inboxRoot
$moved = FullText-Read $readParams
Assert-True ($moved.status -eq 'changed') 'moved item was not detected'

$script:bodyMode = 'delete'
$script:key = [byte[]](1, 2, 3)
$script:mailParent = $inboxRoot
$deleted = FullText-Read $readParams
Assert-True ($deleted.status -eq 'changed') 'deleted item was not detected'
$scopeError = $null
try { [void](FullText-Read ([pscustomobject]@{ id = 'message'; folder = 'deleted'; signature = '010203|1' })) } catch { $scopeError = $_.Exception.Message }
Assert-True ($scopeError -match '^FULLTEXT_FOLDER_SCOPE:') 'deleted requested folder was not rejected'

# A missing change key uses the ephemeral guard only when all three fields
# are present; unavailable fallback fields produce changed before Body read.
$script:bodyMode = 'ready'
$script:key = $null
$script:mailParent = $inboxRoot
$script:testMail.Size = 4
$ephemeralCandidates = FullText-Candidates $scope $candidateParams
$ephemeralGuard = $ephemeralCandidates.messages[0]._fulltext_guard
Assert-True (-not [string]::IsNullOrEmpty([string]$ephemeralGuard)) 'missing change key did not emit the ephemeral candidate guard'
$fallback = FullText-Read ([pscustomobject]@{ id = 'message'; folder = 'inbox'; guard = $ephemeralGuard })
Assert-True ($fallback.status -eq 'ready' -and $null -eq $fallback.signature) 'complete ephemeral guard did not permit a read'
$script:testMail.Size = $null
$script:bodyReads = 0
$incomplete = FullText-Read ([pscustomobject]@{ id = 'message'; folder = 'inbox' })
Assert-True ($incomplete.status -eq 'changed' -and $script:bodyReads -eq 0) 'incomplete ephemeral guard claimed a consistent read'

$readOnly = $false
try { $script:testMail.Body = 'mutation' } catch { $readOnly = $true }
Assert-True $readOnly 'synthetic Body property accepted a write'

[pscustomobject]@{ pass = $true; message = 'synthetic full-text worker checks passed' } | ConvertTo-Json -Compress
