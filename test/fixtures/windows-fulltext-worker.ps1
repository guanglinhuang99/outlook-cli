# Synthetic PowerShell harness for the internal worker helpers.
# It deliberately never loads Outlook or creates a COM object.
param([switch]$BatchProperties)
$ErrorActionPreference = 'Stop'
$script:useBatch = [bool]$BatchProperties
$script:parentReads = 0
$script:batchReads = 0

function Hex-Bytes($Value) {
    return ,[byte[]]@(for ($i = 0; $i -lt $Value.Length; $i += 2) { [Convert]::ToByte($Value.Substring($i, 2), 16) })
}

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
function Get-Mail($Id) { $script:mailLookups++; return ,$script:testMail }

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
    Add-Member -InputObject $accessor -MemberType ScriptMethod -Name GetProperties -Value {
        param($Tags)
        $script:batchReads++
        if (-not $script:useBatch) { throw 'batch properties unavailable' }
        if ($script:batchError) { return ,([object[]]@(-2147221233, -2147221233, -2147221233)) }
        $keyValue = $script:key
        if ($null -eq $keyValue) { $keyValue = -2147221233 }
        $result = [object[]]@($keyValue, (Hex-Bytes $script:mailParent.EntryID), (Hex-Bytes $script:mailParent.StoreID))
        if ($Tags.Count -gt 3) {
            if ($script:metadataErrors) { $result += [object[]]@(-2147221233, -2147221233, -2147221233, -2147221233, -2147221233, -2147221233, -2147221233, -2147221233) }
            else {
                $flags = if ($script:testMail.UnRead) { 0 } else { 1 }
                $result += [object[]]@($script:testMail.Subject,
                    [DateTime]::SpecifyKind($script:testMail.ReceivedTime.ToUniversalTime(), [DateTimeKind]::Unspecified),
                    $flags, [int]$script:testMail.Size, $script:testMail.SenderName, $script:testMail.SenderEmailAddress, $script:testMail.SenderEmailType, -2147221233)
            }
        }
        return ,$result
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
        AccessorValue = New-Accessor
    }
    Add-Member -InputObject $mail -MemberType ScriptProperty -Name PropertyAccessor -Value {
        $script:accessorReads++
        return ,$this.AccessorValue
    }
    Add-Member -InputObject $mail -MemberType ScriptProperty -Name Parent -Value {
        $script:parentReads++
        return ,$script:mailParent
    } -SecondValue { throw 'read-only' }
    Add-Member -InputObject $mail -MemberType ScriptProperty -Name Body -Value {
        $script:bodyReads++
        switch ($script:bodyMode) {
            'null' { return $null }
            'custom' { return $script:customBody }
            'change-key' { $script:key = [byte[]](9, 8, 7); return 'body' }
            'download-change' { $script:testMail.DownloadState = 0; return 'body' }
            'move' { $script:mailParent = $script:sentRoot; return 'body' }
            'delete' { $script:mailParent = $script:deletedRoot; return 'body' }
            'folder-move' { $script:inboxChild.Parent = $script:deletedRoot; return 'body' }
            'ancestor-move' { $script:middleFolder.Parent = $script:deletedRoot; return 'body' }
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
$storeA = New-Store 'AA' $rootsA
$inboxRoot = New-Folder 'Inbox' '01' $storeA $null
$sentRoot = New-Folder 'Sent' '02' $storeA $null
$deletedRoot = New-Folder 'Deleted Items' '03' $storeA $null
$inboxChild = New-Folder 'Child' '04' $storeA $inboxRoot
$rootsA['6'] = $inboxRoot
$rootsA['5'] = $sentRoot
$rootsA['3'] = $deletedRoot
$storeRoot = New-Folder 'Store root' '00' $storeA $null
$script:aboveRootReads = 0
Add-Member -Force -InputObject $storeRoot -MemberType ScriptProperty -Name Parent -Value {
    $script:aboveRootReads++
    throw 'store root has no folder parent'
}
$storeA | Add-Member -MemberType NoteProperty -Name RootFolder -Value $storeRoot
$storeA | Add-Member -MemberType ScriptMethod -Name GetRootFolder -Value { return ,$this.RootFolder }
$inboxRoot.Parent = $storeRoot
$sentRoot.Parent = $storeRoot
$deletedRoot.Parent = $storeRoot

# A second store proves the scope resolves defaults from the requested
# folder's Store rather than from the session's first/default store.
$rootsB = @{}
$storeB = New-Store 'BB' $rootsB
$inboxB = New-Folder 'Inbox' '11' $storeB $null
$deletedB = New-Folder 'Deleted Items' '12' $storeB $null
$outsideB = New-Folder 'Archive' '13' $storeB $null
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
$script:nativeMetadata = ${function:Mail-Metadata}
$script:nativeSender = ${function:Sender-Info}
function Sender-Info($Mail) {
    return @{ name = 'Alice'; address = 'alice@example.test'; address_type = 'SMTP'; smtp_address = 'alice@example.test' }
}
function Person-Match($Mail, $Person, $Role, $SenderInfo = $null) {
    $sender = if ($null -eq $SenderInfo) { Sender-Info $Mail } else { $SenderInfo }
    $matched = $Role -in @('any', 'sender') -and ([string]$Person).ToLowerInvariant() -eq 'alice'
    return @{ matched = $matched; matched_roles = if ($matched) { @('sender') } else { @() }; incomplete = $false; sender = $sender }
}
function Mail-Metadata($Mail, $StoreId, $SenderInfo = $null, $Attachments = $null) {
    switch ($script:metadataMode) {
        'folder-move' { $script:inboxChild.Parent = $script:deletedRoot }
        'ancestor-move' { $script:middleFolder.Parent = $script:deletedRoot }
        'key-change' { $script:key = [byte[]](9, 8, 7) }
        'key-lost' { $script:key = $null }
        'move' { $script:mailParent = $script:sentRoot }
    }
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
$script:accessorReads = 0
[void](FullText-Mail-State $script:testMail $scope (FullText-Accessor $script:testMail))
Assert-True ($script:accessorReads -eq 1) 'state fallback reacquired an available PropertyAccessor'
$script:parentReads = 0
$script:batchReads = 0
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
if ($script:useBatch) {
    Assert-True ($script:batchReads -eq 2 -and $script:parentReads -eq 0) 'binary batch path did not avoid redundant Parent calls'
    $state = FullText-Mail-State $script:testMail $scope (FullText-Accessor $script:testMail) $true
    $sender = & $script:nativeSender $script:testMail $state.metadata_values
    $metadata = & $script:nativeMetadata $script:testMail $scope.store_id $sender $null $state.metadata_values
    Assert-True ($metadata.received_at -eq '2025-01-02T12:00:00.000Z') 'raw UTC delivery time shifted to the host timezone'
    Assert-True ($metadata.subject -eq $script:testMail.Subject -and $metadata.size_bytes -eq 4 -and $metadata.unread) 'typed batched metadata changed its public meaning'
    Assert-True ($metadata.sender.smtp_address -eq 'alice@example.test') 'batched SMTP sender changed'
    $script:metadataErrors = $true
    $state = FullText-Mail-State $script:testMail $scope (FullText-Accessor $script:testMail) $true
    Assert-True (-not $state.metadata_values.ContainsKey('Subject') -and -not $state.metadata_values.ContainsKey('Size')) 'HRESULT became message metadata'
    $sender = & $script:nativeSender $script:testMail $state.metadata_values
    $metadataFallback = & $script:nativeMetadata $script:testMail $scope.store_id $sender $null $state.metadata_values
    Assert-True ($metadataFallback.received_at -eq $metadata.received_at -and $metadataFallback.subject -eq $metadata.subject -and $metadataFallback.size_bytes -eq $metadata.size_bytes) 'unsupported metadata did not use original object-model fields'
    $script:metadataErrors = $false
    $script:batchError = $true
    $fallbackCandidates = FullText-Candidates $scope $candidateParams
    Assert-True ($fallbackCandidates.messages.Count -eq 1 -and $script:parentReads -gt 0) 'per-property HRESULT errors did not use the compatible guard path'
    $script:batchError = $false
}

$inlineParams = [pscustomobject]@{ scan_limit = 50; limit = 50; offset = 0; defer_metadata = $true;
    inline_body_budget = @{ max_bytes = 8; max_body_bytes = 8; budget_ms = 30000 } }
$script:mailLookups = 0
$script:bodyReads = 0
$script:batchReads = 0
$inline = FullText-Candidates $scope $inlineParams
Assert-True ($inline.inline_bodies.Count -eq 1 -and $inline.inline_bodies[0].status -eq 'ready' -and $inline.inline_bodies[0].text -eq 'body') 'inline scan did not return the complete body'
Assert-True ($script:mailLookups -eq 0 -and $script:bodyReads -eq 1) 'inline scan reopened the mail or reread Body'
if ($script:useBatch) { Assert-True ($script:batchReads -eq 2) 'inline scan repeated the pre-body state read' }
$directParams = @{ scan_limit = 50; limit = 50; offset = 0; defer_metadata = $true; direct_scan = $true;
    inline_body_budget = @{ max_bytes = 8; max_body_bytes = 8; budget_ms = 30000 } }
$script:batchReads = 0
$direct = FullText-Candidates $scope $directParams
Assert-True ($direct.inline_bodies[0].status -eq 'unverified') 'direct body was prematurely marked cacheable'
if ($script:useBatch) { Assert-True ($script:batchReads -eq 1) 'direct scan retained the redundant intermediate body guard' }
Assert-True ((FullText-Validate @{ folder = 'inbox'; validation = $direct.validation }).valid) 'final proof did not accept a stable direct scan'
$script:bodyMode = 'custom'
foreach ($case in @(
    @('ASCII body', 'absent', 'unverified_no_match'),
    @('ASCII BODY', 'body', 'unverified'),
    @('ASCII body', 'a*b', 'unverified_no_match'),
    @('', 'body', 'unverified_no_match'),
    @(([string][char]0x0130), 'i', 'unverified'),
    @(([string][char]0x212a), 'k', 'unverified'),
    @(([string][char]0x9879), 'absent', 'unverified_no_match'),
    @(([string][char]0x2019), 'absent', 'unverified_no_match'),
    @(([string][char]0x00c9), 'absent', 'unverified_no_match'),
    @('ASCII body', ([string][char]0x9879), 'unverified')
)) {
    $script:customBody = $case[0]
    $directParams.direct_match_query = $case[1]
    $directParams.inline_body_budget.max_bytes = 100
    $directParams.inline_body_budget.max_body_bytes = 100
    $compact = FullText-Candidates $scope $directParams
    $body = $compact.inline_bodies[0]
    Assert-True ($body.status -eq $case[2]) 'ASCII no-match optimization changed Unicode or hit behavior'
    if ($body.status -eq 'unverified_no_match') {
        Assert-True ($null -eq $body.text -and $body.body_bytes -eq [Text.Encoding]::UTF8.GetByteCount($script:customBody) -and $body.body_code_units -eq $script:customBody.Length -and $body.query -ceq $case[1]) 'compact result lost its byte/query proof'
    } else { Assert-True ($body.text -ceq $script:customBody) 'canonical hit/Unicode body changed' }
    Assert-True ((FullText-Validate @{ folder = 'inbox'; validation = $compact.validation }).valid) 'compact scan skipped final source proof'
}
$directParams.inline_body_budget.max_bytes = 3
$directParams.inline_body_budget.max_body_bytes = 100
$script:customBody = 'body'; $directParams.direct_match_query = 'absent'
$compactCap = FullText-Candidates $scope $directParams
Assert-True ($compactCap.inline_bodies[0].status -eq 'pending') 'omitted body bypassed original byte budget'
$directParams.inline_body_budget.max_bytes = 8; $directParams.inline_body_budget.max_body_bytes = 8
foreach ($mode in @('change-key', 'move', 'delete', 'download-change')) {
    $script:key = [byte[]](1, 2, 3); $script:mailParent = $inboxRoot; $script:testMail.DownloadState = 1
    $script:bodyMode = $mode
    $directChanged = FullText-Candidates $scope $directParams
    $directError = $null
    try { [void](FullText-Validate @{ folder = 'inbox'; validation = $directChanged.validation }) } catch { $directError = $_.Exception.Message }
    Assert-True ($directError -match '^FULLTEXT_SOURCE_CHANGED:') "direct scan final proof missed body mutation: $mode"
}
$script:bodyMode = 'ready'; $script:key = [byte[]](1, 2, 3); $script:mailParent = $inboxRoot; $script:testMail.DownloadState = 1
$inlineParams.inline_body_budget.max_bytes = 3
$limitedInline = FullText-Candidates $scope $inlineParams
Assert-True ($limitedInline.inline_bodies[0].status -eq 'pending' -and $null -eq $limitedInline.inline_bodies[0].text) 'inline scan exceeded its aggregate byte budget'
$inlineParams.inline_body_budget.max_bytes = 8
$script:testMail.DownloadState = 0
$script:bodyReads = 0
$headerInline = FullText-Candidates $scope $inlineParams
Assert-True ($headerInline.inline_bodies[0].status -eq 'missing' -and $script:bodyReads -eq 0) 'inline scan read a header-only body'
$script:testMail.DownloadState = 1
foreach ($mode in @('change-key', 'move', 'delete')) {
    $script:bodyMode = $mode
    $script:key = [byte[]](1, 2, 3)
    $script:mailParent = $inboxRoot
    $changedInline = FullText-Candidates $scope $inlineParams
    Assert-True ($changedInline.inline_bodies[0].status -eq 'changed' -and $null -eq $changedInline.inline_bodies[0].text) "inline scan accepted a body mutation: $mode"
}
$script:bodyMode = 'ready'
foreach ($mode in @('key-change', 'key-lost', 'move')) {
    $script:key = [byte[]](1, 2, 3)
    $script:mailParent = $inboxRoot
    $script:metadataMode = $mode
    $changedInline = FullText-Candidates $scope $inlineParams
    Assert-True ($changedInline.inline_bodies[0].status -eq 'changed') "inline scan blessed a changed metadata version: $mode"
}
$script:metadataMode = $null
$script:key = [byte[]](1, 2, 3)
$script:bodyMode = 'folder-move'
$script:mailParent = $inboxChild
$inboxChild.Items = New-Items @($script:testMail)
$inlineError = $null
try { [void](FullText-Candidates (Resolve-FullText-Scope 'child') $inlineParams) } catch { $inlineError = $_.Exception.Message }
Assert-True ($inlineError -match '^FULLTEXT_SOURCE_CHANGED:') 'inline scan released bodies after the whole folder moved'
$inboxChild.Parent = $inboxRoot
$script:mailParent = $inboxRoot
$script:bodyMode = 'ready'
$script:key = $null
$ephemeralInline = FullText-Candidates $scope $inlineParams
Assert-True ($ephemeralInline.inline_bodies[0].status -eq 'ready' -and $null -eq $ephemeralInline.inline_bodies[0].signature) 'inline scan lost the ephemeral version rules'
$script:key = [byte[]](1, 2, 3)

$validationParams = [pscustomobject]@{ folder = 'inbox'; validation = $candidates.validation }
$script:bodyReads = 0
$validated = FullText-Validate $validationParams
Assert-True ($validated.valid -eq $true -and $script:bodyReads -eq 0) 'warm validation did not pass without reading Body'
$script:testMail.DownloadState = 0
$validationError = $null
try { [void](FullText-Validate $validationParams) } catch { $validationError = $_.Exception.Message }
Assert-True ($validationError -match '^FULLTEXT_SOURCE_CHANGED:') 'warm validation missed changed download availability with an unchanged change key'
$script:testMail.DownloadState = 1
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

# Changes during metadata reads must fail even when the item remains readable.
$script:testMail.Size = 4
$script:bodyMode = 'ready'
foreach ($mode in @('key-change', 'key-lost', 'move')) {
    $script:key = [byte[]](1, 2, 3)
    $script:mailParent = $inboxRoot
    $script:metadataMode = $mode
    $candidateError = $null
    try { [void](FullText-Candidates (Resolve-FullText-Scope 'inbox') $candidateParams) } catch { $candidateError = $_.Exception.Message }
    Assert-True ($candidateError -match '^FULLTEXT_SOURCE_CHANGED:') "candidate scan accepted metadata mutation: $mode"
}
$script:metadataMode = $null
$script:key = [byte[]](1, 2, 3)
$deferredParams = [pscustomobject]@{ scan_limit = 50; limit = 50; offset = 0; defer_metadata = $true }
foreach ($mode in @('key-change', 'key-lost', 'move')) {
    $script:key = [byte[]](1, 2, 3)
    $script:mailParent = $inboxRoot
    $script:metadataMode = $mode
    $deferred = FullText-Candidates (Resolve-FullText-Scope 'inbox') $deferredParams
    Assert-True ($deferred.messages[0]._fulltext_signature -eq '010203|1') 'deferred candidate replaced its original version'
    $validationError = $null
    try { [void](FullText-Validate ([pscustomobject]@{ folder = 'inbox'; validation = $deferred.validation; metadata_ids = @('message|synthetic') })) }
    catch { $validationError = $_.Exception.Message }
    Assert-True ($validationError -match '^FULLTEXT_SOURCE_CHANGED:') "deferred validation accepted metadata mutation: $mode"
}
$script:metadataMode = $null
$script:key = [byte[]](1, 2, 3)
$script:mailParent = $inboxRoot
$batchParams = [pscustomobject]@{
    folder = 'inbox'; candidates = @(@{ id = 'message'; signature = '010203|1' }, @{ id = 'second'; signature = '010203|1' })
    max_bytes = 8; max_body_bytes = 8; budget_ms = 30000
}
$batchReady = FullText-Read-Batch $batchParams
Assert-True ($batchReady.bodies.Count -eq 2 -and $batchReady.bodies[0].text -eq 'body' -and $batchReady.bodies[1].status -eq 'ready') 'batch body read failed'
$batchParams.max_bytes = 4
$batchBounded = FullText-Read-Batch $batchParams
Assert-True ($batchBounded.bodies[0].status -eq 'ready' -and $batchBounded.bodies[1].status -eq 'pending') 'batch byte limit did not preserve pending status'
$batchParams.max_bytes = 8
$script:bodyMode = 'change-key'
$batchChanged = FullText-Read-Batch $batchParams
Assert-True ($batchChanged.bodies[0].status -eq 'changed' -and $batchChanged.bodies[1].status -eq 'changed') 'batch accepted a changed message'
$script:key = [byte[]](1, 2, 3)
$script:bodyMode = 'folder-move'
$script:mailParent = $inboxChild
$batchParams.folder = 'child'
$batchError = $null
try { [void](FullText-Read-Batch $batchParams) } catch { $batchError = $_.Exception.Message }
Assert-True ($batchError -match '^FULLTEXT_SOURCE_CHANGED:') 'batch released bodies from a moved folder'
$inboxChild.Parent = $inboxRoot
$script:bodyMode = 'ready'
$script:mailParent = $inboxRoot
$twinFolder = New-Folder 'Other store Inbox' '01' $storeB $null
$script:mailParent = $twinFolder
$otherStore = FullText-Mail-State $script:testMail $scope (FullText-Accessor $script:testMail)
Assert-True (-not $otherStore.in_scope) 'matching parent EntryID in another store crossed the folder boundary'
$script:mailParent = $inboxChild
$inboxChild.Items = New-Items @($script:testMail)

# The message's parent ID and version stay unchanged while its entire folder
# moves to Deleted. Per-mail parent equality alone must never permit this.
$script:metadataMode = 'folder-move'
$folderError = $null
try { [void](FullText-Candidates (Resolve-FullText-Scope 'child') $candidateParams) } catch { $folderError = $_.Exception.Message }
Assert-True ($folderError -match '^FULLTEXT_SOURCE_CHANGED:') 'candidate scan missed a whole-folder move'
$script:metadataMode = $null
$inboxChild.Parent = $inboxRoot
$script:bodyMode = 'folder-move'
$folderRead = FullText-Read ([pscustomobject]@{ id = 'message'; folder = 'child'; signature = '010203|1' })
Assert-True ($folderRead.status -eq 'changed') 'body read missed a whole-folder move'
$inboxChild.Parent = $inboxRoot
$middleFolder = New-Folder 'Intermediate' '05' $storeA $inboxRoot
$inboxChild.Parent = $middleFolder
$script:metadataMode = 'ancestor-move'
$folderError = $null
try { [void](FullText-Candidates (Resolve-FullText-Scope 'child') $candidateParams) } catch { $folderError = $_.Exception.Message }
Assert-True ($folderError -match '^FULLTEXT_SOURCE_CHANGED:') 'candidate scan missed an ancestor move with unchanged direct parent identity'
$script:metadataMode = $null
$middleFolder.Parent = $inboxRoot
$script:bodyMode = 'ancestor-move'
$ancestorRead = FullText-Read ([pscustomobject]@{ id = 'message'; folder = 'child'; signature = '010203|1' })
Assert-True ($ancestorRead.status -eq 'changed') 'body read missed an ancestor move'
$inboxChild.Parent = $inboxRoot
$script:bodyMode = 'ready'
$childCandidates = FullText-Candidates (Resolve-FullText-Scope 'child') $candidateParams
Add-Member -Force -InputObject $script:testMail -MemberType ScriptProperty -Name DownloadState -Value {
    $script:inboxChild.Parent = $script:deletedRoot
    return 1
}
$folderError = $null
try { [void](FullText-Validate ([pscustomobject]@{ folder = 'child'; validation = $childCandidates.validation })) } catch { $folderError = $_.Exception.Message }
Assert-True ($folderError -match '^FULLTEXT_SOURCE_CHANGED:') 'warm validation missed a whole-folder move'

# All earlier scope/move checks used the real root boundary; store B exercises
# the fallback for providers without GetRootFolder. Never inspect above root.
Assert-True ($script:aboveRootReads -eq 0) 'ancestry walked above the known store root'
Assert-True (Folder-Is-Within $storeRoot $storeRoot $null (Folder-Key $storeRoot)) 'store root equality was incorrectly rejected'
$outsideA = New-Folder 'Archive' '09' $storeA $storeRoot
$script:folderMap['outside-a'] = $outsideA
$scopeError = $null
try { [void](Resolve-FullText-Scope 'outside-a') } catch { $scopeError = $_.Exception.Message }
Assert-True ($scopeError -match '^FULLTEXT_FOLDER_SCOPE:') 'root boundary allowed an unrelated folder'
Assert-True ($script:aboveRootReads -eq 0) 'negative ancestry walked above the known store root'

# Each phase walks a shared ancestor chain once, but the final phase must
# obtain a new proof. Finding Inbox must not stop the Deleted check above it.
$script:walkParentReads = 0
$walkMiddle = New-Folder 'Walk middle' '21' $storeA $inboxRoot
$walkLeaf = New-Folder 'Walk leaf' '22' $storeA $walkMiddle
$walkMiddle | Add-Member -MemberType NoteProperty -Name ActualParent -Value $inboxRoot
$walkLeaf | Add-Member -MemberType NoteProperty -Name ActualParent -Value $walkMiddle
foreach ($node in @($walkMiddle, $walkLeaf)) {
    Add-Member -Force -InputObject $node -MemberType ScriptProperty -Name Parent -Value {
        $script:walkParentReads++
        return ,$this.ActualParent
    }
}
$script:folderMap['walk-leaf'] = $walkLeaf
$walkScope = Resolve-FullText-Scope 'walk-leaf'
Assert-True ($script:walkParentReads -eq 2) 'scope repeated the same shared ancestor Parent reads'
Assert-True (FullText-Scope-Unchanged $walkScope) 'unchanged nested scope failed'
Assert-True ($script:walkParentReads -eq 4) 'final scope check reused an earlier ancestry proof'
$walkMiddle.ActualParent = $deletedRoot
Assert-True (-not (FullText-Scope-Unchanged $walkScope)) 'shared walk missed a moved ancestor'
$walkMiddle.ActualParent = $inboxRoot
$inboxRoot.Parent = $deletedRoot
Assert-True (-not (FullText-Scope-Unchanged $walkScope)) 'shared walk stopped at Inbox before checking Deleted above it'
$inboxRoot.Parent = $storeRoot
$walkMiddle.ActualParent = $walkMiddle
$cycleError = $null
try { [void](Resolve-FullText-Scope 'walk-leaf') } catch { $cycleError = $_.Exception.Message }
Assert-True ($cycleError -match '^FULLTEXT_FOLDER_SCOPE:') 'cyclic ancestry did not fail closed'
Assert-True (-not (FullText-Scope-Unchanged $walkScope)) 'final scope check accepted cyclic ancestry'
$walkMiddle.ActualParent = $inboxRoot

# Read the Exchange sender type once even for remote users and distribution
# lists; retain the same SMTP resolution and PropertyAccessor fallback.
foreach ($entryType in @(0, 5, 1, 99)) {
    $script:senderTypeReads = 0
    $entry = [pscustomobject]@{ TypeValue = $entryType; PrimarySmtpAddress = 'resolved@example.test' }
    Add-Member -InputObject $entry -MemberType ScriptProperty -Name AddressEntryUserType -Value {
        $script:senderTypeReads++; return $this.TypeValue
    }
    Add-Member -InputObject $entry -MemberType ScriptMethod -Name GetExchangeUser -Value { return $this }
    Add-Member -InputObject $entry -MemberType ScriptMethod -Name GetExchangeDistributionList -Value { return $this }
    $senderAccessor = [pscustomobject]@{}
    Add-Member -InputObject $senderAccessor -MemberType ScriptMethod -Name GetProperty -Value { param($Tag); return 'resolved@example.test' }
    $entry | Add-Member -MemberType NoteProperty -Name PropertyAccessor -Value $senderAccessor
    $exchangeMail = [pscustomobject]@{ Sender = $entry; SenderName = 'Sender'; SenderEmailAddress = '/o=example'; SenderEmailType = 'EX' }
    $resolvedSender = & $script:nativeSender $exchangeMail
    Assert-True ($resolvedSender.smtp_address -eq 'resolved@example.test' -and $script:senderTypeReads -eq 1) 'sender optimization changed SMTP resolution or reread its type'
}

[pscustomobject]@{ pass = $true; batch_properties = $script:useBatch; message = 'synthetic full-text worker checks passed' } | ConvertTo-Json -Compress
