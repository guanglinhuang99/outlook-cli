# Persistent JSON-lines worker. Runs only fixed read operations on a single STA thread.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::InputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
$script:idPartsCache = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
$script:encodedIdCache = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)

function Remember-Id($Id, $Kind, $Entry, $Store) {
    # Only immutable identifier components are cached, never COM objects,
    # versions or folder membership. Bound memory in the persistent worker.
    if ($Id.Length -gt 16384 -or $Entry -notmatch '^[a-fA-F0-9]+$' -or $Store -notmatch '^[a-fA-F0-9]+$') { return }
    if ($script:idPartsCache.Count -ge 256) { $script:idPartsCache.Clear() }
    $script:idPartsCache[$Id] = @([string]$Kind, [string]$Entry, [string]$Store)
}

function Track($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { $script:refs.Add($Value) }
    return ,$Value
}
function Make-Id($Kind, $Entry, $Store) {
    $key = [string]$Kind + '|' + [string]$Entry + '|' + [string]$Store
    $cached = $null
    if ($script:encodedIdCache.TryGetValue($key, [ref]$cached)) {
        Remember-Id $cached $Kind ([string]$Entry) ([string]$Store)
        return $cached
    }
    $json = @{ v = 1; kind = $Kind; entry = [string]$Entry; store = [string]$Store } | ConvertTo-Json -Compress
    $id = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json)).TrimEnd('=').Replace('+','-').Replace('/','_')
    Remember-Id $id $Kind ([string]$Entry) ([string]$Store)
    if ($script:idPartsCache.ContainsKey($id)) {
        if ($script:encodedIdCache.Count -ge 256) { $script:encodedIdCache.Clear() }
        $script:encodedIdCache[$key] = $id
    }
    return $id
}
function Parse-Id($Id, $Kind) {
    try {
        $cached = $null
        if ($script:idPartsCache.TryGetValue($Id, [ref]$cached)) {
            if ($cached[0] -ne $Kind) { throw 'bad id kind' }
            return @{ v = 1; kind = $cached[0]; entry = $cached[1]; store = $cached[2] }
        }
        $b64 = $Id.Replace('-','+').Replace('_','/')
        $b64 = $b64.PadRight($b64.Length + ((4 - $b64.Length % 4) % 4), '=')
        $obj = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b64)) | ConvertFrom-Json
        if ($obj.v -ne 1 -or $obj.kind -ne $Kind -or $obj.entry -notmatch '^[a-fA-F0-9]+$' -or $obj.store -notmatch '^[a-fA-F0-9]+$') { throw 'bad id' }
        Remember-Id $Id $Kind ([string]$obj.entry) ([string]$obj.store)
        return $obj
    } catch { throw 'INVALID_ID: Invalid or mismatched Outlook identifier' }
}
function Get-Folder($Id) {
    $defaults = @{ inbox = 6; sent = 5; drafts = 16; deleted = 3; junk = 23; outbox = 4 }
    if ($defaults.ContainsKey($Id)) { return ,(Track ($script:session.GetDefaultFolder($defaults[$Id]))) }
    $decoded = Parse-Id $Id 'folder'
    return ,(Track ($script:session.GetFolderFromID($decoded.entry, $decoded.store)))
}
function Get-Mail($Id) {
    $decoded = Parse-Id $Id 'message'
    $mail = Track ($script:session.GetItemFromID($decoded.entry, $decoded.store))
    if ($mail.Class -ne 43) { throw 'UNSUPPORTED_ITEM: This item is not an email' }
    return ,$mail
}
function Folder-Key($Folder) {
    try {
        $entry = [string]$Folder.EntryID
        $store = [string]$Folder.StoreID
        if ([string]::IsNullOrEmpty($entry) -or [string]::IsNullOrEmpty($store)) { return $null }
        return $store + '|' + $entry
    } catch { return $null }
}
function FullText-Folder-Ancestry($Folder, $InitialKey, $StoreRootKey, $StoreRoot = $null) {
    # One fresh walk supplies both the allowed-root and Deleted checks. Never
    # carry this proof across phases: the end-of-operation walk starts anew.
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $current = $Folder
    for ($depth = 0; $depth -lt 128 -and $null -ne $current; $depth++) {
        if ($null -ne $StoreRootKey -and $null -ne $StoreRoot -and [object]::ReferenceEquals($current, $StoreRoot)) {
            [void]$seen.Add($StoreRootKey)
            return ,$seen
        }
        $key = if ($depth -eq 0 -and $null -ne $InitialKey) { $InitialKey } else { Folder-Key $current }
        # Providers without GetRootFolder may terminate at a non-folder Parent,
        # just as the original ancestry walk did.
        if ($null -eq $key) { return ,$seen }
        if (-not $seen.Add($key)) { return $null }
        if ($null -ne $StoreRootKey -and [string]::Equals($key, $StoreRootKey, [StringComparison]::OrdinalIgnoreCase)) { return ,$seen }
        try { $current = Track ($current.Parent) } catch { return ,$seen }
    }
    if ($null -ne $current) { return $null }
    return ,$seen
}
function Folder-Is-Within($Folder, $Root, $InitialKey = $null, $StoreRootKey = $null, $Ancestry = $null) {
    $rootKey = if ($null -ne $InitialKey -and [object]::ReferenceEquals($Folder, $Root)) { $InitialKey } else { Folder-Key $Root }
    if ($null -eq $rootKey) { return $false }
    if ($null -ne $Ancestry) { return $Ancestry.Contains($rootKey) }
    $seen = @{}
    $current = $Folder
    for ($depth = 0; $depth -lt 128 -and $null -ne $current; $depth++) {
        $key = if ($depth -eq 0 -and $null -ne $InitialKey) { $InitialKey } else { Folder-Key $current }
        if ($null -eq $key) { return $false }
        if ([string]::Equals($key, $rootKey, [StringComparison]::OrdinalIgnoreCase)) { return $true }
        # The store root has no folder ancestor. Avoid probing EntryID/StoreID
        # on its non-folder Parent (expensive COM type discovery on first use).
        if ($null -ne $StoreRootKey -and [string]::Equals($key, $StoreRootKey, [StringComparison]::OrdinalIgnoreCase)) { return $false }
        if ($seen.ContainsKey($key)) { return $false }
        $seen[$key] = $true
        try { $current = Track ($current.Parent) } catch { return $false }
    }
    return $false
}
function Resolve-FullText-Scope($FolderId) {
    if ([string]::IsNullOrWhiteSpace([string]$FolderId)) { throw 'INVALID_ARGUMENT: folder is required for full-text operations' }
    try { $folder = Get-Folder $FolderId } catch { throw 'INVALID_ID: Folder not found' }
    if ($null -eq $folder) { throw 'INVALID_ID: Folder not found' }
    try { $store = Track ($folder.Store) } catch { $store = $null }
    if ($null -eq $store) { throw 'FULLTEXT_FOLDER_SCOPE: Full-text operations require a known Outlook store' }
    $storeRootKey = $null
    $storeRoot = $null
    try { $storeRoot = Track ($store.GetRootFolder()); $storeRootKey = Folder-Key $storeRoot } catch { }
    $inbox = $null
    $sent = $null
    $deleted = $null
    try { $inbox = Track ($store.GetDefaultFolder(6)) } catch { }
    try { $deleted = Track ($store.GetDefaultFolder(3)) } catch { }
    # Deleted is a safety boundary. A missing Sent root is allowed when the
    # requested folder is in the same store's Inbox subtree.
    if ($null -eq $deleted) { throw 'FULLTEXT_FOLDER_SCOPE: Full-text operations require a known Deleted Items root' }
    $folderKey = Folder-Key $folder
    if ($null -eq $folderKey) { throw 'FULLTEXT_FOLDER_SCOPE: Full-text operations require a known folder identity' }
    $ancestry = FullText-Folder-Ancestry $folder $folderKey $storeRootKey $storeRoot
    if ($null -eq $ancestry) { throw 'FULLTEXT_FOLDER_SCOPE: Full-text folder ancestry is cyclic or exceeds its bound' }
    if (Folder-Is-Within $folder $deleted $folderKey $storeRootKey $ancestry) { throw 'FULLTEXT_FOLDER_SCOPE: Full-text operations exclude Deleted Items and its descendants' }
    $root = $null
    if ($null -ne $inbox -and (Folder-Is-Within $folder $inbox $folderKey $storeRootKey $ancestry)) { $root = $inbox }
    else {
        try { $sent = Track ($store.GetDefaultFolder(5)) } catch { }
        if ($null -ne $sent -and (Folder-Is-Within $folder $sent $folderKey $storeRootKey $ancestry)) { $root = $sent }
    }
    if ($null -eq $root) { throw 'FULLTEXT_FOLDER_SCOPE: Full-text operations are limited to the default Inbox and Sent subtrees' }
    # The identity just captured is the scope's version, not a cache. Reuse it
    # for serialization and compare with fresh identity at operation end.
    $folderParts = $folderKey -split '\|', 2
    return ,@{
        folder = $folder
        folder_key = $folderKey
        store_root_key = $storeRootKey
        store_root = $storeRoot
        store_id = $folderParts[0]
        id = Make-Id 'folder' $folderParts[1] $folderParts[0]
        root = $root
        deleted = $deleted
    }
}
function Mail-In-FullText-Scope($Mail, $Scope) {
    # Resolve-FullText-Scope proves the requested folder's ancestry once per
    # operation. Each mail still needs a fresh parent identity check, but not
    # repeated walks through that identical ancestry. Recheck the folder proof
    # at the operation's end to detect a folder/subtree move during the scan.
    try { $current = Track ($Mail.Parent) } catch { return $false }
    $key = Folder-Key $current
    return $null -ne $key -and [string]::Equals($key, $Scope.folder_key, [StringComparison]::OrdinalIgnoreCase)
}
function FullText-Scope-Unchanged($Scope) {
    $key = Folder-Key $Scope.folder
    if ($null -eq $key -or -not [string]::Equals($key, $Scope.folder_key, [StringComparison]::OrdinalIgnoreCase)) { return $false }
    $ancestry = FullText-Folder-Ancestry $Scope.folder $key $Scope.store_root_key $Scope.store_root
    return $null -ne $ancestry -and
        -not (Folder-Is-Within $Scope.folder $Scope.deleted $key $Scope.store_root_key $ancestry) -and
        (Folder-Is-Within $Scope.folder $Scope.root $key $Scope.store_root_key $ancestry)
}
function FullText-Mail-State($Mail, $Scope, $Accessor, $IncludeMetadata = $false) {
    # GetProperties returns individual error values for unsupported properties.
    # Only use the batch fast path for three nonempty binary values; otherwise
    # fall back to the object model, including its complete ephemeral guard.
    try {
        $tags = [object[]]@(
            'http://schemas.microsoft.com/mapi/proptag/0x65E20102',
            'http://schemas.microsoft.com/mapi/proptag/0x0E090102',
            'http://schemas.microsoft.com/mapi/proptag/0x0FFB0102'
        )
        if ($IncludeMetadata) {
            $tags += [object[]]@(
                'http://schemas.microsoft.com/mapi/proptag/0x0037001F', # Subject
                'http://schemas.microsoft.com/mapi/proptag/0x0E060040', # Delivery time (UTC)
                'http://schemas.microsoft.com/mapi/proptag/0x0E070003', # Message flags
                'http://schemas.microsoft.com/mapi/proptag/0x0E080003', # Size
                'http://schemas.microsoft.com/mapi/proptag/0x0C1A001F', # Sender name
                'http://schemas.microsoft.com/mapi/proptag/0x0C1F001F', # Sender address
                'http://schemas.microsoft.com/mapi/proptag/0x0C1E001F', # Sender address type
                'http://schemas.microsoft.com/mapi/proptag/0x0FFF0102'  # Entry ID
            )
        }
        $values = $Accessor.GetProperties($tags)
        if ($values.Count -eq $tags.Count -and
            $values[0] -is [byte[]] -and $values[0].Length -gt 0 -and
            $values[1] -is [byte[]] -and $values[1].Length -gt 0 -and
            $values[2] -is [byte[]] -and $values[2].Length -gt 0) {
            $state = $Mail.DownloadState
            if ($null -ne $state) {
                $change = [BitConverter]::ToString($values[0]).Replace('-', '').ToLowerInvariant()
                $parent = [BitConverter]::ToString($values[1]).Replace('-', '')
                $store = [BitConverter]::ToString($values[2]).Replace('-', '')
                $sameFolder = [string]::Equals($store + '|' + $parent, $Scope.folder_key, [StringComparison]::OrdinalIgnoreCase)
                # Providers may represent equivalent EntryIDs differently.
                # Preserve the object-model identity check on a non-match.
                if (-not $sameFolder) { $sameFolder = Mail-In-FullText-Scope $Mail $Scope }
                $metadata = @{ DownloadState = [int]$state }
                if ($IncludeMetadata) {
                    # Never convert an individual HRESULT into user data. Each
                    # unsupported field is read through the object model later.
                    foreach ($field in @(@(3, 'Subject'), @(7, 'SenderName'), @(8, 'SenderEmailAddress'), @(9, 'SenderEmailType'))) {
                        if ($values[$field[0]] -is [string]) { $metadata[$field[1]] = $values[$field[0]] }
                    }
                    if ($values[4] -is [DateTime]) { $metadata.ReceivedTime = [DateTime]::SpecifyKind($values[4], [DateTimeKind]::Utc).ToLocalTime() }
                    if ($values[5] -is [int] -and $values[5] -ge 0) { $metadata.UnRead = ($values[5] -band 1) -eq 0 }
                    if ($values[6] -is [int] -and $values[6] -ge 0) { $metadata.Size = $values[6] }
                    if ($values[10] -is [byte[]] -and $values[10].Length -gt 0) { $metadata.EntryID = [BitConverter]::ToString($values[10]).Replace('-', '') }
                }
                return @{ signature = $change + '|' + [string][int]$state; ephemeral_guard = $null; download_state = [int]$state; in_scope = $sameFolder; metadata_values = $metadata }
            }
        }
    } catch { }
    $guard = Mail-Guard $Mail $Accessor
    $guard.in_scope = Mail-In-FullText-Scope $Mail $Scope
    return $guard
}
function FullText-Accessor($Mail) {
    try { return ,(Track ($Mail.PropertyAccessor)) } catch { return $null }
}
function Change-Key-Hex($Mail, $Accessor = $null) {
    try {
        if ($null -eq $Accessor) { $Accessor = Track ($Mail.PropertyAccessor) }
        $value = $accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x65E20102')
        if ($value -is [byte[]] -and $value.Length -gt 0) {
            return [BitConverter]::ToString($value).Replace('-', '').ToLowerInvariant()
        }
        if ($value -is [Array] -and $value.Length -gt 0) {
            $bytes = New-Object byte[] $value.Length
            for ($i = 0; $i -lt $value.Length; $i++) { $bytes[$i] = [byte]$value.GetValue($i) }
            return [BitConverter]::ToString($bytes).Replace('-', '').ToLowerInvariant()
        }
    } catch { }
    return $null
}
function Mail-Guard($Mail, $Accessor = $null) {
    $changeKey = Change-Key-Hex $Mail $Accessor
    $downloadState = $null
    $lastModification = $null
    $size = $null
    $downloadStateValue = $null
    $lastModificationValue = $null
    $sizeValue = $null
    try { $downloadStateValue = $Mail.DownloadState } catch { }
    if ($null -ne $downloadStateValue) { try { $downloadState = [int]$downloadStateValue } catch { } }
    if (-not [string]::IsNullOrEmpty($changeKey) -and $null -ne $downloadState) {
        # A versioned message never needs the weaker timestamp/size fallback.
        # Avoid those COM reads and date conversion on every warm guard check.
        return @{ signature = $changeKey + '|' + [string]$downloadState; ephemeral_guard = $null; download_state = $downloadState }
    }
    try { $lastModificationValue = $Mail.LastModificationTime } catch { }
    try { $sizeValue = $Mail.Size } catch { }
    if ($null -ne $lastModificationValue) { try { $lastModification = Iso-Date $lastModificationValue } catch { } }
    if ($null -ne $sizeValue) { try { $size = [string]$sizeValue } catch { } }
    $stateText = if ($null -eq $downloadState) { '<missing>' } else { [string]$downloadState }
    $lastText = if ($null -eq $lastModification) { '<missing>' } else { [string]$lastModification }
    $sizeText = if ($null -eq $size) { '<missing>' } else { [string]$size }
    $ephemeralGuard = $null
    if ($null -ne $lastModification -and $null -ne $size -and $null -ne $downloadState) {
        $ephemeralGuard = $lastText + '|' + $sizeText + '|' + $stateText
    }
    return @{
        signature = $null
        ephemeral_guard = $ephemeralGuard
        download_state = $downloadState
    }
}
function FullText-Guard-Matches($Before, $After, $UseSignature) {
    if ($UseSignature) {
        return $null -ne $Before.signature -and $null -ne $After.signature -and [string]::Equals([string]$Before.signature, [string]$After.signature, [StringComparison]::Ordinal)
    }
    if ($null -eq $Before.ephemeral_guard -or $null -eq $After.ephemeral_guard) { return $false }
    return [string]::Equals([string]$Before.ephemeral_guard, [string]$After.ephemeral_guard, [StringComparison]::Ordinal)
}
function FullText-Result($Status, $Text, $Signature) {
    return @{ status = $Status; text = $Text; signature = $Signature }
}
function FullText-Filtered-Items($Items, $Params) {
    $clauses = [Collections.Generic.List[string]]::new()
    if ($null -ne $Params.unread) {
        if ($Params.unread) { $clauses.Add('[UnRead] = True') } else { $clauses.Add('[UnRead] = False') }
    }
    foreach ($spec in @(@('after','>='), @('before','<'))) {
        $value = $Params.($spec[0])
        if ($value) {
            $local = [DateTimeOffset]::Parse($value).LocalDateTime
            $rounded = $local.AddTicks(-($local.Ticks % [TimeSpan]::TicksPerMinute))
            if ($spec[0] -eq 'before' -and $rounded -lt $local) { $rounded = $rounded.AddMinutes(1) }
            $date = $rounded.ToString('g', [Globalization.CultureInfo]::CurrentCulture)
            $clauses.Add("[ReceivedTime] $($spec[1]) '$date'")
        }
    }
    if ($clauses.Count) { return ,(Track ($Items.Restrict(($clauses -join ' AND ')))) }
    return ,$Items
}
function FullText-Candidates($Scope, $Params) {
    $inlineBudget = if ($null -ne $Params.inline_body_budget) { FullText-Body-Budget $Params.inline_body_budget } else { $null }
    $inlineBodies = [Collections.Generic.List[object]]::new()
    $scanLimit = if ($null -eq $Params.scan_limit) { 50 } else { [int]$Params.scan_limit }
    $limit = if ($null -eq $Params.limit) { 50 } else { [int]$Params.limit }
    $offset = if ($null -eq $Params.offset) { 0 } else { [int]$Params.offset }
    if ($scanLimit -lt 1 -or $scanLimit -gt 50) { throw 'INVALID_ARGUMENT: fulltext_candidates scan_limit must be between 1 and 50' }
    if ($limit -lt 1 -or $limit -gt 50) { throw 'INVALID_ARGUMENT: fulltext_candidates limit must be between 1 and 50' }
    if ($offset -lt 0) { throw 'INVALID_ARGUMENT: offset must be non-negative' }
    $person = if ($null -eq $Params.person) { '' } else { [string]$Params.person }
    $personActive = -not [string]::IsNullOrEmpty($person)
    $personRole = if ($Params.person_role) { [string]$Params.person_role } else { 'any' }
    if (@('any', 'sender', 'to', 'cc') -notcontains $personRole) { throw 'INVALID_ARGUMENT: person_role must be one of any, sender, to, cc' }
    $items = Track ($Scope.folder.Items)
    $items = FullText-Filtered-Items $items $Params
    $items.Sort('[ReceivedTime]', $true)
    $total = [int]$items.Count
    $afterUtc = if ($Params.after) { ([DateTimeOffset]::Parse($Params.after)).UtcDateTime } else { $null }
    $beforeUtc = if ($Params.before) { ([DateTimeOffset]::Parse($Params.before)).UtcDateTime } else { $null }
    $position = [Math]::Min($offset, $total)
    $scanned = 0
    $matches = [Collections.Generic.List[object]]::new()
    $validationCandidates = [Collections.Generic.List[object]]::new()
    $personIncomplete = $false
    while ($position -lt $total -and $scanned -lt $scanLimit -and $matches.Count -lt $limit) {
        $position++
        $scanned++
        $mail = Track ($items.Item($position))
        if ($mail.Class -ne 43) { continue }
        # Capture the source guard before any date, person, or metadata reads.
        # A matching candidate must describe one stable item version.
        $accessor = FullText-Accessor $mail
        $beforeGuard = FullText-Mail-State $mail $Scope $accessor $true
        if ($afterUtc -or $beforeUtc) {
            $receivedUtc = ([DateTime](Mail-Value $mail $beforeGuard.metadata_values 'ReceivedTime')).ToUniversalTime()
            if ($afterUtc -and $receivedUtc -lt $afterUtc) { continue }
            if ($beforeUtc -and $receivedUtc -ge $beforeUtc) { continue }
        }
        if (-not $beforeGuard.in_scope) { continue }
        $personMatch = $null
        $match = $true
        if ($personActive) {
            $personMatch = Person-Match $mail $person $personRole (Sender-Info $mail $beforeGuard.metadata_values)
            $personIncomplete = $personIncomplete -or $personMatch.incomplete
            $match = $personMatch.matched
        }
        if ($match) {
            $senderInfo = if ($personMatch) { $personMatch.sender } else { $null }
            $metadata = Mail-Metadata $mail $Scope.store_id $senderInfo $null $beforeGuard.metadata_values ([bool]$Params.defer_metadata)
            if ($personActive -and $personMatch) { $metadata.matched_person_roles = @($personMatch.matched_roles) }
            # The negotiated search path validates the original guard after
            # matching. Keep that original version, never a later
            # version that could bless metadata read during a mutation.
            $afterGuard = $beforeGuard
            $stable = $beforeGuard.in_scope -and ($null -ne $beforeGuard.signature -or $null -ne $beforeGuard.ephemeral_guard)
            if (-not $Params.defer_metadata) {
                $afterGuard = FullText-Mail-State $mail $Scope $accessor
                $useSignature = $null -ne $beforeGuard.signature -and $null -ne $afterGuard.signature
                $stable = $afterGuard.in_scope -and (FullText-Guard-Matches $beforeGuard $afterGuard $useSignature)
            }
            if (-not $stable) {
                throw 'FULLTEXT_SOURCE_CHANGED: A message changed while its full-text candidate metadata was read; retry the search'
            }
            $metadata['_fulltext_signature'] = $afterGuard.signature
            $metadata['_fulltext_guard'] = if ($null -eq $afterGuard.signature) { $afterGuard.ephemeral_guard } else { $null }
            $metadata['_fulltext_offset'] = $position
            $candidate = @{
                id = $metadata.id
                signature = $afterGuard.signature
                guard = if ($null -eq $afterGuard.signature) { $afterGuard.ephemeral_guard } else { $null }
            }
            $validationCandidates.Add($candidate)
            $matches.Add($metadata)
            if ($null -ne $inlineBudget) {
                # The original guard spans metadata AND body. Reuse only this
                # operation's live objects; no COM reference survives the RPC.
                $prepared = @{ mail = $mail; accessor = $accessor; before = $beforeGuard;
                    defer_validation = [bool]$Params.direct_scan -and [bool]$Params.defer_metadata;
                    direct_query = if ($Params.direct_scan) { [string]$Params.direct_match_query } else { $null } }
                $inlineBodies.Add((FullText-Bounded-Read $candidate $Scope $inlineBudget $prepared))
            }
        }
    }
    $next = $null
    if ($position -lt $total) { $next = $position }
    if (-not (FullText-Scope-Unchanged $Scope)) {
        throw 'FULLTEXT_SOURCE_CHANGED: The requested folder moved out of its verified scope during the candidate scan; retry the search'
    }
    $result = @{
        messages = @($matches.ToArray())
        scanned = $scanned
        scan_limit_applied = $scanLimit
        filtered_item_count = $total
        next_offset = $next
        complete = ($null -eq $next)
        scope = 'single_folder'
        pagination = 'live_offset_not_snapshot'
        person_search_incomplete = if ($personActive) { $personIncomplete } else { $false }
        folder_id = $Scope.id
        fulltext_capabilities = @{ batch_read = $true; deferred_metadata = [bool]$Params.defer_metadata }
        validation = @{ candidates = @($validationCandidates.ToArray()) }
    }
    if ($null -ne $inlineBudget) { $result.inline_bodies = @($inlineBodies.ToArray()) }
    return $result
}
function FullText-Validate($Params) {
    $scope = Resolve-FullText-Scope $Params.folder
    $validation = $Params.validation
    if ($null -eq $validation -or $null -eq $validation.candidates) {
        throw 'INVALID_ARGUMENT: fulltext_validate validation candidates are required'
    }
    $candidates = @($validation.candidates)
    if ($candidates.Count -gt 50) {
        throw 'INVALID_ARGUMENT: fulltext_validate accepts at most 50 candidates'
    }
    $metadataIds = @{}
    foreach ($id in @($Params.metadata_ids)) { if ($null -ne $id) { $metadataIds[[string]$id] = $true } }
    $metadata = [Collections.Generic.List[object]]::new()
    foreach ($candidate in $candidates) {
        if ($null -eq $candidate -or [string]::IsNullOrWhiteSpace([string]$candidate.id)) {
            throw 'INVALID_ARGUMENT: fulltext_validate candidate id is required'
        }
        try { $mail = Get-Mail $candidate.id }
        catch { throw 'FULLTEXT_SOURCE_CHANGED: A message disappeared while its full-text index was being validated; retry the search' }
        if ($metadataIds.ContainsKey([string]$candidate.id)) {
            $attachments = Track ($mail.Attachments)
            $metadata.Add(@{ id = $candidate.id; attachment_count = [int]$attachments.Count })
        }
        # Check after all deferred reads against the pre-metadata candidate
        # version, so no intermediate guard can legitimize changed data.
        $current = FullText-Mail-State $mail $scope (FullText-Accessor $mail)
        if (-not $current.in_scope) {
            throw 'FULLTEXT_SOURCE_CHANGED: A message moved or was deleted while its full-text index was being validated; retry the search'
        }
        $expectedSignature = if ($null -ne $candidate.signature -and -not [string]::IsNullOrEmpty([string]$candidate.signature)) { [string]$candidate.signature } else { $null }
        $expectedGuard = if ($null -ne $candidate.guard -and -not [string]::IsNullOrEmpty([string]$candidate.guard)) { [string]$candidate.guard } else { $null }
        if ($null -eq $expectedSignature -and $null -eq $expectedGuard) {
            throw 'FULLTEXT_SOURCE_CHANGED: A full-text candidate has no usable version guard; retry the search'
        }
        $expected = @{ signature = $expectedSignature; ephemeral_guard = $expectedGuard }
        if (-not (FullText-Guard-Matches $expected $current ($null -ne $expectedSignature))) {
            throw 'FULLTEXT_SOURCE_CHANGED: A message changed while its full-text index was being validated; retry the search'
        }
    }
    if (-not (FullText-Scope-Unchanged $scope)) {
        throw 'FULLTEXT_SOURCE_CHANGED: The requested folder moved out of its verified scope during validation; retry the search'
    }
    if ($metadata.Count -ne $metadataIds.Count) { throw 'INVALID_ARGUMENT: Result metadata must belong to validated candidates' }
    return @{ valid = $true; metadata = @($metadata.ToArray()) }
}
function FullText-Read($Params, $BatchScope = $null, $Prepared = $null) {
    if ([string]::IsNullOrWhiteSpace([string]$Params.id)) { throw 'INVALID_ARGUMENT: id is required for full-text reads' }
    $scope = if ($null -ne $BatchScope) { $BatchScope } else { Resolve-FullText-Scope $Params.folder }
    if ($null -ne $Prepared) {
        $mail = $Prepared.mail
        $accessor = $Prepared.accessor
        $before = $Prepared.before
    } else {
        $mail = Get-Mail $Params.id
        $accessor = FullText-Accessor $mail
        $before = FullText-Mail-State $mail $scope $accessor
    }
    if (-not $before.in_scope) { return (FullText-Result 'changed' $null $null) }
    $expected = $null
    $useSignature = $false
    if ($null -ne $Params.signature -and -not [string]::IsNullOrEmpty([string]$Params.signature)) {
        $expected = [string]$Params.signature
        $useSignature = $true
        if ($null -eq $before.signature -or -not [string]::Equals([string]$before.signature, $expected, [StringComparison]::Ordinal)) {
            return (FullText-Result 'changed' $null $before.signature)
        }
    }
    if (-not $useSignature -and $null -eq $before.ephemeral_guard) {
        return (FullText-Result 'changed' $null $null)
    }
    if (-not $useSignature -and $null -ne $Params.guard -and -not [string]::IsNullOrEmpty([string]$Params.guard)) {
        if (-not [string]::Equals([string]$before.ephemeral_guard, [string]$Params.guard, [StringComparison]::Ordinal)) {
            return (FullText-Result 'changed' $null $null)
        }
    }
    # Outlook's olHeaderOnly DownloadState is 0; never treat its Body as available.
    if ($null -eq $before.download_state) {
        return (FullText-Result 'changed' $null $(if ($useSignature) { $before.signature } else { $null }))
    }
    if ([int]$before.download_state -eq 0) {
        return (FullText-Result 'missing' $null $(if ($useSignature) { $before.signature } else { $null }))
    }
    $rawBody = $null
    $read = $true
    try { $rawBody = $mail.Body } catch { $read = $false }
    if ($read -and $null -eq $rawBody) { $read = $false }
    $body = if ($read) { [string]$rawBody } else { $null }
    if ($null -ne $Prepared -and $Prepared.defer_validation) {
        # A direct scan cannot cache or publish this body yet. Its original
        # guard is checked by fulltext_validate after ALL reads and matching.
        # The distinct status keeps an ordinary index loader from accepting it.
        if (-not $read) { return (FullText-Result 'missing' $null $before.signature) }
        if ($body.Length -gt 16000000) { return (FullText-Result 'truncated' $null $before.signature) }
        return (FullText-Result 'unverified' $body $before.signature)
    }
    $after = FullText-Mail-State $mail $scope $accessor
    if (-not $after.in_scope -or -not (FullText-Guard-Matches $before $after $useSignature) -or ($null -eq $BatchScope -and -not (FullText-Scope-Unchanged $scope))) {
        return (FullText-Result 'changed' $null $(if ($useSignature) { $after.signature } else { $null }))
    }
    $signature = if ($useSignature) { $after.signature } else { $null }
    if (-not $read) { return (FullText-Result 'missing' $null $signature) }
    # Body is exposed as one .NET string, so cap returned UTF-16 code units and keep the status incomplete.
    if ($body.Length -gt 16000000) {
        return (FullText-Result 'truncated' $null $signature)
    }
    return (FullText-Result 'ready' $body $signature)
}
function FullText-Body-Budget($Params) {
    $maxBytes = [long]$Params.max_bytes
    $bodyBytes = [long]$Params.max_body_bytes
    $budgetMs = [int]$Params.budget_ms
    if ($maxBytes -lt 1 -or $maxBytes -gt 134217728 -or $bodyBytes -lt 1 -or $bodyBytes -gt 134217728 -or $budgetMs -lt 1 -or $budgetMs -gt 30000) {
        throw 'INVALID_ARGUMENT: Invalid full-text batch budget'
    }
    return @{ clock = [Diagnostics.Stopwatch]::StartNew(); max_bytes = $maxBytes; max_body_bytes = $bodyBytes;
        budget_ms = $budgetMs; used = 0L; exhausted = $false }
}
function FullText-Bounded-Read($Candidate, $Scope, $Budget, $Prepared = $null) {
    $body = FullText-Result 'pending' $null $Candidate.signature
    if (-not $Budget.exhausted -and $Budget.clock.ElapsedMilliseconds -lt $Budget.budget_ms) {
        try { $body = FullText-Read $Candidate $Scope $Prepared }
        catch { $body = FullText-Result 'changed' $null $Candidate.signature }
        if ($body.status -eq 'ready' -or $body.status -eq 'unverified') {
            $bytes = [Text.Encoding]::UTF8.GetByteCount($body.text)
            if ($bytes -gt $Budget.max_body_bytes) { $body = FullText-Result 'truncated' $null $Candidate.signature }
            elseif ($Budget.used + $bytes -gt $Budget.max_bytes) {
                $body = FullText-Result 'pending' $null $Candidate.signature
                $Budget.exhausted = $true
            } else {
                $Budget.used += $bytes
                # An ASCII query can only match ASCII body characters after
                # JS lowercase within this conservative alphabet. The allowed
                # non-ASCII ranges never lowercase into ASCII (exhaustively
                # checked against the JS runtime in tests). In particular,
                # dotted I and Kelvin sign are excluded. A hit or any other
                # alphabet still goes through the canonical JS path.
                # Count the original body bytes before omitting any payload.
                if ($body.status -eq 'unverified' -and $null -ne $Prepared -and
                    -not [string]::IsNullOrEmpty($Prepared.direct_query) -and
                    $Prepared.direct_query -cnotmatch '[^\x00-\x7f]' -and
                    $body.text -cnotmatch '[^\x00-\x7f\u00a0-\u00ff\u2000-\u206f\u2e80-\u9fff\ufeff\uff00-\uffef]' -and
                    $body.text.IndexOf($Prepared.direct_query, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
                    $body = @{ status = 'unverified_no_match'; text = $null; signature = $body.signature;
                        query = $Prepared.direct_query; body_bytes = $bytes; body_code_units = $body.text.Length }
                }
            }
        }
    }
    $body.id = $Candidate.id
    return $body
}
function FullText-Read-Batch($Params) {
    $candidates = @($Params.candidates)
    if ($candidates.Count -lt 1 -or $candidates.Count -gt 50) { throw 'INVALID_ARGUMENT: fulltext_read_batch accepts 1 to 50 candidates' }
    $budget = FullText-Body-Budget $Params
    $scope = Resolve-FullText-Scope $Params.folder
    $results = [Collections.Generic.List[object]]::new()
    foreach ($candidate in $candidates) {
        $results.Add((FullText-Bounded-Read $candidate $scope $budget))
    }
    # No body from this batch may escape if the shared folder proof expired.
    if (-not (FullText-Scope-Unchanged $scope)) {
        throw 'FULLTEXT_SOURCE_CHANGED: The requested folder moved during a body batch; retry the search'
    }
    return @{ bodies = @($results.ToArray()) }
}
function Iso-Date($Value) {
    if ($null -eq $Value) { return $null }
    return ([DateTime]$Value).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
}
function Mail-Value($Mail, $Values, $Name) {
    if ($null -ne $Values -and $Values.ContainsKey($Name)) { return ,$Values[$Name] }
    return ,$Mail.$Name
}
function Sender-Info($Mail, $Values = $null) {
    $address = [string](Mail-Value $Mail $Values 'SenderEmailAddress')
    $addressType = [string](Mail-Value $Mail $Values 'SenderEmailType')
    $smtp = $null
    if ($addressType -eq 'SMTP') { $smtp = $address }
    if ($addressType -eq 'EX') {
        try {
            $sender = Track ($Mail.Sender)
            $senderType = $sender.AddressEntryUserType
            if ($senderType -eq 0 -or $senderType -eq 5) {
                $exchangeUser = Track ($sender.GetExchangeUser())
                if ($exchangeUser) { $smtp = $exchangeUser.PrimarySmtpAddress }
            } elseif ($senderType -eq 1) {
                $distribution = Track ($sender.GetExchangeDistributionList())
                if ($distribution) { $smtp = $distribution.PrimarySmtpAddress }
            }
            if (-not $smtp) {
                $accessor = Track ($sender.PropertyAccessor)
                $smtp = $accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E')
            }
        } catch { }
    }
    return @{ name = [string](Mail-Value $Mail $Values 'SenderName'); address = $address; address_type = $addressType; smtp_address = $smtp }
}
function Mail-Metadata($Mail, $StoreId, $SenderInfo = $null, $Attachments = $null, $Values = $null, $DeferAttachments = $false) {
    if (-not $DeferAttachments -and $null -eq $Attachments) { $Attachments = Track ($Mail.Attachments) }
    if ($null -eq $SenderInfo) { $SenderInfo = Sender-Info $Mail $Values }
    return @{
        id = Make-Id 'message' (Mail-Value $Mail $Values 'EntryID') $StoreId
        subject = [string](Mail-Value $Mail $Values 'Subject')
        sender = $SenderInfo
        received_at = Iso-Date (Mail-Value $Mail $Values 'ReceivedTime')
        unread = [bool](Mail-Value $Mail $Values 'UnRead')
        size_bytes = [int](Mail-Value $Mail $Values 'Size')
        attachment_count = if ($DeferAttachments) { $null } else { [int]$attachments.Count }
        download_state = [int](Mail-Value $Mail $Values 'DownloadState')
    }
}
function Attachment-List($Mail, $Attachments = $null) {
    if ($null -eq $Attachments) { $Attachments = Track ($Mail.Attachments) }
    return @(for ($i = 1; $i -le $attachments.Count; $i++) {
        $a = Track ($attachments.Item($i))
        @{ index = $i; filename = [string]$a.FileName; size_bytes = [int]$a.Size; type = [int]$a.Type }
    })
}
function Folder-Info($Folder) {
    $items = Track ($Folder.Items)
    $children = Track ($Folder.Folders)
    return @{
        id = Make-Id 'folder' $Folder.EntryID $Folder.StoreID
        name = [string]$Folder.Name
        path = [string]$Folder.FolderPath
        item_count = [int]$items.Count
        unread_count = [int]$Folder.UnReadItemCount
        child_count = [int]$children.Count
    }
}
function Contains-Text($Value, $Query) {
    return (-not [string]::IsNullOrEmpty($Query) -and ([string]$Value).IndexOf($Query, [StringComparison]::OrdinalIgnoreCase) -ge 0)
}
function Recipient-Info($Recipient) {
    $type = $null
    try { $type = [int]$Recipient.Type } catch { return $null }
    if ($type -ne 1 -and $type -ne 2) { return $null }
    $name = $null
    $address = $null
    $addressType = $null
    $smtp = $null
    try { $name = [string]$Recipient.Name } catch { }
    try { $address = [string]$Recipient.Address } catch { }
    $entry = $null
    try {
        $entry = Track ($Recipient.AddressEntry)
        if ($entry) {
            try { $addressType = [string]$entry.Type } catch { }
            if ($addressType -eq 'SMTP' -and $address) { $smtp = $address }
            if (-not $smtp) {
                try {
                    if ($entry.AddressEntryUserType -eq 0 -or $entry.AddressEntryUserType -eq 5) {
                        $exchangeUser = Track ($entry.GetExchangeUser())
                        if ($exchangeUser) { $smtp = $exchangeUser.PrimarySmtpAddress }
                    } elseif ($entry.AddressEntryUserType -eq 1) {
                        $distribution = Track ($entry.GetExchangeDistributionList())
                        if ($distribution) { $smtp = $distribution.PrimarySmtpAddress }
                    }
                } catch { }
            }
            if (-not $smtp) {
                try {
                    $accessor = Track ($entry.PropertyAccessor)
                    $smtp = $accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E')
                } catch { }
            }
        }
    } catch { }
    if ([string]::IsNullOrEmpty($smtp) -and $address -and $address -match '@') { $smtp = $address }
    return @{
        role = if ($type -eq 1) { 'to' } else { 'cc' }
        name = if ([string]::IsNullOrEmpty($name)) { $null } else { $name }
        address = if ([string]::IsNullOrEmpty($address)) { $null } else { $address }
        address_type = if ([string]::IsNullOrEmpty($addressType)) { $null } else { $addressType }
        smtp_address = if ([string]::IsNullOrEmpty($smtp)) { $null } else { [string]$smtp }
    }
}
function Person-Match($Mail, $Person, $Role, $SenderInfo = $null) {
    $matchedRoles = [Collections.Generic.List[string]]::new()
    $incomplete = $false
    $sender = $null
    if ($Role -eq 'any' -or $Role -eq 'sender') {
        if ($null -eq $SenderInfo) {
            try { $SenderInfo = Sender-Info $Mail } catch { $SenderInfo = $null }
        }
        $sender = $SenderInfo
        $senderValues = @($sender.name, $sender.address, $sender.smtp_address)
        $senderNameAvailable = -not [string]::IsNullOrEmpty($sender.name)
        $senderSmtpAvailable = -not [string]::IsNullOrEmpty($sender.smtp_address)
        if (-not $senderNameAvailable -or -not $senderSmtpAvailable) { $incomplete = $true }
        if ($senderValues | Where-Object { Contains-Text $_ $Person }) { $matchedRoles.Add('sender') }
    }
    if ($Role -eq 'any' -or $Role -eq 'to' -or $Role -eq 'cc') {
        $recipients = $null
        try { $recipients = Track ($Mail.Recipients) } catch { $incomplete = $true }
        if ($null -eq $recipients) {
            $incomplete = $true
        } else {
            try {
                for ($i = 1; $i -le $recipients.Count; $i++) {
                    $recipient = Track ($recipients.Item($i))
                    $info = Recipient-Info $recipient
                    if ($null -eq $info) {
                        try {
                            $type = [int]$recipient.Type
                            if ($type -eq 1 -or $type -eq 2) { $incomplete = $true }
                        } catch { $incomplete = $true }
                        continue
                    }
                    if ($Role -ne 'any' -and $info.role -ne $Role) { continue }
                    $values = @($info.name, $info.address, $info.smtp_address)
                    $nameAvailable = -not [string]::IsNullOrEmpty($info.name)
                    $smtpAvailable = -not [string]::IsNullOrEmpty($info.smtp_address)
                    if (-not $nameAvailable -or -not $smtpAvailable) { $incomplete = $true }
                    if (($values | Where-Object { Contains-Text $_ $Person }) -and -not $matchedRoles.Contains($info.role)) {
                        $matchedRoles.Add($info.role)
                    }
                }
            } catch { $incomplete = $true }
        }
    }
    $orderedRoles = @('sender', 'to', 'cc') | Where-Object { $matchedRoles.Contains($_) }
    return @{ matched = ($matchedRoles.Count -gt 0); matched_roles = @($orderedRoles); incomplete = $incomplete; sender = $sender }
}
function Run-Operation($Method, $Params) {
    try { $app = Track ([Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')) }
    catch { throw 'OUTLOOK_UNAVAILABLE: Open and sign in to classic Outlook in the same Windows user session as the service' }
    $script:session = Track ($app.GetNamespace('MAPI'))
    switch ($Method) {
        'health' {
            $stores = Track ($script:session.Stores)
            return @{ outlook_version = [string]$app.Version; connected = $true; offline = [bool]$script:session.Offline; store_count = [int]$stores.Count; exchange_connection_mode = [int]$script:session.ExchangeConnectionMode }
        }
        'stores' {
            $stores = Track ($script:session.Stores)
            $list = @(for ($i = 1; $i -le $stores.Count; $i++) {
                $s = Track ($stores.Item($i))
                $root = Track ($s.GetRootFolder())
                @{ name = [string]$s.DisplayName; root_id = Make-Id 'folder' $root.EntryID $s.StoreID; exchange_store_type = [int]$s.ExchangeStoreType; is_data_file = [bool]$s.IsDataFileStore }
            })
            return @{ stores = $list }
        }
        'folders' {
            $parent = Get-Folder $Params.folder
            $children = Track ($parent.Folders)
            $end = [Math]::Min($children.Count, $Params.offset + $Params.limit)
            $list = @(for ($i = $Params.offset + 1; $i -le $end; $i++) { Folder-Info (Track ($children.Item($i))) })
            $next = $null
            if ($end -lt $children.Count) { $next = $end }
            return @{ parent = Folder-Info $parent; folders = $list; total = [int]$children.Count; next_offset = $next }
        }
        'fulltext_candidates' {
            return (FullText-Candidates (Resolve-FullText-Scope $Params.folder) $Params)
        }
        'fulltext_validate' {
            return (FullText-Validate $Params)
        }
        'fulltext_read' {
            return (FullText-Read $Params)
        }
        'fulltext_read_batch' {
            return (FullText-Read-Batch $Params)
        }
        { $_ -eq 'messages' -or $_ -eq 'search' } {
            $folder = Get-Folder $Params.folder
            $items = Track ($folder.Items)
            $clauses = [Collections.Generic.List[string]]::new()
            if ($null -ne $Params.unread) {
                if ($Params.unread) { $clauses.Add('[UnRead] = True') } else { $clauses.Add('[UnRead] = False') }
            }
            foreach ($spec in @(@('after','>='), @('before','<'))) {
                $value = $Params.($spec[0])
                if ($value) {
                    $local = [DateTimeOffset]::Parse($value).LocalDateTime
                    # Jet uses the Windows locale and minute precision. Widen the
                    # coarse filter, then enforce the exact ISO bounds below.
                    $rounded = $local.AddTicks(-($local.Ticks % [TimeSpan]::TicksPerMinute))
                    if ($spec[0] -eq 'before' -and $rounded -lt $local) { $rounded = $rounded.AddMinutes(1) }
                    $date = $rounded.ToString('g', [Globalization.CultureInfo]::CurrentCulture)
                    $clauses.Add("[ReceivedTime] $($spec[1]) '$date'")
                }
            }
            if ($clauses.Count) { $items = Track ($items.Restrict(($clauses -join ' AND '))) }
            $items.Sort('[ReceivedTime]', $true)
            $total = [int]$items.Count
            $afterUtc = if ($Params.after) { ([DateTimeOffset]::Parse($Params.after)).UtcDateTime } else { $null }
            $beforeUtc = if ($Params.before) { ([DateTimeOffset]::Parse($Params.before)).UtcDateTime } else { $null }
            $position = [Math]::Min([int]$Params.offset, $total)
            $scanned = 0
            $matches = [Collections.Generic.List[object]]::new()
            $query = if ($null -eq $Params.query) { '' } else { [string]$Params.query }
            $person = if ($null -eq $Params.person) { '' } else { [string]$Params.person }
            $personActive = -not [string]::IsNullOrEmpty($person)
            $personRole = if ($Params.person_role) { [string]$Params.person_role } else { 'any' }
            if ($Method -eq 'search' -and @('any', 'sender', 'to', 'cc') -notcontains $personRole) { throw 'INVALID_ARGUMENT: person_role must be one of any, sender, to, cc' }
            $personIncomplete = $false
            while ($position -lt $total -and $scanned -lt $Params.scan_limit -and $matches.Count -lt $Params.limit) {
                $position++
                $scanned++
                $mail = Track ($items.Item($position))
                if ($mail.Class -ne 43) { continue }
                if ($afterUtc -or $beforeUtc) {
                    $receivedUtc = ([DateTime]$mail.ReceivedTime).ToUniversalTime()
                    if ($afterUtc -and $receivedUtc -lt $afterUtc) { continue }
                    if ($beforeUtc -and $receivedUtc -ge $beforeUtc) { continue }
                }
                $match = $true
                $senderInfo = $null
                $personMatch = $null
                if ($Method -eq 'search') {
                    $match = [string]::IsNullOrEmpty($query) -or (Contains-Text $mail.Subject $query) -or (Contains-Text $mail.SenderName $query) -or (Contains-Text $mail.SenderEmailAddress $query)
                    if (-not $match -and $query) {
                        $senderInfo = Sender-Info $mail
                        $match = Contains-Text $senderInfo.smtp_address $query
                    }
                    if (-not $match -and $Params.search_body) { $match = Contains-Text $mail.Body $query }
                    if ($match -and $personActive) {
                        $personMatch = Person-Match $mail $person $personRole $senderInfo
                        $personIncomplete = $personIncomplete -or $personMatch.incomplete
                        $match = $personMatch.matched
                    }
                }
                if ($match) {
                    if ($personMatch -and $personMatch.sender) { $senderInfo = $personMatch.sender }
                    $metadata = Mail-Metadata $mail $folder.StoreID $senderInfo
                    if ($personActive -and $personMatch) { $metadata.matched_person_roles = @($personMatch.matched_roles) }
                    $matches.Add($metadata)
                }
            }
            $next = $null
            if ($position -lt $total) { $next = $position }
            return @{ messages = @($matches.ToArray()); scanned = $scanned; filtered_item_count = $total; next_offset = $next; complete = ($null -eq $next); scope = 'single_folder'; pagination = 'live_offset_not_snapshot'; person_search_incomplete = if ($personActive) { $personIncomplete } else { $false } }
        }
        'read' {
            $mail = Get-Mail $Params.id
            $decoded = Parse-Id $Params.id 'message'
            $attachments = Track ($mail.Attachments)
            $result = Mail-Metadata $mail $decoded.store $null $attachments
            if ($Params.format -eq 'html') { $body = [string]$mail.HTMLBody } else { $body = [string]$mail.Body }
            $offset = [Math]::Min([int]$Params.body_offset, $body.Length)
            $length = [Math]::Min([int]$Params.max_chars, $body.Length - $offset)
            $next = $null
            if ($offset + $length -lt $body.Length) { $next = $offset + $length }
            $result.body = @{ format = $Params.format; text = $body.Substring($offset, $length); total_chars = $body.Length; offset = $offset; next_offset = $next; truncated = ($null -ne $next) }
            $result.to = [string]$mail.To
            $result.cc = [string]$mail.CC
            $result.attachments = @(Attachment-List $mail $attachments)
            return $result
        }
        'attachments' { return @{ attachments = @(Attachment-List (Get-Mail $Params.id)) } }
        'read_attachment' {
            $mail = Get-Mail $Params.id
            $attachments = Track ($mail.Attachments)
            if ($Params.index -gt $attachments.Count) { throw 'INVALID_ARGUMENT: Attachment index is out of range' }
            $attachment = Track ($attachments.Item([int]$Params.index))
            if ($attachment.Type -ne 1) { throw 'UNSUPPORTED_ATTACHMENT: Only by-value file attachments are supported' }
            if ($attachment.Size -gt $Params.max_bytes) { throw 'ATTACHMENT_TOO_LARGE: Attachment exceeds max_bytes' }
            $accessor = Track ($attachment.PropertyAccessor)
            [byte[]]$bytes = $accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x37010102')
            if ($bytes.Length -gt $Params.max_bytes) { throw 'ATTACHMENT_TOO_LARGE: Attachment exceeds max_bytes' }
            $sha = [Security.Cryptography.SHA256]::Create()
            try { $hash = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-','').ToLowerInvariant() } finally { $sha.Dispose() }
            $offset = [Math]::Min([int]$Params.byte_offset, $bytes.Length)
            $count = [Math]::Min([int]$Params.chunk_bytes, $bytes.Length - $offset)
            $next = $null
            if ($offset + $count -lt $bytes.Length) { $next = $offset + $count }
            return @{ filename = [string]$attachment.FileName; size_bytes = $bytes.Length; sha256 = $hash; encoding = 'base64'; offset = $offset; chunk_size = $count; next_offset = $next; data = [Convert]::ToBase64String($bytes, $offset, $count) }
        }
        default { throw 'INVALID_ARGUMENT: Unknown operation' }
    }
}

while ($null -ne ($line = [Console]::ReadLine())) {
    $script:refs = [Collections.Generic.List[object]]::new()
    $request = $null
    try {
        $request = $line | ConvertFrom-Json
        $result = Run-Operation $request.method $request.params
        $response = @{ id = $request.id; result = $result }
    } catch {
        $code = 'OUTLOOK_ERROR'
        $message = 'Outlook could not read the requested item. It may have moved, been deleted, or require an Outlook security prompt.'
        if ($_.Exception.Message -match '^([A-Z_]+): (.+)$') { $code = $Matches[1]; $message = $Matches[2] }
        $response = @{ id = $request.id; error = @{ code = $code; message = $message } }
    } finally {
        for ($i = $script:refs.Count - 1; $i -ge 0; $i--) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:refs[$i]) } catch { }
        }
        $script:refs.Clear()
        $script:session = $null
    }
    [Console]::WriteLine(($response | ConvertTo-Json -Depth 12 -Compress))
    $response = $null
    $result = $null
}
