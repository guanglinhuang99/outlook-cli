# Persistent JSON-lines worker. Runs only fixed read operations on a single STA thread.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::InputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

function Track($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { $script:refs.Add($Value) }
    return ,$Value
}
function Make-Id($Kind, $Entry, $Store) {
    $json = @{ v = 1; kind = $Kind; entry = [string]$Entry; store = [string]$Store } | ConvertTo-Json -Compress
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json)).TrimEnd('=').Replace('+','-').Replace('/','_')
}
function Parse-Id($Id, $Kind) {
    try {
        $b64 = $Id.Replace('-','+').Replace('_','/')
        $b64 = $b64.PadRight($b64.Length + ((4 - $b64.Length % 4) % 4), '=')
        $obj = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b64)) | ConvertFrom-Json
        if ($obj.v -ne 1 -or $obj.kind -ne $Kind -or $obj.entry -notmatch '^[a-fA-F0-9]+$' -or $obj.store -notmatch '^[a-fA-F0-9]+$') { throw 'bad id' }
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
function Same-Folder($Left, $Right) {
    $leftKey = Folder-Key $Left
    $rightKey = Folder-Key $Right
    if ($null -eq $leftKey -or $null -eq $rightKey) { return $false }
    return [string]::Equals($leftKey, $rightKey, [StringComparison]::OrdinalIgnoreCase)
}
function Parent-Folder($Folder) {
    try {
        $parent = Track ($Folder.Parent)
        if ($null -eq $parent -or $null -eq (Folder-Key $parent)) { return $null }
        return ,$parent
    } catch { return $null }
}
function Folder-Is-Within($Folder, $Root) {
    $seen = @{}
    $current = $Folder
    for ($depth = 0; $depth -lt 128 -and $null -ne $current; $depth++) {
        $key = Folder-Key $current
        if ($null -eq $key) { return $false }
        if (Same-Folder $current $Root) { return $true }
        if ($seen.ContainsKey($key)) { return $false }
        $seen[$key] = $true
        $current = Parent-Folder $current
    }
    return $false
}
function Resolve-FullText-Scope($FolderId) {
    if ([string]::IsNullOrWhiteSpace([string]$FolderId)) { throw 'INVALID_ARGUMENT: folder is required for full-text operations' }
    try { $folder = Get-Folder $FolderId } catch { throw 'INVALID_ID: Folder not found' }
    if ($null -eq $folder) { throw 'INVALID_ID: Folder not found' }
    try { $store = Track ($folder.Store) } catch { $store = $null }
    if ($null -eq $store) { throw 'FULLTEXT_FOLDER_SCOPE: Full-text operations require a known Outlook store' }
    $inbox = $null
    $sent = $null
    $deleted = $null
    try { $inbox = Track ($store.GetDefaultFolder(6)) } catch { }
    try { $sent = Track ($store.GetDefaultFolder(5)) } catch { }
    try { $deleted = Track ($store.GetDefaultFolder(3)) } catch { }
    # Deleted is a safety boundary. A missing Sent root is allowed when the
    # requested folder is in the same store's Inbox subtree.
    if ($null -eq $deleted) { throw 'FULLTEXT_FOLDER_SCOPE: Full-text operations require a known Deleted Items root' }
    if (Folder-Is-Within $folder $deleted) { throw 'FULLTEXT_FOLDER_SCOPE: Full-text operations exclude Deleted Items and its descendants' }
    $root = $null
    if ($null -ne $inbox -and (Folder-Is-Within $folder $inbox)) { $root = $inbox }
    elseif ($null -ne $sent -and (Folder-Is-Within $folder $sent)) { $root = $sent }
    if ($null -eq $root) { throw 'FULLTEXT_FOLDER_SCOPE: Full-text operations are limited to the default Inbox and Sent subtrees' }
    return ,@{
        folder = $folder
        id = Make-Id 'folder' $folder.EntryID $folder.StoreID
        root = $root
        root_id = Make-Id 'folder' $root.EntryID $root.StoreID
        deleted = $deleted
    }
}
function Mail-Current-Folder($Mail) {
    try {
        $folder = Track ($Mail.Parent)
        if ($null -eq $folder -or $null -eq (Folder-Key $folder)) { return $null }
        return ,$folder
    } catch { return $null }
}
function Mail-In-FullText-Scope($Mail, $Scope, $RequestedFolder) {
    $current = Mail-Current-Folder $Mail
    if ($null -eq $current) { return $false }
    if (-not (Same-Folder $current $RequestedFolder)) { return $false }
    if (Folder-Is-Within $current $Scope.deleted) { return $false }
    if (-not (Folder-Is-Within $current $Scope.root)) { return $false }
    return $true
}
function Change-Key-Hex($Mail) {
    try {
        $accessor = Track ($Mail.PropertyAccessor)
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
function Mail-Guard($Mail) {
    $changeKey = Change-Key-Hex $Mail
    $downloadState = $null
    $lastModification = $null
    $size = $null
    $downloadStateValue = $null
    $lastModificationValue = $null
    $sizeValue = $null
    try { $downloadStateValue = $Mail.DownloadState } catch { }
    try { $lastModificationValue = $Mail.LastModificationTime } catch { }
    try { $sizeValue = $Mail.Size } catch { }
    if ($null -ne $downloadStateValue) { try { $downloadState = [int]$downloadStateValue } catch { } }
    if ($null -ne $lastModificationValue) { try { $lastModification = Iso-Date $lastModificationValue } catch { } }
    if ($null -ne $sizeValue) { try { $size = [string]$sizeValue } catch { } }
    $stateText = if ($null -eq $downloadState) { '<missing>' } else { [string]$downloadState }
    $lastText = if ($null -eq $lastModification) { '<missing>' } else { [string]$lastModification }
    $sizeText = if ($null -eq $size) { '<missing>' } else { [string]$size }
    $signature = $null
    if (-not [string]::IsNullOrEmpty($changeKey) -and $null -ne $downloadState) { $signature = $changeKey + '|' + $stateText }
    $ephemeralGuard = $null
    if ($null -ne $lastModification -and $null -ne $size -and $null -ne $downloadState) {
        $ephemeralGuard = $lastText + '|' + $sizeText + '|' + $stateText
    }
    return @{
        signature = $signature
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
        $beforeGuard = Mail-Guard $mail
        if ($afterUtc -or $beforeUtc) {
            $receivedUtc = ([DateTime]$mail.ReceivedTime).ToUniversalTime()
            if ($afterUtc -and $receivedUtc -lt $afterUtc) { continue }
            if ($beforeUtc -and $receivedUtc -ge $beforeUtc) { continue }
        }
        if (-not (Mail-In-FullText-Scope $mail $Scope $Scope.folder)) { continue }
        $personMatch = $null
        $match = $true
        if ($personActive) {
            $personMatch = Person-Match $mail $person $personRole
            $personIncomplete = $personIncomplete -or $personMatch.incomplete
            $match = $personMatch.matched
        }
        if ($match) {
            $senderInfo = if ($personMatch) { $personMatch.sender } else { $null }
            $metadata = Mail-Metadata $mail $Scope.folder.StoreID $senderInfo
            if ($personActive -and $personMatch) { $metadata.matched_person_roles = @($personMatch.matched_roles) }
            $afterGuard = Mail-Guard $mail
            $useSignature = $null -ne $beforeGuard.signature -and $null -ne $afterGuard.signature
            if (-not (Mail-In-FullText-Scope $mail $Scope $Scope.folder) -or -not (FullText-Guard-Matches $beforeGuard $afterGuard $useSignature)) {
                throw 'FULLTEXT_SOURCE_CHANGED: A message changed while its full-text candidate metadata was read; retry the search'
            }
            $metadata['_fulltext_signature'] = $afterGuard.signature
            $metadata['_fulltext_guard'] = if ($null -eq $afterGuard.signature) { $afterGuard.ephemeral_guard } else { $null }
            $metadata['_fulltext_offset'] = $position
            $validationCandidates.Add(@{
                id = $metadata.id
                signature = $afterGuard.signature
                guard = if ($null -eq $afterGuard.signature) { $afterGuard.ephemeral_guard } else { $null }
            })
            $matches.Add($metadata)
        }
    }
    $next = $null
    if ($position -lt $total) { $next = $position }
    return @{
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
        validation = @{ candidates = @($validationCandidates.ToArray()) }
    }
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
    foreach ($candidate in $candidates) {
        if ($null -eq $candidate -or [string]::IsNullOrWhiteSpace([string]$candidate.id)) {
            throw 'INVALID_ARGUMENT: fulltext_validate candidate id is required'
        }
        try { $mail = Get-Mail $candidate.id }
        catch { throw 'FULLTEXT_SOURCE_CHANGED: A message disappeared while its full-text index was being validated; retry the search' }
        if (-not (Mail-In-FullText-Scope $mail $scope $scope.folder)) {
            throw 'FULLTEXT_SOURCE_CHANGED: A message moved or was deleted while its full-text index was being validated; retry the search'
        }
        $current = Mail-Guard $mail
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
    return @{ valid = $true }
}
function FullText-Read($Params) {
    if ([string]::IsNullOrWhiteSpace([string]$Params.id)) { throw 'INVALID_ARGUMENT: id is required for full-text reads' }
    $scope = Resolve-FullText-Scope $Params.folder
    $mail = Get-Mail $Params.id
    $requestedFolder = $scope.folder
    $current = Mail-Current-Folder $mail
    if ($null -eq $current -or -not (Mail-In-FullText-Scope $mail $scope $requestedFolder)) { return (FullText-Result 'changed' $null $null) }
    $before = Mail-Guard $mail
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
    $afterFolder = Mail-Current-Folder $mail
    $after = Mail-Guard $mail
    if ($null -eq $afterFolder -or -not (Mail-In-FullText-Scope $mail $scope $requestedFolder) -or -not (FullText-Guard-Matches $before $after $useSignature)) {
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
function Iso-Date($Value) {
    if ($null -eq $Value) { return $null }
    return ([DateTime]$Value).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
}
function Sender-Info($Mail) {
    $address = [string]$Mail.SenderEmailAddress
    $addressType = [string]$Mail.SenderEmailType
    $smtp = $null
    if ($addressType -eq 'SMTP') { $smtp = $address }
    if ($addressType -eq 'EX') {
        try {
            $sender = Track ($Mail.Sender)
            if ($sender.AddressEntryUserType -eq 0 -or $sender.AddressEntryUserType -eq 5) {
                $exchangeUser = Track ($sender.GetExchangeUser())
                if ($exchangeUser) { $smtp = $exchangeUser.PrimarySmtpAddress }
            } elseif ($sender.AddressEntryUserType -eq 1) {
                $distribution = Track ($sender.GetExchangeDistributionList())
                if ($distribution) { $smtp = $distribution.PrimarySmtpAddress }
            }
            if (-not $smtp) {
                $accessor = Track ($sender.PropertyAccessor)
                $smtp = $accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E')
            }
        } catch { }
    }
    return @{ name = [string]$Mail.SenderName; address = $address; address_type = $addressType; smtp_address = $smtp }
}
function Mail-Metadata($Mail, $StoreId, $SenderInfo = $null, $Attachments = $null) {
    if ($null -eq $Attachments) { $Attachments = Track ($Mail.Attachments) }
    if ($null -eq $SenderInfo) { $SenderInfo = Sender-Info $Mail }
    return @{
        id = Make-Id 'message' $Mail.EntryID $StoreId
        subject = [string]$Mail.Subject
        sender = $SenderInfo
        received_at = Iso-Date $Mail.ReceivedTime
        unread = [bool]$Mail.UnRead
        size_bytes = [int]$Mail.Size
        attachment_count = [int]$attachments.Count
        download_state = [int]$Mail.DownloadState
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
