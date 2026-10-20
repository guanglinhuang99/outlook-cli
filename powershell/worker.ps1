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
function Mail-Metadata($Mail, $StoreId) {
    $attachments = Track ($Mail.Attachments)
    return @{
        id = Make-Id 'message' $Mail.EntryID $StoreId
        subject = [string]$Mail.Subject
        sender = Sender-Info $Mail
        received_at = Iso-Date $Mail.ReceivedTime
        unread = [bool]$Mail.UnRead
        size_bytes = [int]$Mail.Size
        attachment_count = [int]$attachments.Count
        download_state = [int]$Mail.DownloadState
    }
}
function Attachment-List($Mail) {
    $attachments = Track ($Mail.Attachments)
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
    return ([string]$Value).IndexOf($Query, [StringComparison]::OrdinalIgnoreCase) -ge 0
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
            $position = [Math]::Min([int]$Params.offset, $total)
            $scanned = 0
            $matches = [Collections.Generic.List[object]]::new()
            while ($position -lt $total -and $scanned -lt $Params.scan_limit -and $matches.Count -lt $Params.limit) {
                $position++
                $scanned++
                $mail = Track ($items.Item($position))
                if ($mail.Class -ne 43) { continue }
                $receivedUtc = ([DateTime]$mail.ReceivedTime).ToUniversalTime()
                if ($Params.after -and $receivedUtc -lt ([DateTimeOffset]::Parse($Params.after)).UtcDateTime) { continue }
                if ($Params.before -and $receivedUtc -ge ([DateTimeOffset]::Parse($Params.before)).UtcDateTime) { continue }
                $metadata = Mail-Metadata $mail $folder.StoreID
                $match = $true
                if ($Method -eq 'search') {
                    $match = (Contains-Text $metadata.subject $Params.query) -or (Contains-Text $metadata.sender.name $Params.query) -or (Contains-Text $metadata.sender.address $Params.query) -or (Contains-Text $metadata.sender.smtp_address $Params.query)
                    if (-not $match -and $Params.search_body) { $match = Contains-Text $mail.Body $Params.query }
                }
                if ($match) { $matches.Add($metadata) }
            }
            $next = $null
            if ($position -lt $total) { $next = $position }
            return @{ messages = @($matches.ToArray()); scanned = $scanned; filtered_item_count = $total; next_offset = $next; complete = ($null -eq $next); scope = 'single_folder'; pagination = 'live_offset_not_snapshot' }
        }
        'read' {
            $mail = Get-Mail $Params.id
            $decoded = Parse-Id $Params.id 'message'
            $result = Mail-Metadata $mail $decoded.store
            if ($Params.format -eq 'html') { $body = [string]$mail.HTMLBody } else { $body = [string]$mail.Body }
            $offset = [Math]::Min([int]$Params.body_offset, $body.Length)
            $length = [Math]::Min([int]$Params.max_chars, $body.Length - $offset)
            $next = $null
            if ($offset + $length -lt $body.Length) { $next = $offset + $length }
            $result.body = @{ format = $Params.format; text = $body.Substring($offset, $length); total_chars = $body.Length; offset = $offset; next_offset = $next; truncated = ($null -ne $next) }
            $result.to = [string]$mail.To
            $result.cc = [string]$mail.CC
            $result.attachments = @(Attachment-List $mail)
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
