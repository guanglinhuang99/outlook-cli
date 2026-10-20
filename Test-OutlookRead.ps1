# Read-only feasibility probe for an already-running classic Outlook session.
# Reports counts and field lengths; does not output mail text or addresses.
[CmdletBinding()]
param([ValidateRange(1, 10)][int]$SampleSize = 3)

$ErrorActionPreference = 'Stop'
$objects = [System.Collections.Generic.List[object]]::new()
function Track-Com($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        $objects.Add($Value)
    }
    return ,$Value
}

try {
    $app = Track-Com ([Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application'))
    $session = Track-Com ($app.GetNamespace('MAPI'))
    $stores = Track-Com ($session.Stores)
    $storeResults = @(for ($i = 1; $i -le $stores.Count; $i++) {
        $store = Track-Com ($stores.Item($i))
        $root = Track-Com ($store.GetRootFolder())
        $folders = Track-Com ($root.Folders)
        [pscustomobject]@{
            Index = $i
            ExchangeStoreType = [int]$store.ExchangeStoreType
            IsDataFileStore = $store.IsDataFileStore
            FileExtension = [IO.Path]::GetExtension($store.FilePath)
            TopLevelFolders = $folders.Count
        }
    })
    $inbox = Track-Com ($session.GetDefaultFolder(6))
    $items = Track-Com ($inbox.Items)
    $items.Sort('[ReceivedTime]', $true)
    $unreadBefore = $inbox.UnReadItemCount
    $samples = @()
    $limit = [Math]::Min($items.Count, 30)
    for ($i = 1; $i -le $limit -and $samples.Count -lt $SampleSize; $i++) {
        $mail = Track-Com ($items.Item($i))
        if ($mail.Class -ne 43) { continue }
        $wasUnread = $mail.UnRead
        $attachments = Track-Com ($mail.Attachments)
        $attachmentSizes = @(for ($a = 1; $a -le $attachments.Count; $a++) {
            $attachment = Track-Com ($attachments.Item($a))
            [pscustomobject]@{ Size = $attachment.Size; NameCharacters = ([string]$attachment.FileName).Length }
        })
        $body = [string]$mail.Body
        $html = [string]$mail.HTMLBody
        $resolved = Track-Com ($session.GetItemFromID($mail.EntryID, $inbox.StoreID))
        $samples += [pscustomobject]@{
            Sample = $samples.Count + 1
            SubjectCharacters = ([string]$mail.Subject).Length
            SenderAddressCharacters = ([string]$mail.SenderEmailAddress).Length
            BodyCharacters = $body.Length
            HTMLCharacters = $html.Length
            DownloadState = [int]$mail.DownloadState
            AttachmentCount = $attachments.Count
            AttachmentMetadata = $attachmentSizes
            LookupByIdSucceeded = ($resolved.EntryID -eq $mail.EntryID)
            UnreadUnchanged = ($wasUnread -eq $mail.UnRead)
        }
        $body = $null
        $html = $null
    }
    $unreadItems = Track-Com ($items.Restrict('[UnRead] = True'))
    [pscustomobject]@{
        CheckedAt = [DateTimeOffset]::Now.ToString('o')
        OutlookVersion = $app.Version
        StoreCount = $stores.Count
        Stores = $storeResults
        Offline = $session.Offline
        ExchangeConnectionMode = [int]$session.ExchangeConnectionMode
        InboxItems = $items.Count
        InboxUnreadBefore = $unreadBefore
        InboxUnreadAfter = $inbox.UnReadItemCount
        FilteredUnreadItems = $unreadItems.Count
        Samples = $samples
    } | ConvertTo-Json -Depth 7
} catch {
    [pscustomobject]@{ Status = 'Failed'; Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
} finally {
    for ($i = $objects.Count - 1; $i -ge 0; $i--) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($objects[$i]) } catch { }
    }
    # Do not call Outlook.Quit(): this is the user's existing session.
}
