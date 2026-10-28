$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../../powershell/worker.ps1') -Raw
Invoke-Expression $source.Substring(0, $source.IndexOf('function Folder-Key'))
function Assert-True($Value, $Message) { if (-not $Value) { throw $Message } }
function Assert-Invalid($Id, $Kind) {
    $errorText = $null
    try { [void](Parse-Id $Id $Kind) } catch { $errorText = $_.Exception.Message }
    Assert-True ($errorText -match '^INVALID_ID:') 'invalid identifier was accepted'
}
$id = Make-Id 'message' 'AABB' 'CCDD'
Assert-True ($script:idPartsCache.ContainsKey($id)) 'generated identifier was not cached'
Assert-True ($script:encodedIdCache.Count -eq 1) 'encoded identifier was not cached'
Assert-True ((Make-Id 'message' 'AABB' 'CCDD') -ceq $id) 'cached encoding changed the opaque identifier'
Assert-True ((Make-Id 'message' 'AABB' 'CCDE') -cne $id) 'encoding cache confused stores'
Assert-True ((Make-Id 'folder' 'AABB' 'CCDD') -cne $id) 'encoding cache confused kinds'
$parsed = Parse-Id $id 'message'
Assert-True ($parsed.entry -eq 'AABB' -and $parsed.store -eq 'CCDD') 'cached decoding differs'
$parsed.entry = 'FFFF'
Assert-True ((Parse-Id $id 'message').entry -eq 'AABB') 'caller changed cached components'
Assert-Invalid $id 'folder'
Assert-True (-not $script:idPartsCache.ContainsKey($id.ToLowerInvariant())) 'base64 identifiers became case insensitive'
Assert-Invalid $id.ToLowerInvariant() 'message'
Assert-Invalid (Make-Id 'message' 'not-hex' 'CCDD') 'message'
Assert-Invalid '!' 'message'
for ($i = 1; $i -le 300; $i++) { [void](Make-Id 'message' $i.ToString('X') 'CCDD') }
Assert-True ($script:idPartsCache.Count -le 256) 'identifier cache grew without bound'
Assert-True ($script:encodedIdCache.Count -le 256) 'encoding cache grew without bound'
Assert-True (-not $script:idPartsCache.ContainsKey($id)) 'bounded cache did not evict old entries'
Assert-True ((Parse-Id $id 'message').entry -eq 'AABB') 'evicted identifier no longer decodes'
$script:lookups = 0
$script:session = [pscustomobject]@{}
$script:session | Add-Member -MemberType ScriptMethod -Name GetItemFromID -Value {
    param($Entry, $Store)
    $script:lookups++
    return [pscustomobject]@{ Class = 43; Version = $script:lookups }
}
Assert-True ((Get-Mail $id).Version -eq 1) 'first lookup failed'
Assert-True ((Get-Mail $id).Version -eq 2) 'cached identifier reused a mail object'
Assert-True ($script:lookups -eq 2) 'fresh source lookup was skipped'
@{ pass = $true } | ConvertTo-Json -Compress
