# outlook-read.ps1 - READ-ONLY access to Classic Outlook via COM (win32).
# No send / no write / no delete / no mark-read. Reading only, by design.
# No app registration, no OAuth, no admin consent (uses the user's own Outlook).
#
# Agent self-sufficiency: pagination (Offset), date filters (Since/Before),
# JSON output, cross-folder search (-Scope all), HTML->text fallback.
param(
  [ValidateSet('help','whoami','folders','list','unread','search','read','digest','attachments','calendar','contacts')]
  [string]$Cmd = 'list',
  [string]$Id = '',
  [string]$Query = '',
  [string]$Folder = '',
  [int]$N = 15,
  [int]$Offset = 0,
  [string]$Since = '',
  [string]$Before = '',
  [string]$Scope = '',
  [switch]$Json
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

try { $ol = New-Object -ComObject Outlook.Application }
catch { Write-Error "Outlook COM unavailable. Is classic Outlook installed and signed in? ($_)"; exit 1 }
$ns = $ol.GetNamespace('MAPI')

# Safe property access (PS 5.1 has no try-as-expression).
function GetVal([scriptblock]$sb, $def = '') {
  try { $v = & $sb; if ($null -eq $v) { return $def } return $v } catch { return $def }
}

function P([string]$s, [int]$max = 200) {
  if ($null -eq $s) { return '' }
  $s = ($s -replace '\s+', ' ').Trim()
  if ($s.Length -gt $max) { $s = $s.Substring(0, $max) + '...' }
  return $s
}

function Html-ToText([string]$html) {
  if (-not $html) { return '' }
  $t = $html -replace '(?is)<(script|style)[^>]*>.*?</\1>', ' '
  $t = $t -replace '(?i)<br\s*/?>', "`n"
  $t = $t -replace '(?i)</p>', "`n"
  $t = $t -replace '(?s)<[^>]+>', ' '
  $t = $t -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"'
  return ($t -replace '[ \t]+', ' ')
}

# Sender SMTP email. Internal Exchange senders give an X500 DN in SenderEmailAddress;
# fall back to the PR_SMTP_ADDRESS MAPI property (no object-model guard prompt).
function Get-FromEmail($m) {
  $e = GetVal { $m.SenderEmailAddress } ''
  $t = GetVal { $m.SenderEmailType } ''
  if ($t -eq 'EX' -or $e -match '^/O=') {
    $smtp = GetVal { $m.Sender.GetExchangeUser().PrimarySmtpAddress } ''
    if ($smtp) { return $smtp }
  }
  return $e
}

function Get-BodyText($m) {
  $b = GetVal { $m.Body } ''
  if ($b -and "$b".Trim().Length -gt 0) { return $b }
  $h = GetVal { $m.HTMLBody } ''
  $t = Html-ToText $h
  if ("$t".Trim().Length -eq 0 -and $h -match '(?i)<img') { return '[image-only email - no text body]' }
  return $t
}

# default-folder aliases (OlDefaultFolders)
$DEFAULT_FOLDERS = @{
  'inbox' = 6; 'sent' = 5; 'sentitems' = 5; 'drafts' = 16; 'deleteditems' = 23;
  'deleted' = 23; 'outbox' = 4; 'calendar' = 9; 'contacts' = 10; 'tasks' = 13; 'notes' = 12; 'journal' = 11
}

function Find-Folder($root, [string]$name, [int]$depth) {
  if ($depth -le 0) { return $null }
  foreach ($f in $root.Folders) {
    if ($f.Name -eq $name) { return $f }
    $r = Find-Folder $f $name ($depth - 1)
    if ($r) { return $r }
  }
  return $null
}

function Resolve-Folder([string]$name, [int]$fallback = 6) {
  if (-not $name) { return $ns.GetDefaultFolder($fallback) }
  $k = ($name.ToLower() -replace '\s', '')
  if ($DEFAULT_FOLDERS.ContainsKey($k)) { return $ns.GetDefaultFolder($DEFAULT_FOLDERS[$k]) }
  foreach ($st in $ns.Stores) {
    foreach ($f in $st.GetRootFolder().Folders) {
      if ($f.Name -eq $name) { return $f }
      $r = Find-Folder $f $name 6
      if ($r) { return $r }
    }
  }
  throw "folder not found: $name"
}

function Get-AllMailFolders($root, [int]$depth) {
  $res = @()
  if ($depth -le 0) { return $res }
  foreach ($f in $root.Folders) {
    if (GetVal { $f.DefaultItemType -eq 0 } $false) { $res += $f }
    $res += Get-AllMailFolders $f ($depth - 1)
  }
  return $res
}

function Parse-Date([string]$s) {
  if (-not $s) { return $null }
  try { return [datetime]::Parse($s) } catch { throw "bad date: $s (use e.g. 2026-09-01)" }
}
$sinceD = Parse-Date $Since
$beforeD = Parse-Date $Before

function Item-Date($m, [string]$kind) {
  if ($kind -eq 'appt') { return (GetVal { $m.Start } $null) }
  return (GetVal { $m.ReceivedTime } $null)
}

function JsonArr($o) { return (ConvertTo-Json -InputObject @($o) -Depth 4) }

# $kind: mail|appt|contact
function Build-Records($items, [string]$kind, [int]$max, [int]$skip, [bool]$json) {
  $recs = @()
  $matched = 0
  $taken = 0
  foreach ($m in $items) {
    $d = Item-Date $m $kind
    if ($sinceD -and $d -and $d -lt $sinceD) { continue }
    if ($beforeD -and $d -and $d -ge $beforeD) { continue }
    $matched++
    if ($matched -le $skip) { continue }
    if ($taken -ge $max) { break }
    $taken++
    $dateStr = if ($d) { $d.ToString('yyyy-MM-dd HH:mm') } else { '' }
    if ($kind -eq 'appt') {
      $recs += [pscustomobject]@{ n = $taken; date = $dateStr; subject = (GetVal { $m.Subject } ''); location = (GetVal { $m.Location } ''); id = (GetVal { $m.EntryID } '') }
    } elseif ($kind -eq 'contact') {
      $recs += [pscustomobject]@{ n = $taken; name = (GetVal { $m.FullName } ''); email = (GetVal { $m.Email1Address } ''); id = (GetVal { $m.EntryID } '') }
    } else {
      $recs += [pscustomobject]@{
        n          = $taken
        unread     = [bool](GetVal { $m.Unread } $false)
        date       = $dateStr
        from       = GetVal { $m.SenderName } ''
        fromEmail  = Get-FromEmail $m
        subject    = GetVal { $m.Subject } ''
        attachments = [int](GetVal { $m.Attachments.Count } 0)
        id         = GetVal { $m.EntryID } ''
      }
    }
  }
  if ($json) { return (JsonArr $recs) }
  if ($recs.Count -eq 0) { return '(no items)' }
  $lines = @()
  foreach ($r in $recs) {
    if ($kind -eq 'appt') { $lines += "[$($r.n)] $($r.date) | $(P $r.subject 100) | $(P $r.location 40) | id=$($r.id)" }
    elseif ($kind -eq 'contact') { $lines += "[$($r.n)] $(P $r.name 40) | $(P $r.email 45) | id=$($r.id)" }
    else {
      $u = if ($r.unread) { '*' } else { ' ' }
      $a = if ($r.attachments -gt 0) { " [$($r.attachments) att]" } else { '' }
      $lines += "[$($r.n)] $u $($r.date) | $(P $r.from 35) <$(P $r.fromEmail 35)> | $(P $r.subject 90)$a | id=$($r.id)"
    }
  }
  return ($lines -join "`n")
}

function Show-Body($m, [int]$max) {
  $att = [int](GetVal { $m.Attachments.Count } 0)
  $attNames = if ($att -gt 0) { ($m.Attachments | ForEach-Object { $_.FileName }) -join ', ' } else { '' }
  if ($Json) {
    return (ConvertTo-Json -InputObject ([pscustomobject]@{
      from        = GetVal { $m.SenderName } ''
      fromEmail   = Get-FromEmail $m
      to          = GetVal { $m.To } ''
      cc          = GetVal { $m.CC } ''
      subject     = GetVal { $m.Subject } ''
      date        = "$(GetVal { $m.ReceivedTime } '')"
      attachments = $attNames
      body        = (P (Get-BodyText $m) $max)
    }) -Depth 4)
  }
  "From: $(GetVal { $m.SenderName } '') <$(Get-FromEmail $m)>"
  "To: $(GetVal { $m.To } '')"
  "Subject: $(GetVal { $m.Subject } '')"
  "Date: $(GetVal { $m.ReceivedTime } '')"
  if ($att -gt 0) { "Attachments: $attNames" }
  "---"
  P (Get-BodyText $m) $max
}

switch ($Cmd) {
  'help' {
@"
outlook-read - READ-ONLY Outlook access (no send, no write, no delete)

Commands:
  whoami                     list accounts
  folders [-Folder N]        list folders (or children of N)
  list    [-Folder N] [-N n] [-Offset k] [-Since d] [-Before d] [-Json]
  unread  [-N n] [-Offset k] [-Json]
  search  -Query "text" [-Folder N] [-Scope all] [-N n] [-Offset k] [-Json]
  read    -Id <entryid> [-Json]
  digest  [-Folder N] [-N n] [-Offset k] [-Since d] [-Before d]
  attachments -Id <entryid> [-Json]
  calendar [-N n] [-Offset k] [-Since d] [-Before d] [-Json]
  contacts [-N n] [-Offset k] [-Json]

Dates: YYYY-MM-DD or ISO.  -Scope all searches every mail folder.
"@
  }
  'whoami' {
    $o = foreach ($a in $ns.Accounts) { [pscustomobject]@{ account = $a.SmtpAddress; display = $a.DisplayName; type = $a.AccountType } }
    if ($Json) { JsonArr $o } else { $o | ForEach-Object { "account=$($_.account) display=$($_.display) type=$($_.type)" } }
  }
  'folders' {
    $f = if ($Folder) { Resolve-Folder $Folder } else { $ns.GetDefaultFolder(6).Parent }
    $o = foreach ($sub in $f.Folders) { [pscustomobject]@{ folder = $sub.Name; items = $sub.Items.Count } }
    if ($Json) { JsonArr $o } else { $o | ForEach-Object { "folder=$($_.folder) items=$($_.items)" } }
  }
  'list' {
    $f = Resolve-Folder $Folder
    $it = $f.Items; try { $it.Sort('[ReceivedTime]', $true) } catch {}
    Build-Records $it 'mail' $N $Offset $Json
  }
  'unread' {
    $f = Resolve-Folder $Folder
    $it = $f.Items.Restrict('[Unread]=true'); try { $it.Sort('[ReceivedTime]', $true) } catch {}
    Build-Records $it 'mail' $N $Offset $Json
  }
  'search' {
    $q = [regex]::Escape($Query)
    if ($Scope -eq 'all') {
      $found = @()
      foreach ($fold in (Get-AllMailFolders $ns.GetDefaultFolder(6).Parent 6)) {
        try { foreach ($m in $fold.Items) {
          $hay = "$(GetVal { $m.Subject } '') $(GetVal { $m.SenderName } '') $(P (Get-BodyText $m) 200)"
          if ($hay -match $q) { $found += $m }
        } } catch {}
      }
      $found = $found | Sort-Object { GetVal { $_.ReceivedTime } ([datetime]::MinValue) } -Descending
      Build-Records $found 'mail' $N $Offset $Json
    } else {
      $f = Resolve-Folder $Folder
      $it = $f.Items; try { $it.Sort('[ReceivedTime]', $true) } catch {}
      $found = @()
      # ponytail: linear scan of cached folder; Restrict/DASL if >10k items
      foreach ($m in $it) {
        $hay = "$(GetVal { $m.Subject } '') $(GetVal { $m.SenderName } '') $(P (Get-BodyText $m) 200)"
        if ($hay -match $q) { $found += $m }
      }
      Build-Records $found 'mail' $N $Offset $Json
    }
  }
  'read' { Show-Body $ns.GetItemFromID($Id) 8000 }
  'digest' {
    $f = Resolve-Folder $Folder
    $it = $f.Items; try { $it.Sort('[ReceivedTime]', $true) } catch {}
    $i = 0; $skipped = 0
    foreach ($m in $it) {
      $d = Item-Date $m 'mail'
      if ($sinceD -and $d -and $d -lt $sinceD) { continue }
      if ($beforeD -and $d -and $d -ge $beforeD) { continue }
      if ($skipped -lt $Offset) { $skipped++; continue }
      if ($i -ge $N) { break }
      $i++
      $date = if ($d) { $d.ToString('yyyy-MM-dd HH:mm') } else { '' }
      $unread = if (GetVal { $m.Unread } $false) { '*' } else { ' ' }
      "=== [$i] $unread $date | $(P (GetVal { $m.SenderName } '') 45) <$(Get-FromEmail $m)> | $(P (GetVal { $m.Subject } '') 120)"
      P (Get-BodyText $m) 700
      ""
    }
  }
  'attachments' {
    $m = $ns.GetItemFromID($Id)
    $o = foreach ($a in $m.Attachments) { [pscustomobject]@{ file = $a.FileName; size = $a.Size; type = $a.Type } }
    if ($Json) { ConvertTo-Json -InputObject ([pscustomobject]@{ subject = (GetVal { $m.Subject } ''); attachments = @($o) }) -Depth 4 }
    else {
      "Subject: $(GetVal { $m.Subject } '')"
      if ($m.Attachments.Count -eq 0) { '(no attachments)' }
      $o | ForEach-Object { "att=$($_.file) size=$($_.size) type=$($_.type)" }
    }
  }
  'calendar' {
    $f = Resolve-Folder '' 9
    $it = $f.Items; try { $it.Sort('[Start]', $true) } catch {}
    Build-Records $it 'appt' $N $Offset $Json
  }
  'contacts' {
    $f = Resolve-Folder '' 10
    Build-Records $f.Items 'contact' $N $Offset $Json
  }
}
