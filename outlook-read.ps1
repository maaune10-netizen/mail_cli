# outlook-read.ps1 - READ-ONLY access to Classic Outlook via COM (win32).
# No send / no write / no delete / no move / no mark-read. Reading only, by design.
# The only thing written is an attachment copy into a LOCAL temp folder
# (attachment-save) - never to the mailbox.
# No app registration, no OAuth, no admin consent (uses the user's own Outlook).
#
# Agent self-sufficiency: pagination (Offset), date filters (Since/Before),
# JSON output, cross-folder search (-Scope all), HTML->text fallback,
# message metadata (importance/categories/flag), link extraction.
param(
  [ValidateSet('help','doctor','whoami','folders','list','unread','search','read','digest','attachments','attachment-save','links','calendar','contacts')]
  [string]$Cmd = 'list',
  [string]$Id = '',
  [string]$Query = '',
  [string]$Folder = '',
  [int]$N = 15,
  [int]$Offset = 0,
  [string]$Since = '',
  [string]$Before = '',
  [string]$Scope = '',
  [int]$Index = -1,
  [string]$Dest = '',
  [switch]$Json
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

# Create the COM object, but do NOT die here - so `doctor` can diagnose failures.
$ol = $null; $comError = $null
try { $ol = New-Object -ComObject Outlook.Application } catch { $comError = $_.Exception.Message }
$ns = $null
if ($ol) { try { $ns = $ol.GetNamespace('MAPI') } catch { $comError = $_.Exception.Message } }
if (-not $ol -and $Cmd -ne 'doctor') {
  Write-Error "Outlook COM unavailable. Is Classic Outlook installed and signed in? ($comError)"
  exit 1
}

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

# Strip any path from an attachment name and make it filesystem-safe.
function Sanitize-Name([string]$n) {
  if (-not $n) { return 'attachment.bin' }
  $n = $n -replace '.*[\\/]', ''
  $n = $n -replace '[<>:"|?*\x00-\x1f]', '_'
  $n = $n.Trim('. ')
  if (-not $n) { return 'attachment.bin' }
  return $n
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

function Get-BodyText($m) {
  $b = GetVal { $m.Body } ''
  if ($b -and "$b".Trim().Length -gt 0) { return $b }
  $h = GetVal { $m.HTMLBody } ''
  $t = Html-ToText $h
  if ("$t".Trim().Length -eq 0 -and $h -match '(?i)<img') { return '[image-only email - no text body]' }
  return $t
}

# Sender SMTP email. Internal Exchange senders give an X500 DN in SenderEmailAddress;
# resolve the real SMTP address via GetExchangeUser().
function Get-FromEmail($m) {
  $e = GetVal { $m.SenderEmailAddress } ''
  $t = GetVal { $m.SenderEmailType } ''
  if ($t -eq 'EX' -or $e -match '^/O=') {
    $smtp = GetVal { $m.Sender.GetExchangeUser().PrimarySmtpAddress } ''
    if ($smtp) { return $smtp }
  }
  return $e
}

$IMPORTANCE = @{ 0 = 'Low'; 1 = 'Normal'; 2 = 'High' }
$FLAGSTATUS = @{ 0 = 'None'; 1 = 'Complete'; 2 = 'Flagged' }
function Get-Importance($m) { $v = [int](GetVal { $m.Importance } 1); if ($IMPORTANCE.ContainsKey($v)) { return $IMPORTANCE[$v] } return 'Normal' }
function Get-Flag($m) { $v = [int](GetVal { $m.FlagStatus } 0); if ($FLAGSTATUS.ContainsKey($v)) { return $FLAGSTATUS[$v] } return 'None' }
function Get-Categories($m) { return (GetVal { "$($m.Categories)" } '') }

# Distinct links from the HTML (href) and the plain-text body.
function Get-Links($m) {
  $urls = New-Object System.Collections.Generic.List[string]
  $html = GetVal { $m.HTMLBody } ''
  if ($html) {
    foreach ($mm in [regex]::Matches($html, 'href\s*=\s*["'']([^"'']+)["'']', 'IgnoreCase')) {
      $u = $mm.Groups[1].Value
      if ($u -match '^https?://') { $urls.Add($u) }
    }
  }
  $body = GetVal { $m.Body } ''
  if ($body) {
    foreach ($mm in [regex]::Matches($body, 'https?://[^\s<>"'')]+')) { $urls.Add($mm.Value.TrimEnd('.', ',', ';', ')')) }
  }
  return ($urls | Select-Object -Unique)
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
        n           = $taken
        unread      = [bool](GetVal { $m.Unread } $false)
        date        = $dateStr
        from        = GetVal { $m.SenderName } ''
        fromEmail   = Get-FromEmail $m
        subject     = GetVal { $m.Subject } ''
        importance  = Get-Importance $m
        categories  = Get-Categories $m
        flag        = Get-Flag $m
        attachments = [int](GetVal { $m.Attachments.Count } 0)
        id          = GetVal { $m.EntryID } ''
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
      $imp = if ($r.importance -ne 'Normal') { " <$($r.importance)>" } else { '' }
      $cat = if ($r.categories) { " {$($r.categories)}" } else { '' }
      $lines += "[$($r.n)] $u $($r.date) | $(P $r.from 35) <$(P $r.fromEmail 35)> | $(P $r.subject 90)$imp$cat$a | id=$($r.id)"
    }
  }
  return ($lines -join "`n")
}

function Show-Body($m, [int]$max) {
  $att = [int](GetVal { $m.Attachments.Count } 0)
  $attNames = if ($att -gt 0) { ($m.Attachments | ForEach-Object { $_.FileName }) -join ', ' } else { '' }
  $links = @(Get-Links $m)
  if ($Json) {
    return (ConvertTo-Json -InputObject ([pscustomobject]@{
      from        = GetVal { $m.SenderName } ''
      fromEmail   = Get-FromEmail $m
      to          = GetVal { $m.To } ''
      cc          = GetVal { $m.CC } ''
      subject     = GetVal { $m.Subject } ''
      date        = "$(GetVal { $m.ReceivedTime } '')"
      importance  = Get-Importance $m
      categories  = Get-Categories $m
      flag        = Get-Flag $m
      attachments = $attNames
      links       = @($links)
      body        = (P (Get-BodyText $m) $max)
    }) -Depth 4)
  }
  "From: $(GetVal { $m.SenderName } '') <$(Get-FromEmail $m)>"
  "To: $(GetVal { $m.To } '')"
  if ("$(GetVal { $m.CC } '')".Trim().Length -gt 0) { "Cc: $(GetVal { $m.CC } '')" }
  "Subject: $(GetVal { $m.Subject } '')"
  "Date: $(GetVal { $m.ReceivedTime } '')"
  "Importance: $(Get-Importance $m) | Flag: $(Get-Flag $m) | Categories: $(Get-Categories $m)"
  if ($att -gt 0) { "Attachments: $attNames" }
  "---"
  P (Get-BodyText $m) $max
  if ($links.Count -gt 0) {
    ""
    "Links ($($links.Count)):"
    $links | Select-Object -First 30 | ForEach-Object { "- $_" }
  }
}

$script:docFail = 0
function Doc([string]$label, [string]$value, [bool]$bad = $false) {
  if ($bad) { $script:docFail++ }
  $mark = 'OK  '; if ($bad) { $mark = 'FAIL' }
  return "$mark  $label : $value"
}

switch ($Cmd) {
  'help' {
@"
outlook-read - READ-ONLY Outlook access (no send, no write, no delete)

Commands:
  doctor                     check the environment (Outlook, COM, account, inbox)
  whoami                     list accounts
  folders [-Folder N]        list folders (or children of N)
  list    [-Folder N] [-N n] [-Offset k] [-Since d] [-Before d] [-Json]
  unread  [-N n] [-Offset k] [-Json]
  search  -Query "text" [-Folder N] [-Scope all] [-N n] [-Offset k] [-Json]
  read    -Id <entryid> [-Json]
  digest  [-Folder N] [-N n] [-Offset k] [-Since d] [-Before d]
  attachments    -Id <entryid> [-Json]
  attachment-save -Id <entryid> [-Index n] [-Dest dir] [-Json]
  links   -Id <entryid> [-Json]
  calendar [-N n] [-Offset k] [-Since d] [-Before d] [-Json]
  contacts [-N n] [-Offset k] [-Json]

Dates: YYYY-MM-DD or ISO.  -Scope all searches every mail folder.
attachment-save writes a copy to a LOCAL temp folder only (never the mailbox).
"@
  }
  'doctor' {
    Doc 'PowerShell' "$($PSVersionTable.PSVersion)"
    $olk = Get-Process olk -ErrorAction SilentlyContinue
    if ($olk) { Doc 'New Outlook' 'running (classic is still required)' $true } else { Doc 'New Outlook' 'not running' }
    $proc = Get-Process OUTLOOK -ErrorAction SilentlyContinue
    if ($proc) { Doc 'Classic Outlook process' "running (pid $($proc[0].Id))" } else { Doc 'Classic Outlook process' 'not running (COM can start it)' }
    Doc 'Outlook COM' $(if ($ol) { 'available' } else { "FAILED: $comError" }) ($null -eq $ol)
    if ($ns) {
      $accs = @($ns.Accounts)
      if ($accs.Count -gt 0) { foreach ($a in $accs) { Doc 'Account' "$($a.SmtpAddress) ($($a.DisplayName))" } }
      else { Doc 'Account' 'none found in the Outlook profile' $true }
      $ib = GetVal { $ns.GetDefaultFolder(6) } $null
      if ($ib) { Doc 'Inbox' "$($ib.Items.Count) items" } else { Doc 'Inbox' 'not accessible' $true }
      Doc 'Stores' "$(@($ns.Stores).Count)"
    }
    if ($script:docFail -eq 0) { "DOCTOR: OK" } else { "DOCTOR: $($script:docFail) problem(s) found"; exit 1 }
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
  'links' {
    $m = $ns.GetItemFromID($Id)
    $links = @(Get-Links $m)
    if ($Json) { ConvertTo-Json -InputObject ([pscustomobject]@{ subject = (GetVal { $m.Subject } ''); count = $links.Count; links = $links }) -Depth 4 }
    else { if ($links.Count -eq 0) { '(no links)' } else { $links | ForEach-Object { $_ } } }
  }
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
      $imp = Get-Importance $m; if ($imp -eq 'Normal') { $imp = '' } else { $imp = " <$imp>" }
      "=== [$i] $unread $date | $(P (GetVal { $m.SenderName } '') 45) <$(Get-FromEmail $m)> | $(P (GetVal { $m.Subject } '') 120)$imp"
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
  'attachment-save' {
    $m = $ns.GetItemFromID($Id)
    $att = $m.Attachments
    if ($att.Count -eq 0) { '(no attachments)'; break }
    if (-not $Dest) {
      $key = (($Id -replace '[^A-Za-z0-9]', ''))
      if ($key.Length -gt 24) { $key = $key.Substring(0, 24) }
      $Dest = Join-Path $env:TEMP ("outlook-read\$key")
    }
    New-Item -ItemType Directory -Force -Path $Dest | Out-Null
    $saved = @()
    $n = 0
    foreach ($a in $att) {
      $n++
      if ($Index -ge 0 -and $n -ne $Index) { continue }
      $name = Sanitize-Name $a.FileName
      $path = Join-Path $Dest $name
      if (Test-Path $path) {
        $base = [IO.Path]::GetFileNameWithoutExtension($name)
        $ext = [IO.Path]::GetExtension($name)
        $path = Join-Path $Dest "$base-$n$ext"
      }
      $a.SaveAsFile($path)
      $saved += [pscustomobject]@{ index = $n; file = $path; size = $a.Size }
    }
    if ($Json) { JsonArr $saved } else { $saved | ForEach-Object { "saved=$($_.file) size=$($_.size)" } }
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
