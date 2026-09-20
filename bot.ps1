# bot.ps1 - resident Telegram bot for the read-only Outlook CLI.
#
# Features:
#   * commands from Telegram (/check, /time, /status, /unread, /last, /search, ...)
#   * real-time alerts for HIGH-importance (or flagged) mail
#   * scheduled digests at configurable times (default 10:00, 18:00)
#   * runs only when the user is logged on (Outlook COM needs the session)
#
# Config: %LOCALAPPDATA%\outlook-read\telegram.json
#   { "token": "...", "chat_id": "123", "times": ["10:00","18:00"],
#     "poll_seconds": 60, "alert_importance": "High", "alert_flagged": true,
#     "digest_count": 50 }
param(
  [string]$Config = (Join-Path $env:LOCALAPPDATA 'outlook-read\telegram.json'),
  [string]$Cmd = '',
  [switch]$Once,
  [switch]$Test,
  [switch]$Verbose2
)
$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$Repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$Engine = Join-Path $Repo 'outlook-read.ps1'
$DataDir = Join-Path $env:LOCALAPPDATA 'outlook-read'
$StateFile = Join-Path $DataDir 'last-run.txt'      # digest watermark
$AlertFile = Join-Path $DataDir 'last-alert.txt'    # alert watermark
$AlertIds = Join-Path $DataDir 'alerted.json'       # ids already alerted
$ReportFile = Join-Path $DataDir 'last-report.txt'
$OffsetFile = Join-Path $DataDir 'bot-offset.txt'
$LogFile = Join-Path $DataDir 'bot.log'
New-Item -ItemType Directory -Force -Path $DataDir | Out-Null

function Log([string]$m) {
  $line = "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))  $m"
  Add-Content -Path $LogFile -Value $line -Encoding UTF8
  if ($Verbose2) { Write-Host $line }
}

# ---------- config ----------
function Get-Cfg {
  if (-not (Test-Path $Config)) { throw "config not found: $Config" }
  return (Get-Content $Config -Raw -Encoding UTF8 | ConvertFrom-Json)
}
$cfg = Get-Cfg
$script:token = "$($cfg.token)"
$script:chat = "$($cfg.chat_id)".Trim()
if (-not $script:token) { throw 'token missing in config' }
if (-not $script:chat) { throw 'chat_id missing in config' }
# alert modes: off | important | all
$script:alertMode = "$($cfg.alert_mode)"
if (-not $script:alertMode) { if ("$($cfg.alert_importance)" -eq 'off') { $script:alertMode = 'off' } else { $script:alertMode = 'important' } }

# ---------- telegram ----------
function Tg([string]$method, $params, [int]$timeoutSec = 35) {
  $uri = "https://api.telegram.org/bot$($script:token)/$method"
  $json = $params | ConvertTo-Json -Compress -Depth 6
  $bytes = [Text.Encoding]::UTF8.GetBytes($json)
  return Invoke-RestMethod -Uri $uri -Method Post -ContentType 'application/json; charset=utf-8' -Body $bytes -TimeoutSec $timeoutSec
}

function Send-Text([string]$text) {
  $text = "$text".Trim()
  if (-not $text) { return }
  while ($text.Length -gt 3900) {
    $chunk = $text.Substring(0, 3900)
    $nl = $chunk.LastIndexOf("`n")
    if ($nl -gt 500) { $chunk = $chunk.Substring(0, $nl) }
    [void](Tg 'sendMessage' @{ chat_id = $script:chat; text = $chunk; disable_web_page_preview = $true })
    $text = $text.Substring($chunk.Length).TrimStart()
  }
  if ($text) { [void](Tg 'sendMessage' @{ chat_id = $script:chat; text = $text; disable_web_page_preview = $true }) }
}

# ---------- engine ----------
function Engine([string[]]$a) {
  return ((& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Engine @a) -join "`n")
}
function EngineJson([string[]]$a) {
  $raw = Engine $a
  try { return @($raw | ConvertFrom-Json) } catch { Log "json parse failed: $($_.Exception.Message)"; return @() }
}

function Format-Rows($rows, [int]$max = 10) {
  $t = @()
  $i = 0
  foreach ($r in $rows) {
    if ($i -ge $max) { break }
    $i++
    $u = if ($r.unread) { '🔵' } else { '⚪' }
    $imp = if ($r.importance -and $r.importance -ne 'Normal') { " [$($r.importance)]" } else { '' }
    $att = if ($r.attachments -gt 0) { " 📎$($r.attachments)" } else { '' }
    $t += "$u $($r.date) | $($r.from)$imp$att"
    $t += "     $($r.subject)"
  }
  return ($t -join "`n")
}

# ---------- actions ----------
function Send-Digest([switch]$Update) {
  $count = 50; if ($cfg.digest_count) { $count = [int]$cfg.digest_count }
  $a = @('-Cmd', 'report', '-N', "$count", '-Out', $ReportFile, '-State', $StateFile)
  if ($Update) { $a += '-Update' }
  [void](Engine $a)
  if (Test-Path $ReportFile) { Send-Text (Get-Content $ReportFile -Raw -Encoding UTF8) }
  else { Send-Text '(فشل توليد التقرير)' }
}

function Get-Alerted {
  if (Test-Path $AlertIds) { try { return @(Get-Content $AlertIds -Raw -Encoding UTF8 | ConvertFrom-Json) } catch {} }
  return @()
}
function Save-Alerted($ids) {
  $bounded = @($ids | Select-Object -Last 500)
  [IO.File]::WriteAllText($AlertIds, (ConvertTo-Json -Compress -InputObject @($bounded)), (New-Object System.Text.UTF8Encoding($false)))
}

function Check-Alerts {
  if ($script:alertMode -eq 'off') { return }
  $since = $null
  if (Test-Path $AlertFile) {
    $raw = "$(Get-Content $AlertFile -Raw -Encoding UTF8)".Trim()
    if ($raw) { try { $since = [datetime]::Parse($raw).AddMinutes(-5) } catch {} }
  }
  if (-not $since) { $since = (Get-Date).AddMinutes(-30) }
  $rows = EngineJson @('-Cmd', 'list', '-Json', '-Snippet', '-N', '50', '-Since', $since.ToString('s'))
  $re = "$($cfg.alert_regex)"; if (-not $re) { $re = $IMPORTANT_RE }
  $alerted = Get-Alerted
  $sent = 0
  foreach ($r in $rows) {
    if ($alerted -contains $r.id) { continue }
    $reasons = @()
    if ("$($r.importance)" -eq 'High') { $reasons += 'أهمية عالية' }
    if ($cfg.alert_flagged -and "$($r.flag)" -eq 'Flagged') { $reasons += 'معلّمة' }
    if ("$($r.subject) $($r.snippet)" -match $re) { $reasons += 'تحتاج انتباه' }
    $isImportant = $reasons.Count -gt 0
    # mode 'important' -> only flagged-by-content; mode 'all' -> every new mail
    if ($script:alertMode -eq 'important' -and -not $isImportant) { continue }

    if ($isImportant) {
      $full = EngineJson @('-Cmd', 'read', '-Json', '-Id', $r.id)
      $body = "$($full.body)"; if (-not $body) { $body = "$($r.snippet)" }
      $msg = "🚨 رسالة مهمة ($($reasons -join ' + '))`n`nمن: $($r.from) <$($r.fromEmail)>`nالتاريخ: $($r.date)`nالموضوع: $($r.subject)`n`n$body"
    } else {
      $snip = "$($r.snippet)"; if ($snip.Length -gt 400) { $snip = $snip.Substring(0, 400) + '...' }
      $msg = "📧 رسالة جديدة`n`nمن: $($r.from) <$($r.fromEmail)>`nالتاريخ: $($r.date)`nالموضوع: $($r.subject)"
      if ($snip) { $msg += "`n`n$snip" }
    }
    Send-Text $msg
    $alerted += $r.id
    $sent++
  }
  if ($sent -gt 0) { Save-Alerted $alerted; Log "alerts sent: $sent (mode $($script:alertMode))" }
  [IO.File]::WriteAllText($AlertFile, (Get-Date).ToString('o'), (New-Object System.Text.UTF8Encoding($false)))
}

# ---------- commands ----------
# Student-focused. ASCII names for the Telegram menu + Arabic aliases.
$ALIAS = @{
  'start' = 'help'; 'help' = 'help'; 'مساعدة' = 'help'
  'check' = 'check'; 'جديد' = 'check'
  'action' = 'action'; 'المطلوب' = 'action'; 'مطلوب' = 'action'
  'exams' = 'exams'; 'الاختبارات' = 'exams'; 'اختبارات' = 'exams'
  'unread' = 'unread'; 'غير_مقروء' = 'unread'; 'غيرمقروء' = 'unread'
  'last' = 'last'; 'آخر' = 'last'
  'read' = 'read'; 'قراءة' = 'read'
  'search' = 'search'; 'بحث' = 'search'
  'course' = 'course'; 'مقرر' = 'course'
  'time' = 'time'; 'وقت' = 'time'
  'alerts' = 'alerts'; 'تنبيه' = 'alerts'; 'تنبيهات' = 'alerts'
  'interval' = 'interval'; 'فحص' = 'interval'
}
$ACTION_RE = 'مطلوب|يجب عليك|الرجاء|يرجى|تسجيل|سجّل|دفع|رسوم|غرامة|تحقق|وثيقة|مستند|ارفع|رفع|نموذج|استبيان|عبّئ|عبئ|آخر موعد|deadline|submit|upload|verify|register|payment|أكمل|اكمل|renew'
$EXAM_RE = 'اختبار|كويز|امتحان|exam|quiz|midterm|final exam'
# What counts as "important" for real-time alerts (content, not just the High flag).
$IMPORTANT_RE = "$ACTION_RE|$EXAM_RE|جدول|نتيجة|رسوم|موعد|مراجعة|تنبيه"

$script:pending = ''
# Tap-only menu (no typing). Telegram sends the label back as a message.
$BUTTONS = [ordered]@{
  '🆕 الجديد'     = 'check'
  '📌 المطلوب'     = 'action'
  '📝 الاختبارات'  = 'exams'
  '🔵 غير المقروء' = 'unread'
  '📥 آخر الرسائل' = 'last'
  '🔎 بحث'         = 'search'
  '🎓 مقرر'        = 'course'
  '⏰ الأوقات'     = 'time'
  '🔔 التنبيهات'   = 'alerts'
  '❓ مساعدة'      = 'help'
}

function Send-Menu([string]$text) {
  $rows = @(
    @('🆕 الجديد', '📌 المطلوب'),
    @('📝 الاختبارات', '🔵 غير المقروء'),
    @('📥 آخر الرسائل', '🔎 بحث'),
    @('🎓 مقرر', '⏰ الأوقات'),
    @('🔔 التنبيهات', '❓ مساعدة')
  )
  $markup = @{ keyboard = $rows; resize_keyboard = $true; is_persistent = $true }
  [void](Tg 'sendMessage' @{ chat_id = $script:chat; text = $text; reply_markup = $markup })
}

$HELP = @"
🤖 بوت بريد الجامعة — كل شي بأزرار، بلا كتابة

🆕 الجديد — الرسائل الجديدة من آخر تشغيل
📌 المطلوب — رسائل تحتاج منك إجراء
📝 الاختبارات — مواعيد الاختبارات والكوِزات
🔵 غير المقروء — الرسائل غير المقروءة
📥 آخر الرسائل — آخر 10
🔎 بحث / 🎓 مقرر — يطلبان منك كلمة واحدة
⏰ الأوقات — تغيير وقت التقارير
🔔 التنبيهات — بدّل: مطفية / المهم فقط / كل رسالة
⏰ الأوقات — تغيير وقت التقارير
"@

function Handle([string]$text) {
  $text = "$text".Trim()
  if (-not $text) { return }

  $arg = ''
  if ($BUTTONS.Contains($text)) {
    $cmd = $BUTTONS[$text]
  }
  elseif ($script:pending -and -not $text.StartsWith('/')) {
    $cmd = $script:pending
    $arg = $text
    $script:pending = ''
  }
  else {
    $parts = $text -split '\s+', 2
    $key = $parts[0].ToLower().TrimStart('/')
    if ($parts.Count -gt 1) { $arg = $parts[1].Trim() }
    $cmd = if ($ALIAS.ContainsKey($key)) { $ALIAS[$key] } else { $key }
  }

  switch ($cmd) {
    'help' { Send-Menu $HELP }
    'check' {
      Send-Text '⏳ جارٍ الفحص...'
      Send-Digest -Update
    }
    'action' {
      $since = (Get-Date).AddDays(-30).ToString('s')
      $rows = EngineJson @('-Cmd', 'search', '-Scope', 'all', '-Regex', '-Query', $ACTION_RE, '-Since', $since, '-Json', '-N', '15')
      $body = Format-Rows $rows 15
      Send-Text ("📌 رسائل يبدو أنها تحتاج إجراء (آخر 30 يوم)`n`n" + $(if ($body) { $body } else { '(لا شيء)' }))
    }
    'exams' {
      $since = (Get-Date).AddDays(-90).ToString('s')
      $rows = EngineJson @('-Cmd', 'search', '-Scope', 'all', '-Regex', '-Query', $EXAM_RE, '-Since', $since, '-Json', '-N', '15')
      $body = Format-Rows $rows 15
      Send-Text ("📝 مواعيد الاختبارات والكوِزات`n`n" + $(if ($body) { $body } else { '(لا يوجد)' }))
    }
    'unread' {
      $rows = EngineJson @('-Cmd', 'unread', '-Json', '-N', '10')
      $body = Format-Rows $rows 10
      Send-Text ("🔵 غير مقروء`n`n" + $(if ($body) { $body } else { '(لا شيء)' }))
    }
    'last' {
      $n = if ($arg -match '^\d+$') { [int]$arg } else { 10 }      if ($n -gt 30) { $n = 30 }
      $rows = EngineJson @('-Cmd', 'list', '-Json', '-N', "$n")
      Send-Text ("📥 آخر $n رسالة`n`n" + (Format-Rows $rows $n))
    }
    'read' {
      if ($arg -notmatch '^\d+$') { Send-Text 'اكتب رقم الرسالة من /last — مثال: /read 3'; return }
      $n = [int]$arg; if ($n -lt 1) { $n = 1 }; if ($n -gt 30) { $n = 30 }
      $rows = EngineJson @('-Cmd', 'list', '-Json', '-N', "$n")
      if (@($rows).Count -lt $n) { Send-Text 'ما فيه رسالة بهذا الرقم. استخدم /last'; return }
      $id = @($rows)[$n - 1].id
      Send-Text (Engine @('-Cmd', 'read', '-Id', $id))
    }
    'search' {
      if (-not $arg) { $script:pending = 'search'; Send-Text '🔎 اكتب كلمة البحث 👇'; return }
      $rows = EngineJson @('-Cmd', 'search', '-Scope', 'all', '-Query', $arg, '-Json', '-N', '8')
      $body = Format-Rows $rows 8
      Send-Text ("🔎 نتائج: $arg`n`n" + $(if ($body) { $body } else { '(لا نتائج)' }))
    }
    'course' {
      if (-not $arg) { $script:pending = 'course'; Send-Text '🎓 اكتب رمز المقرر (مثل GR101) 👇'; return }
      $rows = EngineJson @('-Cmd', 'search', '-Scope', 'all', '-Query', $arg, '-Json', '-N', '10')
      $body = Format-Rows $rows 10
      Send-Text ("🎓 مقرر $arg`n`n" + $(if ($body) { $body } else { '(لا رسائل)' }))
    }
    'time' {
      if (-not $arg) { Send-Text "الأوقات الحالية: $((@($cfg.times)) -join ', ')`nللتغيير: /time 09:00,21:00"; return }
      $newTimes = @($arg -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d{1,2}:\d{2}$' })
      if ($newTimes.Count -eq 0) { Send-Text 'صيغة خاطئة. مثال: /time 09:00,21:00'; return }
      $c = Get-Cfg; $c.times = @($newTimes)
      [IO.File]::WriteAllText($Config, (ConvertTo-Json -InputObject $c -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
      $script:cfg = $c
      Send-Text "✅ صار وقت التقرير: $($newTimes -join ', ')"
    }
    'alerts' {
      $c = Get-Cfg
      $mode = "$($c.alert_mode)"; if (-not $mode) { $mode = $script:alertMode }
      if ($arg -in @('off', 'important', 'all')) { $mode = $arg }
      elseif (-not $arg) {
        if ($mode -eq 'off') { $mode = 'important' } elseif ($mode -eq 'important') { $mode = 'all' } else { $mode = 'off' }
      }
      $c.alert_mode = $mode
      [IO.File]::WriteAllText($Config, (ConvertTo-Json -InputObject $c -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
      $script:cfg = $c; $script:alertMode = $mode
      $desc = switch ($mode) { 'off' { 'مطفية' } 'all' { 'كل رسالة جديدة' } default { 'المهم فقط (نص كامل)' } }
      Send-Text "🔔 التنبيهات: $desc`n`nاضغط الزر مرة ثانية للتبديل، أو /alerts off|important|all"
    }
    'interval' {
      if ($arg -notmatch '^\d+$') { Send-Text "⏱ الفحص الحالي كل $($cfg.poll_seconds) ثانية — للتغيير: /interval 30"; return }
      $sec = [int]$arg; if ($sec -lt 20) { $sec = 20 }; if ($sec -gt 600) { $sec = 600 }
      $c = Get-Cfg; $c.poll_seconds = $sec
      [IO.File]::WriteAllText($Config, (ConvertTo-Json -InputObject $c -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
      $script:cfg = $c
      Send-Text "⏱ صار الفحص كل $sec ثانية"
    }
    default { Send-Menu "ما فهمت: $cmd`n`n$HELP" }
  }
}

# ---------- main loop ----------
$offset = 0
if (Test-Path $OffsetFile) { $raw = "$(Get-Content $OffsetFile -Raw -Encoding UTF8)".Trim(); if ($raw -match '^\d+$') { $offset = [int]$raw } }
function Save-Offset([int]$v) { [IO.File]::WriteAllText($OffsetFile, "$v", (New-Object System.Text.UTF8Encoding($false))) }

if ($Cmd) { Handle $Cmd; Log "cmd test: $Cmd"; exit 0 }

if ($Test) {
  Send-Text "🧪 اختبار البوت الجديد"
  Handle '/help'
  Handle '/action'
  Handle '/exams'
  Handle '/unread'
  Handle '/last 3'
  # confirm and clear pending updates so the normal run does not replay them
  try { $r = Tg 'getUpdates' @{ offset = $offset; timeout = 0; allowed_updates = @('message') } 20; foreach ($u in @($r.result)) { $offset = [int]$u.update_id + 1 } } catch {}
  Save-Offset $offset
  Log 'test mode done'
  exit 0
}

$lastPoll = [datetime]::MinValue
$fired = @{}
Log "bot started (repo=$Repo, once=$Once)"
Send-Menu "🟢 البوت شغّال. اختر من الأزرار 👇"

while ($true) {
  try {
    # 1) telegram commands (long poll)
    $r = Tg 'getUpdates' @{ offset = $offset; timeout = 20; allowed_updates = @('message') } 30
    foreach ($u in @($r.result)) {
      $offset = [int]$u.update_id + 1
      $m = $u.message
      if (-not $m) { continue }
      if ("$($m.chat.id)" -ne $script:chat) { Log "ignored message from chat $($m.chat.id)"; continue }
      if ($m.text) { Handle $m.text }
    }
    Save-Offset $offset

    # 2) real-time high-priority alerts (time-based)
    $now = Get-Date
    $pollSec = [int]("$($cfg.poll_seconds)"); if ($pollSec -lt 20) { $pollSec = 20 }
    if (($now - $lastPoll).TotalSeconds -ge $pollSec) {
      Check-Alerts
      $lastPoll = $now
    }

    # 3) scheduled digests
    foreach ($t in @($cfg.times)) {
      $target = $null; try { $target = [datetime]::Parse($t) } catch {}
      if ($target -and $now -ge $target -and $fired[$t] -ne $now.ToString('yyyy-MM-dd')) {
        Log "scheduled digest at $t"
        Send-Digest -Update
        $fired[$t] = $now.ToString('yyyy-MM-dd')
      }
    }
  } catch {
    Log "loop error: $($_.Exception.Message)"
    Start-Sleep -Seconds 5
  }
  if ($Once) { break }
  Start-Sleep -Seconds 2
}
if ($Once) { try { $r = Tg 'getUpdates' @{ offset = $offset; timeout = 0; allowed_updates = @('message') } 20; foreach ($u in @($r.result)) { $offset = [int]$u.update_id + 1 }; Save-Offset $offset } catch {} }
Log 'bot stopped'
