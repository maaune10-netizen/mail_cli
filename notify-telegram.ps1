# notify-telegram.ps1 - send the mail report to a Telegram chat.
# The bot token and chat id are read from a JSON config OUTSIDE the repo
# (default: %LOCALAPPDATA%\outlook-read\telegram.json). Never commit secrets.
#
# Config format:
#   { "token": "<bot token>", "chat_id": "<chat id>" }
#
# Usage:
#   powershell -File notify-telegram.ps1
#   powershell -File notify-telegram.ps1 -Message "hello"
#   powershell -File notify-telegram.ps1 -Report other.txt
param(
  [string]$Config = (Join-Path $env:LOCALAPPDATA 'outlook-read\telegram.json'),
  [string]$Report = (Join-Path $env:LOCALAPPDATA 'outlook-read\last-report.txt'),
  [string]$Message = '',
  [int]$MaxLen = 4000
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Config)) { Write-Error "config not found: $Config"; exit 2 }
$cfg = Get-Content $Config -Raw -Encoding UTF8 | ConvertFrom-Json
$token = $cfg.token
$chat = "$($cfg.chat_id)".Trim()
if (-not $token) { Write-Error "bot token missing in $Config"; exit 2 }
if (-not $chat) { Write-Error "chat_id missing in $Config (send a message to the bot, then set it)"; exit 2 }

if ($Message) { $text = $Message }
elseif (Test-Path $Report) { $text = Get-Content $Report -Raw -Encoding UTF8 }
else { Write-Error "report not found: $Report"; exit 2 }

$text = "$text".Trim()
if ($text.Length -eq 0) { $text = '(تقرير فارغ)' }
if ($text.Length -gt $MaxLen) { $text = $text.Substring(0, $MaxLen) + "`n… (مقطوع)" }

$payload = @{ chat_id = $chat; text = $text; disable_web_page_preview = $true } | ConvertTo-Json -Compress
$bytes = [Text.Encoding]::UTF8.GetBytes($payload)
$uri = "https://api.telegram.org/bot$token/sendMessage"

try {
  $r = Invoke-RestMethod -Uri $uri -Method Post -ContentType 'application/json; charset=utf-8' -Body $bytes -TimeoutSec 30
  "sent: ok=$($r.ok) message_id=$($r.result.message_id)"
} catch {
  $msg = $_.Exception.Message
  if ($_.ErrorDetails.Message) { $msg = "$msg | $($_.ErrorDetails.Message)" }
  Write-Error "telegram send failed: $msg"
  exit 1
}
