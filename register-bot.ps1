# register-bot.ps1 - run the Telegram bot as a resident, auto-restarting task.
#
# The bot polls Telegram for commands, sends real-time alerts for important
# mail, and sends the digest at the configured times - so there is no need for
# separate scheduled report tasks. It runs at logon in the user's session
# (Classic Outlook COM requires it).
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File register-bot.ps1
#   powershell -ExecutionPolicy Bypass -File register-bot.ps1 -KeepReportTask
#   powershell -ExecutionPolicy Bypass -File register-bot.ps1 -Remove
param(
  [string]$TaskName = 'AOU-Mail-Bot',
  [string]$OldReportTask = 'AOU-Mail-Report',
  [switch]$KeepReportTask,
  [switch]$Remove,
  [switch]$NoStart
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$bot = Join-Path $repo 'bot.ps1'

if ($Remove) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
  "Removed scheduled task: $TaskName"
  exit 0
}

if (-not (Test-Path $bot)) { throw "bot not found: $bot" }

if (-not $KeepReportTask) {
  if (Get-ScheduledTask -TaskName $OldReportTask -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $OldReportTask -Confirm:$false
    "Removed old task: $OldReportTask (the bot handles scheduling now)"
  }
}

$argument = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$bot`""
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument -WorkingDirectory $repo
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet `
  -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
  -StartWhenAvailable `
  -MultipleInstances IgnoreNew `
  -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) `
  -ExecutionTimeLimit ([TimeSpan]::Zero)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null

"Registered: $TaskName (runs at logon, auto-restart)"
if (-not $NoStart) {
  Start-ScheduledTask -TaskName $TaskName
  "Started now."
}
""
"Status : Get-ScheduledTask -TaskName $TaskName"
"Log    : $env:LOCALAPPDATA\outlook-read\bot.log"
"Stop   : powershell -File register-bot.ps1 -Remove"
