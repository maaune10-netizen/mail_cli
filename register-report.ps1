# register-report.ps1 - schedule the daily mail report (Windows Task Scheduler).
#
# Creates (or updates) a task that runs the read-only report twice a day.
# Runs only when the user is logged on, in the user's session, so Classic
# Outlook COM works.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File register-report.ps1
#   powershell -ExecutionPolicy Bypass -File register-report.ps1 -Times 10:00,18:00
#   powershell -ExecutionPolicy Bypass -File register-report.ps1 -Remove
param(
  [string]$Times = '10:00,18:00',
  [string]$TaskName = 'AOU-Mail-Report',
  [int]$Count = 50,
  [switch]$Remove
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$engine = Join-Path $repo 'outlook-read.ps1'
$outDir = Join-Path $env:LOCALAPPDATA 'outlook-read'
$outFile = Join-Path $outDir 'last-report.txt'
$stateFile = Join-Path $outDir 'last-run.txt'

if ($Remove) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
  "Removed scheduled task: $TaskName"
  exit 0
}

if (-not (Test-Path $engine)) { throw "engine not found: $engine" }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$argument = "-NoProfile -ExecutionPolicy Bypass -File `"$engine`" -Cmd report -Update -N $Count -State `"$stateFile`" -Out `"$outFile`""
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument -WorkingDirectory $repo
$timesList = @($Times -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$triggers = foreach ($t in $timesList) { New-ScheduledTaskTrigger -Daily -At ([datetime]::Parse($t)) }
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -Force | Out-Null

"Registered: $TaskName"
"  times : $($timesList -join ', ')"
"  report: $outFile"
"  state : $stateFile"
""
"Run now to test:  Start-ScheduledTask -TaskName $TaskName"
"Remove:           powershell -File register-report.ps1 -Remove"
