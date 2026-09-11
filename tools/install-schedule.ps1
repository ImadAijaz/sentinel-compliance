# Installs the 15-minute Sentinel master sync as a Windows Scheduled Task.
#
# Why a scheduled task and not a cloud cron: the master folder is ~30 GB across ~28,000 files
# and both it and the destination are already OneDrive-synced onto this machine. Running here
# is plain local file I/O - no Graph API, no rate limits, no function timeouts, and nothing
# pushed through the cloud. OneDrive uploads the organized result to SharePoint by itself.
# Vercel's free plan allows one cron run a DAY, so it cannot do this at all.
#
# Runs as the signed-in user (it needs that user's OneDrive folders), starts hidden, and skips
# a run if the previous one is still going. -WakeToRun plus the battery settings keep it going
# through sleep, and -StartWhenAvailable catches up a run the machine missed while off.
#
#   powershell -ExecutionPolicy Bypass -File tools\install-schedule.ps1
#   powershell -ExecutionPolicy Bypass -File tools\install-schedule.ps1 -Remove

param([switch]$Remove, [int]$Minutes = 15)

$TaskName = "Sentinel Master Sync"
$Repo     = Split-Path -Parent $PSScriptRoot
$Script   = Join-Path $Repo "tools\sync-master.js"

if ($Remove) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
  Write-Host "Removed scheduled task '$TaskName'."
  return
}

$node = (Get-Command node -ErrorAction SilentlyContinue).Source
if (-not $node) { Write-Error "Node.js not found on PATH. Install Node, then re-run this."; return }
if (-not (Test-Path $Script)) { Write-Error "Cannot find $Script"; return }

# -NoProfile keeps startup fast; --quiet keeps it silent (everything still goes to the log file
# under Sentinel\_sync\sync-log.txt).
$action = New-ScheduledTaskAction -Execute $node -Argument "`"$Script`" --apply --quiet" -WorkingDirectory $Repo

# A daily trigger that repeats through the day. [TimeSpan]::MaxValue is rejected by the task
# scheduler on this build ("Duration: P99999999DT23H59M59S ... out of range"), and a bare -Once
# trigger stops repeating once its duration elapses. Bounding the repetition to 24 hours and
# letting the daily trigger re-arm it gives a run every $Minutes minutes, indefinitely.
$trigger = New-ScheduledTaskTrigger -Daily -At (Get-Date).Date.AddMinutes(2)
$trigger.Repetition = (New-ScheduledTaskTrigger -Once -At (Get-Date) `
  -RepetitionInterval (New-TimeSpan -Minutes $Minutes) `
  -RepetitionDuration (New-TimeSpan -Hours 24)).Repetition

$settings = New-ScheduledTaskSettingsSet `
  -MultipleInstances IgnoreNew `
  -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
  -StartWhenAvailable `
  -DontStopIfGoingOnBatteries `
  -AllowStartIfOnBatteries `
  -WakeToRun

$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
  -Settings $settings -Principal $principal `
  -Description "Files documents from the WCGTX master OneDrive folder into the Sentinel tree in Corporate Archives, every $Minutes minutes." -ErrorAction Stop | Out-Null

# Confirm rather than assume: Register-ScheduledTask reports some failures without stopping,
# and a first attempt printed "Installed" while nothing had actually been registered.
$check = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if (-not $check) { Write-Error "Task did not register."; return }

Write-Host "Installed '$TaskName' - runs every $Minutes minutes as $env:USERNAME."
Write-Host "  script : $Script"
Write-Host '  log    : (OneDrive)\Corporate Archives Directory - Documents\Sama Farooqui\Sentinel\_sync\sync-log.txt'
Write-Host ""
Write-Host "Run it once now with:  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "Check it with:         Get-ScheduledTaskInfo -TaskName '$TaskName'"
Write-Host 'Remove it with:        powershell -ExecutionPolicy Bypass -File tools\install-schedule.ps1 -Remove'
