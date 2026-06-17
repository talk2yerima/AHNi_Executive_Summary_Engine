# schedule_task.ps1
# Registers a Windows Task Scheduler task to run dhis2_pull.py every 4 hours.
# Run once as Administrator: .\schedule_task.ps1

$TaskName   = "AHNi_DHIS2_Pull"
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Python     = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$Script     = Join-Path $ScriptDir "dhis2_pull.py"
$LogDir     = Join-Path $ScriptDir "logs"

if (-not (Test-Path $Python)) {
    Write-Error "Python venv not found at $Python. Run setup.ps1 first."
    exit 1
}

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

$Action  = New-ScheduledTaskAction `
    -Execute $Python `
    -Argument "--mode incremental" `
    -WorkingDirectory $ScriptDir

$Trigger = @(
    New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Hours 4) `
        -Once -At (Get-Date "06:00")
)

$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 10) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable

$Principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType S4U `
    -RunLevel Highest

# Remove existing task if present
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed existing task '$TaskName'."
}

Register-ScheduledTask `
    -TaskName  $TaskName `
    -Action    $Action `
    -Trigger   $Trigger `
    -Settings  $Settings `
    -Principal $Principal `
    -Description "Pulls DHIS2 ACEBAY data every 4 hours and writes to output Excel."

Write-Host ""
Write-Host "Task '$TaskName' registered successfully." -ForegroundColor Green
Write-Host "  Schedule : every 4 hours starting at 06:00"
Write-Host "  Mode     : incremental (use --mode full for a one-off manual pull)"
Write-Host "  Log file : $LogDir\dhis2_pull.log"
Write-Host ""
Write-Host "To run immediately: Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "To remove:          Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false"
