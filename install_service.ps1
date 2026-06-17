# install_service.ps1
# Installs AHNi Executive Service Engine as a Windows Service using NSSM.
# Double-click or run from any PowerShell window - auto-elevates to Admin.

$ServiceName = "AHNi_Executive_Service_Engine"
$DisplayName = "AHNi Executive Service Engine"
$Description = "Pulls PEPFAR HIV indicators from DHIS2 every 3 hours and uploads to Azure Blob Storage."

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Python    = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$Runner    = Join-Path $ScriptDir "service_runner.py"
$LogDir    = Join-Path $ScriptDir "logs"
$NssmDir   = Join-Path $ScriptDir "nssm"
$NssmExe   = Join-Path $NssmDir "nssm.exe"

# --- Auto-elevate to Administrator ---
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]"Administrator"
)
if (-not $isAdmin) {
    Write-Host "Relaunching as Administrator ..."
    $elevateArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Definition)`""
    Start-Process powershell -ArgumentList $elevateArgs -Verb RunAs
    exit
}

# --- Check prerequisites ---
if (-not (Test-Path $Python)) {
    Write-Error "Python venv not found at: $Python"
    Write-Host "Run:  python -m venv .venv  then  .venv\Scripts\pip install -r requirements.txt"
    Read-Host "Press Enter to exit"
    exit 1
}
if (-not (Test-Path (Join-Path $ScriptDir ".env"))) {
    Write-Error ".env not found. Copy .env.example to .env and fill in credentials."
    Read-Host "Press Enter to exit"
    exit 1
}

# --- Download NSSM if not present ---
if (-not (Test-Path $NssmExe)) {
    Write-Host "Downloading NSSM ..."
    $NssmZip     = Join-Path $env:TEMP "nssm.zip"
    $NssmExtract = Join-Path $env:TEMP "nssm_extract"
    try {
        Invoke-WebRequest -Uri "https://nssm.cc/release/nssm-2.24.zip" -OutFile $NssmZip -UseBasicParsing
        Expand-Archive -Path $NssmZip -DestinationPath $NssmExtract -Force
        New-Item -ItemType Directory -Force -Path $NssmDir | Out-Null
        Copy-Item "$NssmExtract\nssm-2.24\win64\nssm.exe" $NssmExe -Force
        Remove-Item $NssmZip, $NssmExtract -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "NSSM ready." -ForegroundColor Green
    }
    catch {
        Write-Error "NSSM download failed: $_"
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# --- Create logs dir FIRST (needed before short-path resolution) ---
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

# --- Resolve 8.3 short paths (no spaces - safe for NSSM) ---
function Get-ShortPath($p) {
    $s = & cmd /c "for %I in (`"$p`") do @echo %~sI" 2>$null
    return $s.Trim()
}
$PythonShort = Get-ShortPath $Python
$RunnerShort = Get-ShortPath $Runner
$DirShort    = Get-ShortPath $ScriptDir
$LogShort    = Get-ShortPath $LogDir

Write-Host "Paths resolved (8.3 short form):"
Write-Host "  Python : $PythonShort"
Write-Host "  Runner : $RunnerShort"
Write-Host "  Dir    : $DirShort"

# --- Force-remove any existing service (NSSM + sc.exe fallback) ---
Write-Host "Checking for existing service '$ServiceName' ..."
$svcQuery = & sc.exe query $ServiceName 2>&1
if ("$svcQuery" -notmatch "does not exist") {
    Write-Host "Existing service found - removing ..." -ForegroundColor Yellow
    & $NssmExe stop   $ServiceName confirm 2>&1 | Out-Null
    Start-Sleep -Seconds 3
    & $NssmExe remove $ServiceName confirm 2>&1 | Out-Null
    Start-Sleep -Seconds 2
    & sc.exe delete $ServiceName 2>&1 | Out-Null
    Start-Sleep -Seconds 3
    Write-Host "Removed." -ForegroundColor Yellow
} else {
    Write-Host "No existing service found." -ForegroundColor Green
}

# --- Install service with short paths ---
Write-Host "Installing service '$ServiceName' ..."
& $NssmExe install $ServiceName $PythonShort $RunnerShort

& $NssmExe set $ServiceName Application   $PythonShort
& $NssmExe set $ServiceName AppParameters $RunnerShort
& $NssmExe set $ServiceName AppDirectory  $DirShort
& $NssmExe set $ServiceName DisplayName   $DisplayName
& $NssmExe set $ServiceName Description   $Description
& $NssmExe set $ServiceName Start         SERVICE_AUTO_START

$StdoutLog = Join-Path $LogShort "AESE_stdout.log"
$StderrLog = Join-Path $LogShort "AESE_stderr.log"
& $NssmExe set $ServiceName AppStdout                    $StdoutLog
& $NssmExe set $ServiceName AppStderr                    $StderrLog
& $NssmExe set $ServiceName AppStdoutCreationDisposition 4
& $NssmExe set $ServiceName AppStderrCreationDisposition 4
& $NssmExe set $ServiceName AppRotateFiles               1
& $NssmExe set $ServiceName AppRotateOnline              1
& $NssmExe set $ServiceName AppRotateBytes               5242880
& $NssmExe set $ServiceName AppRestartDelay              30000
& $NssmExe set $ServiceName AppThrottle                  60000

# --- Run initial full pull before starting the service ---
Write-Host ""
Write-Host "Running initial full pull (this may take 15-20 minutes) ..." -ForegroundColor Cyan
& $PythonShort (Get-ShortPath $Runner) --mode full
if ($LASTEXITCODE -ne 0) {
    Write-Warning "Initial pull exited with code $LASTEXITCODE - check logs. Continuing with service start."
} else {
    Write-Host "Initial full pull completed." -ForegroundColor Green
}

# --- Start service ---
Write-Host "Starting service ..."
& $NssmExe start $ServiceName
Start-Sleep -Seconds 3
$status = & $NssmExe status $ServiceName

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " $ServiceName" -ForegroundColor Green
Write-Host " Status : $status"
Write-Host " Logs   : $LogDir"
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Commands:"
Write-Host "  .\nssm\nssm.exe status  $ServiceName"
Write-Host "  .\nssm\nssm.exe stop    $ServiceName"
Write-Host "  .\nssm\nssm.exe start   $ServiceName"
Write-Host "  .\nssm\nssm.exe remove  $ServiceName confirm"
Write-Host ""
Write-Host "Logs:"
Write-Host "  Get-Content '$LogDir\AESE.log' -Tail 50 -Wait"
Write-Host ""
Write-Host "Schedule: full pull at midnight | incremental every 3 hours"
Read-Host "Press Enter to close"
