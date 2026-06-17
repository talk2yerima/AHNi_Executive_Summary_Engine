# =============================================================================
# setup.ps1  —  First-time setup for ACEBAY DHIS2 pipeline
# Run from the project root:
#     Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#     .\setup.ps1
# =============================================================================

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  ACEBAY DHIS2 Pipeline  —  Setup"           -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# ── 1. Check Python ──────────────────────────────────────────────────────────
Write-Host "[1/4] Checking Python ..." -ForegroundColor Yellow
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "  ERROR: Python not found." -ForegroundColor Red
    Write-Host "  Download from https://www.python.org/downloads/" -ForegroundColor Red
    Write-Host "  Check 'Add Python to PATH' during install." -ForegroundColor Red
    exit 1
}
$pyVer = python --version
Write-Host "  OK — $pyVer" -ForegroundColor Green

# ── 2. Create virtual environment ────────────────────────────────────────────
Write-Host ""
Write-Host "[2/4] Setting up virtual environment ..." -ForegroundColor Yellow
if (-not (Test-Path ".venv")) {
    python -m venv .venv
    Write-Host "  OK — .venv created" -ForegroundColor Green
} else {
    Write-Host "  OK — .venv already exists" -ForegroundColor Green
}

# ── 3. Install packages ──────────────────────────────────────────────────────
Write-Host ""
Write-Host "[3/4] Installing packages ..." -ForegroundColor Yellow
& .\.venv\Scripts\Activate.ps1
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Host "  ERROR: pip install failed." -ForegroundColor Red
    exit 1
}
Write-Host "  OK — packages installed" -ForegroundColor Green

# ── 4. Folders and .env ──────────────────────────────────────────────────────
Write-Host ""
Write-Host "[4/4] Preparing folders and config ..." -ForegroundColor Yellow
New-Item -ItemType Directory -Force -Path "output" | Out-Null
New-Item -ItemType Directory -Force -Path "cache"  | Out-Null
Write-Host "  OK — output\  and  cache\  ready" -ForegroundColor Green

if (-not (Test-Path ".env")) {
    Copy-Item ".env.example" ".env"
    Write-Host "  OK — .env created from .env.example" -ForegroundColor Green
    Write-Host ""
    Write-Host "  *** ACTION REQUIRED ***" -ForegroundColor Yellow
    Write-Host "  Open .env and fill in your credentials:" -ForegroundColor Yellow
    Write-Host "    DHIS2_URL, DHIS2_USER, DHIS2_PASS" -ForegroundColor Gray
    Write-Host "    AZURE_CONNECTION_STRING, AZURE_CONTAINER_NAME" -ForegroundColor Gray
} else {
    Write-Host "  OK — .env already exists" -ForegroundColor Green
}

# ── Done ─────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Setup complete!"                            -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Run the pipeline:" -ForegroundColor White
Write-Host "  Full pull  :  .\.venv\Scripts\python dhis2_pull.py --mode full"        -ForegroundColor Gray
Write-Host "  Incremental:  .\.venv\Scripts\python dhis2_pull.py --mode incremental" -ForegroundColor Gray
Write-Host ""
