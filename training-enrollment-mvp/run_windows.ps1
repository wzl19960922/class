$ErrorActionPreference = 'Stop'

Write-Host "== Training Enrollment MVP (Windows) ==" -ForegroundColor Cyan

$py = Get-Command python -ErrorAction SilentlyContinue
if (-not $py) {
  $py = Get-Command py -ErrorAction SilentlyContinue
}
if (-not $py) {
  Write-Host "Python not found. Please install Python 3 and ensure it is on PATH." -ForegroundColor Red
  exit 1
}

if (-not (Test-Path .venv)) {
  Write-Host "Creating virtual environment..." -ForegroundColor Yellow
  & $py.Source -m venv .venv
}

Write-Host "Activating virtual environment..." -ForegroundColor Yellow
if (Test-Path .\.venv\Scripts\Activate.ps1) {
  . .\.venv\Scripts\Activate.ps1
} elseif (Test-Path .\.venv\Scripts\activate.bat) {
  Write-Host "Activate.ps1 not found, using activate.bat fallback." -ForegroundColor Yellow
  cmd /c .\.venv\Scripts\activate.bat ^&^& python -m pip install --upgrade pip ^&^& python -m pip install pandas openpyxl ^&^& python main.py
  exit $LASTEXITCODE
} else {
  Write-Host "Virtual environment activation script not found. Try deleting .venv and re-running." -ForegroundColor Red
  exit 1
}

Write-Host "Installing dependencies..." -ForegroundColor Yellow
python -m pip install --upgrade pip
python -m pip install pandas openpyxl

Write-Host "Running application..." -ForegroundColor Green
python main.py
