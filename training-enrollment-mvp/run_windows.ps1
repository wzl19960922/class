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

Write-Host "Using virtual environment..." -ForegroundColor Yellow
$venvPython = Join-Path -Path .venv -ChildPath "Scripts/python.exe"
if (-not (Test-Path $venvPython)) {
  Write-Host "Virtual environment python not found. Try deleting .venv and re-running." -ForegroundColor Red
  exit 1
}

Write-Host "Installing dependencies..." -ForegroundColor Yellow
& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install pandas openpyxl

Write-Host "Running application..." -ForegroundColor Green
& $venvPython -m app.cli
