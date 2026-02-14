$ErrorActionPreference = 'Stop'

Write-Host "== Repair Training Enrollment MVP entry files ==" -ForegroundColor Cyan

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  Write-Host "git not found. Please install Git first." -ForegroundColor Red
  exit 1
}

if (-not (Test-Path .git)) {
  Write-Host "Current directory is not a git repository root." -ForegroundColor Red
  exit 1
}

Write-Host "Restoring tracked entry files from git..." -ForegroundColor Yellow
git restore main.py app/cli.py run_windows.ps1 README.md

Write-Host "Pulling latest changes..." -ForegroundColor Yellow
git pull

Write-Host "Repair done. Run project with:" -ForegroundColor Green
Write-Host "  python -m app.cli" -ForegroundColor Green
Write-Host "or:" -ForegroundColor Green
Write-Host "  .\\run_windows.ps1" -ForegroundColor Green
