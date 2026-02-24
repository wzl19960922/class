$ErrorActionPreference = 'Stop'

Write-Host "== Training Enrollment MVP (Windows One-Click) ==" -ForegroundColor Cyan

function Ensure-Conda {
    if ($env:CONDA_EXE -and (Test-Path $env:CONDA_EXE)) {
        return $env:CONDA_EXE
    }

    $condaCmd = Get-Command conda -ErrorAction SilentlyContinue
    if ($condaCmd -and $condaCmd.Source) {
        return $condaCmd.Source
    }

    $candidates = @(
        "D:\conda\Scripts\conda.exe",
        "D:\anaconda3\Scripts\conda.exe",
        "D:\miniconda3\Scripts\conda.exe",
        "$env:USERPROFILE\anaconda3\Scripts\conda.exe",
        "$env:USERPROFILE\miniconda3\Scripts\conda.exe"
    )
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }
    return $null
}

function Ensure-EnvWithConda {
    param(
        [string]$CondaExe,
        [string]$EnvName
    )

    Write-Host "检测到 Conda，准备环境: $EnvName" -ForegroundColor Yellow

    $envExists = (& $CondaExe env list | Select-String -Pattern "^$EnvName\s|[\\/]$EnvName$")
    if (-not $envExists) {
        Write-Host "创建 Conda 环境 $EnvName ..." -ForegroundColor Yellow
        & $CondaExe create -n $EnvName python=3.11 -y
    }

    Write-Host "安装依赖..." -ForegroundColor Yellow
    & $CondaExe run -n $EnvName python -m pip install --upgrade pip
    & $CondaExe run -n $EnvName python -m pip install flask pandas openpyxl python-docx qrcode[pil]

    Write-Host "当前将使用以下解释器启动：" -ForegroundColor Yellow
    & $CondaExe run -n $EnvName python -c "import sys; print(sys.executable)"

    Write-Host "运行环境检查..." -ForegroundColor Yellow
    & $CondaExe run -n $EnvName python env_check.py

    Write-Host "启动服务: http://127.0.0.1:5000" -ForegroundColor Green
    & $CondaExe run -n $EnvName python main.py
}

function Ensure-EnvWithVenv {
    Write-Host "未检测到 Conda，回退到 .venv 方案" -ForegroundColor Yellow

    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCmd) {
        $pythonCmd = Get-Command py -ErrorAction SilentlyContinue
    }
    if (-not $pythonCmd) {
        Write-Host "未找到 Python。请先安装 Python 3.10+ 或 Anaconda。" -ForegroundColor Red
        exit 1
    }

    if (-not (Test-Path .venv)) {
        Write-Host "创建 .venv ..." -ForegroundColor Yellow
        if ($pythonCmd.Name -ieq 'py.exe' -or $pythonCmd.Name -ieq 'py') {
            & $pythonCmd.Source -3 -m venv .venv
        } else {
            & $pythonCmd.Source -m venv .venv
        }
    }

    $venvPython = Join-Path (Get-Location) '.venv\Scripts\python.exe'
    if (-not (Test-Path $venvPython)) {
        Write-Host "未找到 .venv Python，可删除 .venv 后重试。" -ForegroundColor Red
        exit 1
    }

    Write-Host "安装依赖..." -ForegroundColor Yellow
    & $venvPython -m pip install --upgrade pip
    & $venvPython -m pip install flask pandas openpyxl python-docx qrcode[pil]

    Write-Host "运行环境检查..." -ForegroundColor Yellow
    & $venvPython env_check.py

    Write-Host "启动服务: http://127.0.0.1:5000" -ForegroundColor Green
    & $venvPython main.py
}

$condaExe = Ensure-Conda
if ($condaExe) {
    Ensure-EnvWithConda -CondaExe $condaExe -EnvName 'training-mvp'
} else {
    Ensure-EnvWithVenv
}
