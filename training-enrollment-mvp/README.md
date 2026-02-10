# Training Enrollment MVP

本项目是本地单机运行的学员培训管理系统 MVP（Flask + SQLite）。

## Windows 一键运行（推荐）

### CMD
```bat
run_windows.cmd
```

### PowerShell
```powershell
.\run_windows.ps1
```

脚本会自动：
1. 优先使用 Conda（环境名 `training-mvp`），若无 Conda 则回退 `.venv`；
2. 安装依赖 `flask pandas openpyxl`；
3. 运行 `env_check.py`；
4. 启动服务 `http://127.0.0.1:5000`。

## 手动运行

```bash
python env_check.py
python main.py
```

## 主要文件

- `env_check.py`：环境自检脚本
- `main.py`：Flask 后端与 API
- `templates/index.html`：单页前端
- `static/app.js`：前端交互逻辑
- `run_windows.ps1` / `run_windows.cmd`：Windows 一键启动脚本
