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
2. 安装依赖 `flask pandas openpyxl python-docx qrcode[pil]`；
3. 运行 `env_check.py`；
4. 启动服务 `http://127.0.0.1:5000`。

## 手动运行

```bash
python env_check.py
python main.py
```

## 常见启动报错排查

### 1) `View function mapping is overwriting an existing endpoint function: index`
### 2) `View function mapping is overwriting an existing endpoint function: import_course_word`

这两个报错通常表示：你本地 `main.py` 不是最新版本（常见于冲突后重复粘贴、旧文件未覆盖）。

建议按下面顺序处理（Windows PowerShell）：

```powershell
cd D:\survey\class\training-enrollment-mvp
git status
git fetch --all
git reset --hard origin/<你的分支名>
python -m py_compile main.py
```

如果你不方便 `reset --hard`，至少先对比以下关键行是否存在：

- `@app.route("/", endpoint="home_page")`
- `@app.route("/api/course/import", methods=["POST"], endpoint="api_course_import")`

若不存在，说明仍是旧版本文件，请先同步代码再运行。

## 主要文件

- `env_check.py`：环境自检脚本
- `main.py`：Flask 后端与 API
- `templates/index.html`：单页前端
- `static/app.js`：前端交互逻辑
- `run_windows.ps1` / `run_windows.cmd`：Windows 一键启动脚本
