# Training Enrollment MVP

本项目为学员培训管理系统 MVP，支持 Excel 报名表导入与年度统计。仅使用 Python + SQLite。

## 目录结构

```
app/
  db/
  importer/
data/
README.md
main.py
```

## 快速开始

### Windows（PowerShell，一键）

在项目根目录执行：

```powershell
.\run_windows.ps1
```

> 说明：脚本会自动创建虚拟环境、安装依赖并运行 `main.py`。

### 通用（Linux/macOS）

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install pandas openpyxl
python main.py
```

运行后将创建数据库并导入 `data/sample_enrollments.csv`（同样支持 Excel），输出年度统计结果。

## Excel 导入示例

```python
from pathlib import Path

from app.db.database import initialize_database
from app.db.queries import create_training_session
from app.importer.excel_importer import import_enrollments

initialize_database()

session_id = create_training_session(
    title="季度培训",
    start_date="2024-06-01",
    end_date="2024-06-03",
    location_text="上海",
)

stats = import_enrollments(Path("data/sample_enrollments.csv"), session_id=session_id)
print(stats)
```

## 年度统计调用示例

```python
from app.db.queries import (
    count_enrollments_for_year,
    count_unique_people_for_year,
    count_repeat_people_for_year,
    top_learners_for_year,
)

year = 2024
print(count_enrollments_for_year(year))
print(count_unique_people_for_year(year))
print(count_repeat_people_for_year(year))
print(top_learners_for_year(year, limit=5))
```

## 约束说明

- 同一手机号在不同 Excel 或培训期只生成一个 `Person`。
- 每一行 Excel 都会生成一条 `Enrollment` 记录，不跨期去重。
- 年度统计基于 `training_session.start_date` 的年份。
- SQLite 数据库文件位于 `app/db/training.db`，可重复使用。
