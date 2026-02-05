# Training Enrollment MVP

本项目是一个可在本地运行的学员培训管理系统 MVP（Python + SQLite），用于准确统计年度参训人次与参训人数。

## 目录结构

```text
training-enrollment-mvp/
├── app/
│   ├── db/
│   │   ├── database.py
│   │   ├── queries.py
│   │   └── training.db  # 运行后自动创建
│   └── importer/
│       ├── excel_importer.py
│       └── phone.py
├── data/
│   └── sample_enrollments.csv
├── main.py
└── README.md
```

## 环境与安装

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install pandas openpyxl  # 如只导入 CSV，可不安装
```

## 运行最小演示

```bash
python main.py
```

运行后会完成：
1. 初始化 SQLite 数据库 `app/db/training.db`。
2. 手动创建一个 `training_session` 并返回 `session_id`。
3. 导入 `data/sample_enrollments.csv`（CSV 或 Excel；Excel 支持多个工作表）。
4. 打印年度统计结果。

## Excel 导入说明

- CSV 使用标准库 `csv` 读取；Excel 使用 `pandas.read_excel(..., sheet_name=None)` 读取所有工作表。
- 忽略完全空行。
- 每一行 Excel 始终生成一条 `enrollment`。
- 同一手机号在同一/不同培训期可产生多条 `enrollment`，但只对应一个 `person`。
- 手机号为空或明显非法（非 11 位手机号）将抛出异常并停止导入。
- 为避免二进制文件限制，仓库示例数据默认提供为 CSV。

### 导入示例代码

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

## 年度统计函数（可直接调用）

```python
from app.db.queries import (
    count_enrollments_for_year,
    count_unique_people_for_year,
    count_repeat_people_for_year,
    top_learners_for_year,
)

year = 2024
print(count_enrollments_for_year(year))      # 参训人次
print(count_unique_people_for_year(year))    # 参训人数（去重 person）
print(count_repeat_people_for_year(year))    # 重复参训人数（>=2）
print(top_learners_for_year(year, limit=5))  # Top N（手机号 + 姓名 + 次数）
```

## 自检结论

- 同一手机号跨 Excel / 跨培训期只生成一个 `person`（由 `phone_norm` 唯一约束 + 查找创建逻辑保证）。
- 同一手机号在一年内多次报名会生成多条 `enrollment`。
- 年度统计按 `training_session.start_date` 年份聚合，与导入行数逻辑一致。
- `app/db/training.db` 为本地 SQLite 文件，可重复复用。

## 严格限制

- 不使用 Web 框架。
- 不使用 MongoDB / Redis / 云服务。
- 不设计 UI。
- 不引入权限与登录。
- 不进行单位、地区字段的过早标准化（保持原文）。
