from pathlib import Path

from app.db.database import initialize_database
from app.db.queries import (
    count_enrollments_for_year,
    count_repeat_people_for_year,
    count_unique_people_for_year,
    create_training_session,
    top_learners_for_year,
)
from app.importer.excel_importer import import_enrollments


def main() -> None:
    initialize_database()

    session_id = create_training_session(
        title="年度培训示例",
        start_date="2024-05-01",
        end_date="2024-05-03",
        location_text="北京",
    )

    excel_path = Path("data") / "sample_enrollments.csv"
    stats = import_enrollments(excel_path, session_id=session_id)

    year = 2024
    print(f"导入总行数: {stats.total_rows}")
    print(f"成功导入行数: {stats.imported_rows}")
    print(f"{year}年参训人次: {count_enrollments_for_year(year)}")
    print(f"{year}年参训人数: {count_unique_people_for_year(year)}")
    print(f"{year}年重复参训人数: {count_repeat_people_for_year(year)}")

    print(f"{year}年参训次数 Top 3 学员:")
    for learner in top_learners_for_year(year, limit=3):
        print(f"- {learner.phone_norm} {learner.name_latest} ({learner.enrollments}次)")


if __name__ == "__main__":
    main()
