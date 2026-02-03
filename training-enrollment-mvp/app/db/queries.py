from __future__ import annotations

from dataclasses import dataclass
from typing import List, Tuple

from app.db.database import get_connection


@dataclass
class TopLearner:
    phone_norm: str
    name_latest: str | None
    enrollments: int


def create_training_session(
    title: str,
    start_date: str,
    end_date: str,
    location_text: str,
) -> int:
    with get_connection() as conn:
        cursor = conn.execute(
            """
            INSERT INTO training_session (title, start_date, end_date, location_text, created_at)
            VALUES (?, ?, ?, ?, datetime('now'))
            """,
            (title, start_date, end_date, location_text),
        )
        conn.commit()
        return int(cursor.lastrowid)


def count_enrollments_for_year(year: int) -> int:
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT COUNT(*) AS total
            FROM enrollment
            JOIN training_session ON enrollment.session_id = training_session.session_id
            WHERE strftime('%Y', training_session.start_date) = ?
            """,
            (str(year),),
        ).fetchone()
        return int(row["total"])


def count_unique_people_for_year(year: int) -> int:
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT COUNT(DISTINCT enrollment.person_id) AS total
            FROM enrollment
            JOIN training_session ON enrollment.session_id = training_session.session_id
            WHERE strftime('%Y', training_session.start_date) = ?
            """,
            (str(year),),
        ).fetchone()
        return int(row["total"])


def count_repeat_people_for_year(year: int) -> int:
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT COUNT(*) AS total
            FROM (
                SELECT enrollment.person_id
                FROM enrollment
                JOIN training_session ON enrollment.session_id = training_session.session_id
                WHERE strftime('%Y', training_session.start_date) = ?
                GROUP BY enrollment.person_id
                HAVING COUNT(*) >= 2
            ) AS repeaters
            """,
            (str(year),),
        ).fetchone()
        return int(row["total"])


def top_learners_for_year(year: int, limit: int = 5) -> List[TopLearner]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT person.phone_norm, person.name_latest, COUNT(*) AS enrollments
            FROM enrollment
            JOIN person ON enrollment.person_id = person.person_id
            JOIN training_session ON enrollment.session_id = training_session.session_id
            WHERE strftime('%Y', training_session.start_date) = ?
            GROUP BY person.person_id
            ORDER BY enrollments DESC, person.phone_norm ASC
            LIMIT ?
            """,
            (str(year), limit),
        ).fetchall()
        return [
            TopLearner(
                phone_norm=row["phone_norm"],
                name_latest=row["name_latest"],
                enrollments=row["enrollments"],
            )
            for row in rows
        ]
