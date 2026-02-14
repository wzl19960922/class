import sqlite3
from pathlib import Path
from typing import Iterable, Optional

DB_PATH = Path(__file__).resolve().parent / "training.db"

SCHEMA_STATEMENTS: Iterable[str] = (
    """
    CREATE TABLE IF NOT EXISTS person (
        person_id INTEGER PRIMARY KEY,
        phone_norm TEXT UNIQUE NOT NULL,
        name_latest TEXT,
        org_text_latest TEXT
    );
    """,
    """
    CREATE TABLE IF NOT EXISTS training_session (
        session_id INTEGER PRIMARY KEY,
        title TEXT,
        start_date TEXT,
        end_date TEXT,
        location_text TEXT,
        created_at TEXT
    );
    """,
    """
    CREATE TABLE IF NOT EXISTS enrollment (
        enrollment_id INTEGER PRIMARY KEY,
        session_id INTEGER NOT NULL,
        person_id INTEGER NOT NULL,
        enrolled_at TEXT,
        name_snapshot TEXT,
        org_text TEXT,
        region_text TEXT,
        room_preference TEXT,
        remote_id_snapshot TEXT,
        FOREIGN KEY (session_id) REFERENCES training_session(session_id),
        FOREIGN KEY (person_id) REFERENCES person(person_id)
    );
    """,
)


def get_connection(db_path: Optional[Path] = None) -> sqlite3.Connection:
    target = db_path or DB_PATH
    conn = sqlite3.connect(target)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def initialize_database(db_path: Optional[Path] = None) -> Path:
    target = db_path or DB_PATH
    target.parent.mkdir(parents=True, exist_ok=True)
    conn = get_connection(target)
    try:
        for statement in SCHEMA_STATEMENTS:
            conn.execute(statement)
        conn.commit()
    finally:
        conn.close()
    return target
