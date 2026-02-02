from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Optional

import pandas as pd

from app.db.database import get_connection
from app.importer.phone import PhoneNormalizationError, normalize_phone


COLUMN_ALIASES = {
    "name": {"姓名", "名字", "name"},
    "phone": {"手机", "手机号", "联系电话", "phone", "mobile"},
    "org": {"单位", "公司", "组织", "org", "organization"},
    "region": {"地区", "区域", "region", "province"},
    "remote_id": {"远程网id", "远程网ID", "远程id", "remote id", "remote_id"},
    "room": {"住宿偏好", "住宿", "room", "room preference"},
}


@dataclass
class ImportStats:
    total_rows: int
    imported_rows: int


def _normalize_columns(columns: Iterable[str]) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for column in columns:
        normalized = str(column).strip()
        lowered = normalized.lower()
        for key, aliases in COLUMN_ALIASES.items():
            if normalized in aliases or lowered in aliases:
                mapping[key] = column
    return mapping


def _read_value(row: pd.Series, mapping: Dict[str, str], key: str) -> Optional[str]:
    column = mapping.get(key)
    if column is None:
        return None
    value = row.get(column)
    if pd.isna(value):
        return None
    return str(value).strip()


def import_enrollments(excel_path: Path, session_id: int) -> ImportStats:
    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(excel_path)

    dataframes = _load_dataframes(excel_path)
    total_rows = 0
    imported_rows = 0

    with get_connection() as conn:
        for sheet_name, dataframe in dataframes.items():
            if dataframe.empty:
                continue
            dataframe = dataframe.dropna(axis=0, how="all")
            if dataframe.empty:
                continue
            mapping = _normalize_columns(dataframe.columns)

            for _, row in dataframe.iterrows():
                total_rows += 1
                phone_raw = _read_value(row, mapping, "phone")
                try:
                    phone_norm = normalize_phone(phone_raw)
                except PhoneNormalizationError as exc:
                    raise PhoneNormalizationError(
                        f"Sheet '{sheet_name}' has invalid phone: {phone_raw}"
                    ) from exc

                name = _read_value(row, mapping, "name")
                org_text = _read_value(row, mapping, "org")
                region_text = _read_value(row, mapping, "region")
                remote_id = _read_value(row, mapping, "remote_id")
                room_preference = _read_value(row, mapping, "room")

                person_id = _get_or_create_person(conn, phone_norm, name, org_text)
                conn.execute(
                    """
                    INSERT INTO enrollment (
                        session_id,
                        person_id,
                        enrolled_at,
                        name_snapshot,
                        org_text,
                        region_text,
                        room_preference,
                        remote_id_snapshot
                    ) VALUES (?, ?, datetime('now'), ?, ?, ?, ?, ?)
                    """,
                    (
                        session_id,
                        person_id,
                        name,
                        org_text,
                        region_text,
                        room_preference,
                        remote_id,
                    ),
                )
                imported_rows += 1
        conn.commit()

    return ImportStats(total_rows=total_rows, imported_rows=imported_rows)


def _load_dataframes(source_path: Path) -> Dict[str, pd.DataFrame]:
    suffix = source_path.suffix.lower()
    if suffix == ".csv":
        dataframe = pd.read_csv(source_path)
        return {"csv": dataframe}
    return pd.read_excel(source_path, sheet_name=None)


def _get_or_create_person(
    conn, phone_norm: str, name: Optional[str], org_text: Optional[str]
) -> int:
    existing = conn.execute(
        "SELECT person_id FROM person WHERE phone_norm = ?", (phone_norm,)
    ).fetchone()
    if existing:
        conn.execute(
            """
            UPDATE person
            SET name_latest = COALESCE(?, name_latest),
                org_text_latest = COALESCE(?, org_text_latest)
            WHERE person_id = ?
            """,
            (name, org_text, existing["person_id"]),
        )
        return int(existing["person_id"])

    cursor = conn.execute(
        """
        INSERT INTO person (phone_norm, name_latest, org_text_latest)
        VALUES (?, ?, ?)
        """,
        (phone_norm, name, org_text),
    )
    return int(cursor.lastrowid)
