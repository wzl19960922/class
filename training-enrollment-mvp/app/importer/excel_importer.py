from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Optional

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


def _normalize_text(value: object) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _normalize_columns(columns: Iterable[str]) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for column in columns:
        normalized = str(column).strip()
        lowered = normalized.lower()
        for key, aliases in COLUMN_ALIASES.items():
            lowered_aliases = {alias.lower() for alias in aliases}
            if normalized in aliases or lowered in lowered_aliases:
                mapping[key] = normalized
    return mapping


def _read_value(row: Dict[str, object], mapping: Dict[str, str], key: str) -> Optional[str]:
    column = mapping.get(key)
    if column is None:
        return None
    return _normalize_text(row.get(column))


def import_enrollments(data_path: Path, session_id: int) -> ImportStats:
    data_path = Path(data_path)
    if not data_path.exists():
        raise FileNotFoundError(data_path)

    rows_by_sheet = _load_rows(data_path)
    total_rows = 0
    imported_rows = 0

    with get_connection() as conn:
        for sheet_name, rows in rows_by_sheet.items():
            if not rows:
                continue
            mapping = _normalize_columns(rows[0].keys())

            for row in rows:
                if not any(_normalize_text(value) for value in row.values()):
                    continue

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


def _load_rows(source_path: Path) -> Dict[str, list[Dict[str, object]]]:
    suffix = source_path.suffix.lower()
    if suffix == ".csv":
        with source_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
            reader = csv.DictReader(csv_file)
            return {"csv": [dict(row) for row in reader]}

    if suffix in {".xlsx", ".xls"}:
        try:
            import pandas as pd
        except ModuleNotFoundError as exc:
            raise ModuleNotFoundError(
                "Importing Excel requires pandas/openpyxl. Please install them or use CSV input."
            ) from exc

        dataframes = pd.read_excel(source_path, sheet_name=None)
        rows_by_sheet: Dict[str, list[Dict[str, object]]] = {}
        for sheet_name, dataframe in dataframes.items():
            if dataframe.empty:
                rows_by_sheet[sheet_name] = []
                continue
            dataframe = dataframe.dropna(axis=0, how="all")
            rows_by_sheet[sheet_name] = dataframe.to_dict(orient="records")
        return rows_by_sheet

    raise ValueError("Only CSV and Excel files (.csv/.xlsx/.xls) are supported.")


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
