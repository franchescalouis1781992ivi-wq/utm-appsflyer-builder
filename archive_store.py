import json
import sqlite3
from pathlib import Path


DB_PATH = Path("data/link_builder_archive.db")


def get_connection() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_archive_db() -> None:
    with get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS campaign_archives (
                archive_id INTEGER PRIMARY KEY AUTOINCREMENT,
                archive_name TEXT NOT NULL,
                utm_campaign TEXT,
                theme TEXT,
                sender TEXT,
                campaign_date TEXT,
                utm_term_base TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                metadata_json TEXT NOT NULL,
                rows_json TEXT NOT NULL,
                result_json TEXT NOT NULL
            )
            """
        )


def _dump_json(payload: dict | list) -> str:
    return json.dumps(payload, ensure_ascii=False)


def _load_json(payload: str) -> dict | list:
    return json.loads(payload) if payload else {}


def create_archive(
    archive_name: str,
    utm_campaign: str,
    theme: str,
    sender: str,
    campaign_date: str,
    utm_term_base: str,
    created_at: str,
    updated_at: str,
    metadata: dict,
    rows: list[dict],
    result: dict,
) -> int:
    with get_connection() as conn:
        cursor = conn.execute(
            """
            INSERT INTO campaign_archives (
                archive_name,
                utm_campaign,
                theme,
                sender,
                campaign_date,
                utm_term_base,
                created_at,
                updated_at,
                metadata_json,
                rows_json,
                result_json
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                archive_name,
                utm_campaign,
                theme,
                sender,
                campaign_date,
                utm_term_base,
                created_at,
                updated_at,
                _dump_json(metadata),
                _dump_json(rows),
                _dump_json(result),
            ),
        )
        return int(cursor.lastrowid)


def update_archive(
    archive_id: int,
    archive_name: str,
    utm_campaign: str,
    theme: str,
    sender: str,
    campaign_date: str,
    utm_term_base: str,
    updated_at: str,
    metadata: dict,
    rows: list[dict],
    result: dict,
) -> None:
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE campaign_archives
            SET archive_name = ?,
                utm_campaign = ?,
                theme = ?,
                sender = ?,
                campaign_date = ?,
                utm_term_base = ?,
                updated_at = ?,
                metadata_json = ?,
                rows_json = ?,
                result_json = ?
            WHERE archive_id = ?
            """,
            (
                archive_name,
                utm_campaign,
                theme,
                sender,
                campaign_date,
                utm_term_base,
                updated_at,
                _dump_json(metadata),
                _dump_json(rows),
                _dump_json(result),
                archive_id,
            ),
        )


def delete_archive(archive_id: int) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM campaign_archives WHERE archive_id = ?", (archive_id,))


def get_archive(archive_id: int) -> dict | None:
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT archive_id,
                   archive_name,
                   utm_campaign,
                   theme,
                   sender,
                   campaign_date,
                   utm_term_base,
                   created_at,
                   updated_at,
                   metadata_json,
                   rows_json,
                   result_json
            FROM campaign_archives
            WHERE archive_id = ?
            """,
            (archive_id,),
        ).fetchone()

    if row is None:
        return None

    payload = dict(row)
    payload["metadata"] = _load_json(payload.pop("metadata_json"))
    payload["rows"] = _load_json(payload.pop("rows_json"))
    payload["result"] = _load_json(payload.pop("result_json"))
    return payload


def list_archives(search: str = "") -> list[dict]:
    pattern = f"%{search.strip()}%"
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT archive_id,
                   archive_name,
                   utm_campaign,
                   theme,
                   sender,
                   campaign_date,
                   utm_term_base,
                   created_at,
                   updated_at
            FROM campaign_archives
            WHERE ? = ''
               OR archive_name LIKE ?
               OR utm_campaign LIKE ?
               OR theme LIKE ?
               OR sender LIKE ?
               OR campaign_date LIKE ?
               OR utm_term_base LIKE ?
            ORDER BY updated_at DESC, archive_id DESC
            """,
            (search.strip(), pattern, pattern, pattern, pattern, pattern, pattern),
        ).fetchall()

    return [dict(row) for row in rows]
