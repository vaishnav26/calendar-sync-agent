import os
import sqlite3
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


DB_PATH = Path(os.environ.get("SYNC_TRACKER_DB", "/tmp/calendar_agent_sync_tracker.db"))
SYNC_TZ = ZoneInfo(os.environ.get("SYNC_LIMIT_TIMEZONE", "Asia/Kolkata"))


def today_key() -> str:
    return datetime.now(SYNC_TZ).strftime("%Y-%m-%d")


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH, timeout=10) as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS sync_log (
                user_key TEXT NOT NULL,
                sync_date TEXT NOT NULL,
                count INTEGER NOT NULL DEFAULT 0,
                PRIMARY KEY (user_key, sync_date)
            )
            """
        )


def check_and_increment_sync(user_key: str, daily_limit: int = 5) -> tuple[bool, int]:
    if not user_key:
        raise ValueError("Missing signed-in user identity.")

    init_db()
    sync_date = today_key()
    with sqlite3.connect(DB_PATH, timeout=10, isolation_level=None) as conn:
        conn.execute("BEGIN IMMEDIATE")
        try:
            row = conn.execute(
                "SELECT count FROM sync_log WHERE user_key = ? AND sync_date = ?",
                (user_key, sync_date),
            ).fetchone()
            current_count = int(row[0]) if row else 0
            if current_count >= daily_limit:
                conn.execute("COMMIT")
                return False, current_count

            new_count = current_count + 1
            conn.execute(
                """
                INSERT INTO sync_log (user_key, sync_date, count)
                VALUES (?, ?, ?)
                ON CONFLICT(user_key, sync_date)
                DO UPDATE SET count = excluded.count
                """,
                (user_key, sync_date, new_count),
            )
            conn.execute("COMMIT")
            return True, new_count
        except Exception:
            conn.execute("ROLLBACK")
            raise


def get_sync_count(user_key: str) -> int:
    if not user_key:
        return 0
    init_db()
    with sqlite3.connect(DB_PATH, timeout=10) as conn:
        row = conn.execute(
            "SELECT count FROM sync_log WHERE user_key = ? AND sync_date = ?",
            (user_key, today_key()),
        ).fetchone()
    return int(row[0]) if row else 0
