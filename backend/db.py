from __future__ import annotations

from contextlib import contextmanager

from psycopg.rows import dict_row

from db_config import connect, get_database_settings, load_project_env


load_project_env()


@contextmanager
def get_connection():
    conn = connect(row_factory=dict_row)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def ping_database() -> dict[str, str]:
    settings = get_database_settings()
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("select current_database() as database_name, current_user as user_name")
            row = cur.fetchone()
    return {
        "host": settings["host"],
        "database": row["database_name"],
        "user": row["user_name"],
    }
