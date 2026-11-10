from __future__ import annotations

import json
import time
import uuid

import psycopg
from psycopg.rows import dict_row

from .config import DATABASE_URL

MIGRATIONS = [
    """
    CREATE TABLE document (
        document_id uuid PRIMARY KEY,
        name text NOT NULL,
        created_at timestamptz NOT NULL DEFAULT now()
    );
    CREATE TABLE document_version (
        document_id uuid NOT NULL REFERENCES document(document_id),
        version integer NOT NULL,
        sha256 char(64) NOT NULL,
        size bigint NOT NULL,
        meta jsonb NOT NULL,
        created_at timestamptz NOT NULL DEFAULT now(),
        PRIMARY KEY (document_id, version)
    );
    """,
    # Ai đã đổi gì. Ô người dùng tự gõ là "nhập tay" (D9), cần biết ai và lúc nào.
    """
    ALTER TABLE document_version ADD COLUMN author text NOT NULL DEFAULT '';
    ALTER TABLE document_version ADD COLUMN summary text NOT NULL DEFAULT '';
    ALTER TABLE document_version ADD COLUMN restored_from integer;
    CREATE TABLE cell_edit (
        document_id uuid NOT NULL,
        version integer NOT NULL,
        sheet text NOT NULL,
        cell text NOT NULL,
        value jsonb,
        formula text,
        author text NOT NULL,
        created_at timestamptz NOT NULL DEFAULT now(),
        PRIMARY KEY (document_id, version, sheet, cell),
        FOREIGN KEY (document_id, version) REFERENCES document_version(document_id, version)
    );
    """,
]


class StaleVersion(Exception):
    """Bản người dùng đang sửa không còn là bản mới nhất."""


def _connect():
    return psycopg.connect(DATABASE_URL, row_factory=dict_row)


def migrate(wait_seconds: int = 60) -> None:
    deadline = time.monotonic() + wait_seconds
    while True:
        try:
            conn = _connect()
            break
        except psycopg.OperationalError:
            if time.monotonic() > deadline:
                raise
            time.sleep(1)
    with conn, conn.cursor() as cur:
        cur.execute("SELECT pg_advisory_xact_lock(7342001)")
        cur.execute("CREATE TABLE IF NOT EXISTS schema_migration (n integer PRIMARY KEY)")
        cur.execute("SELECT coalesce(max(n), 0) AS n FROM schema_migration")
        done = cur.fetchone()["n"]
        for n, sql in enumerate(MIGRATIONS, 1):
            if n > done:
                cur.execute(sql)
                cur.execute("INSERT INTO schema_migration (n) VALUES (%s)", (n,))


def ping() -> None:
    with _connect() as conn:
        conn.execute("SELECT 1")


_SELECT = """
    SELECT d.document_id, d.name, v.version, v.sha256, v.size, v.meta, v.created_at, v.author, v.summary, v.restored_from
    FROM document d JOIN document_version v USING (document_id)
"""


def create_document(name: str, sha: str, size: int, meta: dict, author: str = "") -> dict:
    document_id = uuid.uuid4()
    with _connect() as conn:
        conn.execute("INSERT INTO document (document_id, name) VALUES (%s, %s)", (document_id, name))
        conn.execute("INSERT INTO document_version (document_id, version, sha256, size, meta, author, summary) VALUES (%s, 1, %s, %s, %s, %s, %s)",
                     (document_id, sha, size, json.dumps(meta, ensure_ascii=False), author, "Tải lên"))
        return conn.execute(_SELECT + " WHERE d.document_id = %s AND v.version = 1", (document_id,)).fetchone()


def add_version(document_id: uuid.UUID, base_version: int, sha: str, size: int, meta: dict, author: str, summary: str,
                edits: list[dict] = (), restored_from: int | None = None) -> dict:
    """Thêm bản mới và bản ghi từng ô trong cùng một giao dịch: lỗi giữa chừng thì không có bản dở dang."""
    with _connect() as conn:
        # Khóa dòng tài liệu để hai lần lưu cùng lúc không cùng nhận một số bản.
        conn.execute("SELECT 1 FROM document WHERE document_id = %s FOR UPDATE", (document_id,))
        latest_n = conn.execute("SELECT max(version) AS n FROM document_version WHERE document_id = %s", (document_id,)).fetchone()["n"]
        if latest_n != base_version:
            raise StaleVersion
        version = latest_n + 1
        conn.execute("INSERT INTO document_version (document_id, version, sha256, size, meta, author, summary, restored_from) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
                     (document_id, version, sha, size, json.dumps(meta, ensure_ascii=False), author, summary, restored_from))
        for e in edits:
            conn.execute("INSERT INTO cell_edit (document_id, version, sheet, cell, value, formula, author) VALUES (%s, %s, %s, %s, %s, %s, %s)",
                         (document_id, version, e["sheet"], e["cell"], json.dumps(e.get("value")), e.get("formula"), author))
        return conn.execute(_SELECT + " WHERE d.document_id = %s AND v.version = %s", (document_id, version)).fetchone()


def latest(document_id: uuid.UUID) -> dict | None:
    with _connect() as conn:
        return conn.execute(_SELECT + " WHERE d.document_id = %s ORDER BY v.version DESC LIMIT 1", (document_id,)).fetchone()


def version(document_id: uuid.UUID, n: int) -> dict | None:
    with _connect() as conn:
        return conn.execute(_SELECT + " WHERE d.document_id = %s AND v.version = %s", (document_id, n)).fetchone()


def versions(document_id: uuid.UUID) -> list[dict]:
    with _connect() as conn:
        return conn.execute(_SELECT + " WHERE d.document_id = %s ORDER BY v.version DESC", (document_id,)).fetchall()


def recent(limit: int = 20) -> list[dict]:
    with _connect() as conn:
        return conn.execute(
            """
            SELECT d.document_id, d.name, v.version, v.created_at, v.author
            FROM document d
            JOIN LATERAL (SELECT version, created_at, author FROM document_version WHERE document_id = d.document_id ORDER BY version DESC LIMIT 1) v ON true
            ORDER BY v.created_at DESC LIMIT %s
            """,
            (limit,),
        ).fetchall()
