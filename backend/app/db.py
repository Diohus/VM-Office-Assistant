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
]


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
    SELECT d.document_id, d.name, v.version, v.sha256, v.size, v.meta, v.created_at
    FROM document d JOIN document_version v USING (document_id)
"""


def create_document(name: str, sha: str, size: int, meta: dict) -> dict:
    document_id = uuid.uuid4()
    with _connect() as conn:
        conn.execute("INSERT INTO document (document_id, name) VALUES (%s, %s)", (document_id, name))
        conn.execute("INSERT INTO document_version (document_id, version, sha256, size, meta) VALUES (%s, 1, %s, %s, %s)",
                     (document_id, sha, size, json.dumps(meta, ensure_ascii=False)))
        return conn.execute(_SELECT + " WHERE d.document_id = %s AND v.version = 1", (document_id,)).fetchone()


def latest(document_id: uuid.UUID) -> dict | None:
    with _connect() as conn:
        return conn.execute(_SELECT + " WHERE d.document_id = %s ORDER BY v.version DESC LIMIT 1", (document_id,)).fetchone()


def recent(limit: int = 20) -> list[dict]:
    with _connect() as conn:
        return conn.execute("SELECT document_id, name, created_at FROM document ORDER BY created_at DESC LIMIT %s", (limit,)).fetchall()
