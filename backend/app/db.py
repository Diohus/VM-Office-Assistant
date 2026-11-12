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
    # Lát 1 nối AI (Q133). File nguồn chỉ lưu SHA và tên, nội dung nằm trong kho; không lưu lưới đã đọc của nguồn.
    # Lô: draft (vừa tạo, chưa mở, không tính D10), open (đang duyệt hoặc còn ô chờ), done. `targets`: change_id -> "sheet!A1".
    # `decisions`: change_id -> {state: pending|rejected|written|blank, how, version}.
    """
    CREATE TABLE source_file (
        source_id uuid PRIMARY KEY,
        document_id uuid NOT NULL REFERENCES document(document_id),
        name text NOT NULL,
        sha256 char(64) NOT NULL,
        size bigint NOT NULL,
        created_at timestamptz NOT NULL DEFAULT now()
    );
    CREATE TABLE fill_batch (
        batch_id uuid PRIMARY KEY,
        document_id uuid NOT NULL REFERENCES document(document_id),
        source_id uuid NOT NULL REFERENCES source_file(source_id),
        base_version integer NOT NULL,
        status text NOT NULL CHECK (status IN ('draft', 'open', 'done')),
        batch jsonb NOT NULL,
        targets jsonb NOT NULL,
        decisions jsonb NOT NULL,
        created_at timestamptz NOT NULL DEFAULT now(),
        opened_at timestamptz
    );
    CREATE INDEX fill_batch_document ON fill_batch (document_id, status);
    """,
    # Lát 2 khung chat (Q135). chat_call: mỗi lượt gọi model một dòng, KHÔNG có câu gõ, giá trị ô, payload hay câu trả lời.
    # `at` theo đồng hồ của api (đồng hồ thử được), `vn_day` là ngày giờ Việt Nam để tính trần ngày.
    # cell_edit.by_ai: ô do AI sửa qua khung chat rồi người dùng bấm Lưu (lịch sử bản ghi rõ, Q135).
    """
    CREATE TABLE chat_call (
        call_id bigserial PRIMARY KEY,
        document_id uuid NOT NULL REFERENCES document(document_id),
        at timestamptz NOT NULL,
        vn_day date NOT NULL,
        status text NOT NULL,
        reason text,
        cells_sent integer NOT NULL,
        masked_groups integer NOT NULL,
        ops integer NOT NULL DEFAULT 0,
        rejected_codes text[] NOT NULL DEFAULT '{}',
        input_tokens integer,
        output_tokens integer,
        estimated_cost_usd double precision NOT NULL,
        cost_usd double precision NOT NULL
    );
    CREATE INDEX chat_call_day ON chat_call (vn_day);
    ALTER TABLE cell_edit ADD COLUMN by_ai boolean NOT NULL DEFAULT false;
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
            conn.execute("INSERT INTO cell_edit (document_id, version, sheet, cell, value, formula, author, by_ai) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
                         (document_id, version, e["sheet"], e["cell"], json.dumps(e.get("value")), e.get("formula"), author, bool(e.get("by_ai"))))
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


class BatchChanged(Exception):
    """Quyết định của lô đã đổi ở chỗ khác (tab khác) từ lúc đọc tới lúc ghi."""


class FillConflict(Exception):
    """Ô của lô này đang chờ trong một lô khác đang mở của cùng tài liệu (D10)."""

    def __init__(self, cells: list[str]):
        super().__init__("conflict")
        self.cells = cells


def add_source(document_id: uuid.UUID, name: str, sha: str, size: int) -> dict:
    source_id = uuid.uuid4()
    with _connect() as conn:
        conn.execute("INSERT INTO source_file (source_id, document_id, name, sha256, size) VALUES (%s, %s, %s, %s, %s)",
                     (source_id, document_id, name, sha, size))
    return {"source_id": str(source_id), "name": name}


def source(document_id: uuid.UUID, source_id: uuid.UUID) -> dict | None:
    with _connect() as conn:
        return conn.execute("SELECT source_id, name, sha256 FROM source_file WHERE document_id = %s AND source_id = %s",
                            (document_id, source_id)).fetchone()


def _pending_cells(conn, document_id: uuid.UUID, exclude: uuid.UUID | None) -> dict[str, str]:
    """Ô đang chờ của mọi lô đang mở trên tài liệu: "sheet!A1" -> batch_id."""
    rows = conn.execute("SELECT batch_id, targets, decisions FROM fill_batch WHERE document_id = %s AND status = 'open' AND batch_id <> %s",
                        (document_id, exclude or uuid.UUID(int=0))).fetchall()
    out = {}
    for row in rows:
        for change_id, target in row["targets"].items():
            if row["decisions"].get(change_id, {}).get("state") == "pending":
                out[target] = str(row["batch_id"])
    return out


def create_batch(document_id: uuid.UUID, source_id: uuid.UUID, base_version: int, batch: dict, targets: dict, decisions: dict) -> uuid.UUID:
    """Lưu lô `draft`. Lô `draft` khác của tài liệu bị xóa: lô bỏ dở không nằm lại (acceptance_lat1 6.1)."""
    batch_id = uuid.uuid4()
    with _connect() as conn:
        conn.execute("SELECT 1 FROM document WHERE document_id = %s FOR UPDATE", (document_id,))
        conn.execute("DELETE FROM fill_batch WHERE document_id = %s AND status = 'draft'", (document_id,))
        conn.execute("INSERT INTO fill_batch (batch_id, document_id, source_id, base_version, status, batch, targets, decisions) "
                     "VALUES (%s, %s, %s, %s, 'draft', %s, %s, %s)",
                     (batch_id, document_id, source_id, base_version, json.dumps(batch, ensure_ascii=False),
                      json.dumps(targets, ensure_ascii=False), json.dumps(decisions)))
    return batch_id


def batch(batch_id: uuid.UUID) -> dict | None:
    with _connect() as conn:
        return conn.execute("SELECT b.*, s.name AS source_name FROM fill_batch b JOIN source_file s USING (source_id) WHERE b.batch_id = %s",
                            (batch_id,)).fetchone()


def open_batches(document_id: uuid.UUID) -> list[dict]:
    with _connect() as conn:
        return conn.execute("SELECT b.batch_id, b.created_at, s.name AS source_name FROM fill_batch b JOIN source_file s USING (source_id) "
                            "WHERE b.document_id = %s AND b.status = 'open' ORDER BY b.created_at", (document_id,)).fetchall()


def drop_batch(batch_id: uuid.UUID) -> dict:
    """Bỏ lần điền (N5). Lô nháp thì xóa hẳn. Lô đang mở: ô còn chờ thành `dropped`, lô xong; ô đã ghi giữ nguyên.

    Khóa dòng tài liệu như apply_batch: không đua với một lần Ghi đang chạy.
    """
    with _connect() as conn:
        row = conn.execute("SELECT document_id FROM fill_batch WHERE batch_id = %s", (batch_id,)).fetchone()
        if row is None:
            return {"deleted": False, "dropped": 0}
        conn.execute("SELECT 1 FROM document WHERE document_id = %s FOR UPDATE", (row["document_id"],))
        row = conn.execute("SELECT status, decisions FROM fill_batch WHERE batch_id = %s FOR UPDATE", (batch_id,)).fetchone()
        if row is None:
            return {"deleted": False, "dropped": 0}
        if row["status"] == "draft":
            conn.execute("DELETE FROM fill_batch WHERE batch_id = %s", (batch_id,))
            return {"deleted": True, "dropped": 0}
        if row["status"] != "open":
            return {"deleted": False, "dropped": 0}
        decisions = dict(row["decisions"])
        dropped = 0
        for cid, d in decisions.items():
            if d.get("state") == "pending":
                decisions[cid] = {"state": "dropped"}
                dropped += 1
        conn.execute("UPDATE fill_batch SET decisions = %s, status = 'done' WHERE batch_id = %s", (json.dumps(decisions), batch_id))
        return {"deleted": False, "dropped": dropped}


def open_batch(batch_id: uuid.UUID) -> None:
    """`draft` thành `open`. D10 kiểm trong cùng giao dịch, có khóa dòng tài liệu: hai tab cùng lúc chỉ một lô được mở."""
    with _connect() as conn:
        row = conn.execute("SELECT document_id FROM fill_batch WHERE batch_id = %s", (batch_id,)).fetchone()
        if row is None:
            raise LookupError(batch_id)
        conn.execute("SELECT 1 FROM document WHERE document_id = %s FOR UPDATE", (row["document_id"],))
        row = conn.execute("SELECT document_id, status, targets FROM fill_batch WHERE batch_id = %s", (batch_id,)).fetchone()
        if row is None:
            raise LookupError(batch_id)
        if row["status"] != "draft":
            return
        if not row["targets"]:  # lô 0 ô: không có gì để duyệt, kết thúc ngay, không báo "còn ô chờ" (N8)
            conn.execute("UPDATE fill_batch SET status = 'done', opened_at = now() WHERE batch_id = %s", (batch_id,))
            return
        busy = _pending_cells(conn, row["document_id"], batch_id)
        clash = sorted({t for t in row["targets"].values() if t in busy})
        if clash:
            raise FillConflict(clash)
        conn.execute("UPDATE fill_batch SET status = 'open', opened_at = now() WHERE batch_id = %s", (batch_id,))


def apply_batch(batch_id: uuid.UUID, document_id: uuid.UUID, base_version: int, decisions: dict, write: dict | None, expected: dict) -> dict | None:
    """Lưu quyết định, và nếu có ô để ghi thì thêm bản mới, trong một giao dịch có khóa dòng tài liệu.

    `write`: {sha, size, meta, author, summary, edits} hoặc None. Kiểm lại bản mới nhất và D10 ngay trong giao dịch:
    có lô khác vừa ghi hay vừa giữ ô của lô này thì không ghi gì (StaleVersion, FillConflict).
    `expected`: quyết định đã dùng để tính. Trong giao dịch, quyết định đang lưu khác bản đó (một lần áp khác vừa xong)
    thì không ghi gì (BatchChanged): không đè mất quyết định của lần kia.
    """
    with _connect() as conn:
        conn.execute("SELECT 1 FROM document WHERE document_id = %s FOR UPDATE", (document_id,))
        row = conn.execute("SELECT status, targets, decisions FROM fill_batch WHERE batch_id = %s FOR UPDATE", (batch_id,)).fetchone()
        if row is None or row["status"] != "open":
            raise LookupError(batch_id)
        if row["decisions"] != expected:
            raise BatchChanged
        latest_n = conn.execute("SELECT max(version) AS n FROM document_version WHERE document_id = %s", (document_id,)).fetchone()["n"]
        if latest_n != base_version:
            raise StaleVersion
        busy = _pending_cells(conn, document_id, batch_id)
        clash = sorted({row["targets"][cid] for cid, d in decisions.items()
                        if d.get("state") == "written" and d.get("version") is None and row["targets"].get(cid) in busy})
        if clash:
            raise FillConflict(clash)
        version_row = None
        if write is not None:
            version = latest_n + 1
            try:
                conn.execute("INSERT INTO document_version (document_id, version, sha256, size, meta, author, summary) VALUES (%s, %s, %s, %s, %s, %s, %s)",
                             (document_id, version, write["sha"], write["size"], json.dumps(write["meta"], ensure_ascii=False), write["author"], write["summary"]))
            except psycopg.errors.UniqueViolation as exc:
                # Lưới an toàn: khóa dòng tài liệu ở trên làm việc này không xảy ra. Nếu có, đó là bản lệch, không phải lỗi 500.
                raise StaleVersion from exc
            for e in write["edits"]:
                conn.execute("INSERT INTO cell_edit (document_id, version, sheet, cell, value, formula, author) VALUES (%s, %s, %s, %s, %s, %s, %s)",
                             (document_id, version, e["sheet"], e["cell"], json.dumps(e.get("value")), None, write["author"]))
            for d in decisions.values():
                if d.get("state") == "written" and d.get("version") is None:
                    d["version"] = version
            version_row = conn.execute(_SELECT + " WHERE d.document_id = %s AND v.version = %s", (document_id, version)).fetchone()
        status = "open" if any(d.get("state") == "pending" for d in decisions.values()) else "done"
        conn.execute("UPDATE fill_batch SET decisions = %s, status = %s WHERE batch_id = %s", (json.dumps(decisions), status, batch_id))
        return version_row


def chat_spent(day) -> float:
    """Tiền đã dùng trong một ngày giờ Việt Nam (theo bảng giá, không phải hóa đơn)."""
    with _connect() as conn:
        return float(conn.execute("SELECT coalesce(sum(cost_usd), 0) AS s FROM chat_call WHERE vn_day = %s", (day,)).fetchone()["s"])


def add_chat_call(row: dict) -> None:
    """Một lượt gọi. Chỉ số đếm và mã, không có nội dung."""
    cols = ("document_id", "at", "vn_day", "status", "reason", "cells_sent", "masked_groups", "ops", "rejected_codes",
            "input_tokens", "output_tokens", "estimated_cost_usd", "cost_usd")
    with _connect() as conn:
        conn.execute(f"INSERT INTO chat_call ({', '.join(cols)}) VALUES ({', '.join(['%s'] * len(cols))})", tuple(row[c] for c in cols))
