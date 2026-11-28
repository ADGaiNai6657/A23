import hashlib
import json
import sqlite3
from collections.abc import Callable
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

from config import DATA
from schema import ParsedDoc

SCHEMA = """
CREATE TABLE IF NOT EXISTS documents (
    name TEXT PRIMARY KEY, digest TEXT NOT NULL, content BLOB NOT NULL
);
CREATE TABLE IF NOT EXISTS blocks (
    doc TEXT NOT NULL REFERENCES documents(name) ON DELETE CASCADE,
    number INTEGER NOT NULL, type TEXT NOT NULL, text TEXT NOT NULL, locator TEXT NOT NULL,
    PRIMARY KEY(doc, number)
);
CREATE TABLE IF NOT EXISTS facts (
    doc TEXT NOT NULL, mode TEXT NOT NULL, block_no INTEGER NOT NULL,
    entity TEXT NOT NULL, metric TEXT NOT NULL, metric_key TEXT NOT NULL,
    value TEXT NOT NULL, value_num TEXT, unit TEXT NOT NULL, snippet TEXT NOT NULL,
    period TEXT NOT NULL, period_block_no INTEGER, scope TEXT NOT NULL,
    FOREIGN KEY(doc, block_no) REFERENCES blocks(doc, number) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS facts_lookup ON facts(entity, metric_key, period, scope);
CREATE INDEX IF NOT EXISTS facts_document ON facts(doc, mode, block_no);
CREATE TABLE IF NOT EXISTS extraction_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT, doc TEXT NOT NULL, source_sha256 TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS published_versions (
    doc TEXT PRIMARY KEY, request_id INTEGER NOT NULL, source_sha256 TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS extraction_tasks (
    id TEXT PRIMARY KEY, doc TEXT NOT NULL, source_sha256 TEXT NOT NULL,
    request_id INTEGER, status TEXT NOT NULL, submitted_at TEXT NOT NULL,
    started_at TEXT, finished_at TEXT, result TEXT, error TEXT
);
"""


class Store:
    def __init__(self, path: Path = DATA / "a23.db"):
        self.path = path

    @contextmanager
    def connect(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        try:
            connection.execute("PRAGMA foreign_keys = ON")
            connection.executescript(SCHEMA)
            with connection:
                yield connection
        finally:
            connection.close()

    def save(self, parsed: ParsedDoc, content: bytes, mode: str | None = None, facts=()) -> None:
        source_sha256 = hashlib.sha256(content).hexdigest()
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            same_document = self._write_document(db, parsed, content, mode, facts)
            if mode is None and same_document:
                # Reimporting identical bytes and blocks does not publish a new version.
                return
            request = db.execute(
                "INSERT INTO extraction_requests(doc, source_sha256) VALUES (?, ?)",
                (parsed.file, source_sha256),
            )
            db.execute("""INSERT INTO published_versions VALUES (?, ?, ?)
                          ON CONFLICT(doc) DO UPDATE SET request_id=excluded.request_id,
                          source_sha256=excluded.source_sha256""",
                       (parsed.file, request.lastrowid, source_sha256))

    @staticmethod
    def _write_document(db, parsed: ParsedDoc, content: bytes, mode: str | None, facts) -> bool:
        digest = hashlib.sha256(content).hexdigest()
        if parsed.file.lower().endswith((".txt", ".md")):
            digest = "line-number:" + digest
        incoming = [
            (block_no, block.type, block.text, json.dumps(block.locator, ensure_ascii=False, sort_keys=True))
            for block_no, block in parsed.numbered_blocks()
        ]
        current = db.execute("SELECT digest, content FROM documents WHERE name = ?", (parsed.file,)).fetchone()
        same_blocks = False
        if current is not None and current["digest"] == digest and current["content"] == content:
            saved = [
                (row["number"], row["type"], row["text"],
                 json.dumps(json.loads(row["locator"]), ensure_ascii=False, sort_keys=True))
                for row in db.execute("SELECT number, type, text, locator FROM blocks WHERE doc = ? ORDER BY number",
                                      (parsed.file,))
            ]
            same_blocks = saved == incoming
        if not same_blocks:
            db.execute("DELETE FROM documents WHERE name = ?", (parsed.file,))
            db.execute("INSERT INTO documents VALUES (?, ?, ?)", (parsed.file, digest, content))
            db.executemany("INSERT INTO blocks VALUES (?, ?, ?, ?, ?)", (
                (parsed.file, *block) for block in incoming
            ))
        if mode is not None:
            db.execute("DELETE FROM facts WHERE doc = ? AND mode = ?", (parsed.file, mode))
            db.executemany("INSERT INTO facts VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                           ((parsed.file, mode, *fact) for fact in facts))
        return same_blocks

    def register_extraction(self, doc: str, content: bytes) -> int:
        with self.connect() as db:
            saved = db.execute("INSERT INTO extraction_requests(doc, source_sha256) VALUES (?, ?)",
                               (doc, hashlib.sha256(content).hexdigest()))
            return saved.lastrowid

    def publish_extraction(self, parsed: ParsedDoc, content: bytes, mode: str, facts,
                           request_id: int, result: dict, task_id: str | None = None) -> bool:
        """Commit facts and a task's terminal state together, respecting request order."""
        digest = hashlib.sha256(content).hexdigest()
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            request = db.execute("SELECT doc, source_sha256 FROM extraction_requests WHERE id = ?",
                                 (request_id,)).fetchone()
            if request is None or request["doc"] != parsed.file or request["source_sha256"] != digest:
                raise ValueError("抽取请求与原文件版本不一致")
            winner = db.execute("SELECT request_id FROM published_versions WHERE doc = ?",
                                (parsed.file,)).fetchone()
            published = winner is None or request_id >= winner["request_id"]
            if published:
                self._write_document(db, parsed, content, mode, facts)
                db.execute("""INSERT INTO published_versions VALUES (?, ?, ?)
                              ON CONFLICT(doc) DO UPDATE SET request_id=excluded.request_id,
                              source_sha256=excluded.source_sha256""",
                           (parsed.file, request_id, digest))
            if task_id is not None:
                updated = db.execute("""
                    UPDATE extraction_tasks SET status=?, finished_at=?, result=?, error=NULL
                    WHERE id=? AND request_id=? AND status='running'
                """, ("succeeded" if published else "superseded",
                      datetime.now(timezone.utc).isoformat(timespec="milliseconds"),
                      json.dumps(result, ensure_ascii=False), task_id, request_id))
                if updated.rowcount != 1:
                    raise ValueError("识别任务状态与提交版本不一致")
        return published

    def content(self, doc: str) -> bytes | None:
        with self.connect() as db:
            row = db.execute("SELECT content FROM documents WHERE name = ?", (doc,)).fetchone()
            return row["content"] if row else None

    def blocks(self, doc: str) -> list[dict]:
        with self.connect() as db:
            return self._blocks(db, doc)

    @staticmethod
    def _blocks(db, doc: str) -> list[dict]:
        return [
            {**dict(row), "locator": json.loads(row["locator"])}
            for row in db.execute("SELECT * FROM blocks WHERE doc = ? ORDER BY number", (doc,))
        ]

    def facts(self, doc: str | None = None) -> list[dict]:
        with self.connect() as db:
            return self._facts(db, doc)

    def record_fact_decision(self, criteria: dict, decide: Callable[[list[dict]], dict]) -> dict:
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            db.execute("""
                CREATE TABLE IF NOT EXISTS fact_decisions (
                    id INTEGER PRIMARY KEY, payload TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
            """)
            digests = {row["name"]: row["digest"].rsplit(":", 1)[-1]
                       for row in db.execute("SELECT name, digest FROM documents")}
            candidates = [{**fact, "source_sha256": digests[fact["doc"]]} for fact in self._facts(db, None)]
            payload = {"criteria": criteria, **decide(candidates)}
            saved = db.execute("INSERT INTO fact_decisions(payload) VALUES (?)",
                               (json.dumps(payload, ensure_ascii=False),))
            return {"decision_id": saved.lastrowid, **payload}

    def fact_decision(self, decision_id: int) -> dict | None:
        if not self.path.is_file():
            return None
        with self.connect() as db:
            exists = db.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='fact_decisions'").fetchone()
            if exists is None:
                return None
            row = db.execute("SELECT payload, created_at FROM fact_decisions WHERE id = ?", (decision_id,)).fetchone()
            return {"decision_id": decision_id, "created_at": row["created_at"], **json.loads(row["payload"])} if row else None

    @staticmethod
    def _facts(db, doc: str | None) -> list[dict]:
        statement = "SELECT * FROM facts"
        args = ()
        if doc is not None:
            statement += " WHERE doc = ?"
            args = (doc,)
        return [dict(row) for row in db.execute(statement + " ORDER BY doc, block_no, rowid", args)]

    def snapshot(self, doc: str, *, with_facts: bool, expected_publication_id: int | None = None
                 ) -> tuple[bytes, list[dict], list[dict], int | None] | None:
        with self.connect() as db:
            db.execute("BEGIN")
            row = db.execute("SELECT content FROM documents WHERE name = ?", (doc,)).fetchone()
            if row is None:
                return None
            publication = db.execute("SELECT request_id FROM published_versions WHERE doc = ?",
                                     (doc,)).fetchone()
            publication_id = publication["request_id"] if publication else None
            if expected_publication_id is not None and publication_id != expected_publication_id:
                raise ValueError("事实版本冲突：已发布的抽取批次改变，请重新识别并填写")
            return (row["content"], self._blocks(db, doc),
                    self._facts(db, doc) if with_facts else [], publication_id)

    def block(self, doc: str, number: int) -> dict | None:
        with self.connect() as db:
            row = db.execute(
                """
                SELECT blocks.*, documents.digest AS document_digest
                FROM blocks JOIN documents ON documents.name = blocks.doc
                WHERE blocks.doc = ? AND number = ?
                """, (doc, number),
            ).fetchone()
            return {**dict(row), "locator": json.loads(row["locator"])} if row else None
