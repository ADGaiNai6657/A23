import hashlib
import json
import sqlite3
from contextlib import contextmanager
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
        digest = hashlib.sha256(content).hexdigest()
        if parsed.file.lower().endswith((".txt", ".md")):
            digest = "line-number:" + digest
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            current = db.execute(
                "SELECT digest FROM documents WHERE name = ?", (parsed.file,),
            ).fetchone()
            if current is None or current["digest"] != digest:
                db.execute("DELETE FROM documents WHERE name = ?", (parsed.file,))
                db.execute("INSERT INTO documents VALUES (?, ?, ?)", (parsed.file, digest, content))
                db.executemany("INSERT INTO blocks VALUES (?, ?, ?, ?, ?)", (
                    (parsed.file, block_no, block.type, block.text, json.dumps(block.locator, ensure_ascii=False))
                    for block_no, block in parsed.numbered_blocks()
                ))
            if mode is not None:
                db.execute("DELETE FROM facts WHERE doc = ? AND mode = ?", (parsed.file, mode))
                db.executemany("""
                    INSERT INTO facts VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, ((parsed.file, mode, *fact) for fact in facts))

    def content(self, doc: str) -> bytes | None:
        with self.connect() as db:
            row = db.execute("SELECT content FROM documents WHERE name = ?", (doc,)).fetchone()
            return row["content"] if row else None

    def blocks(self, doc: str) -> list[dict]:
        with self.connect() as db:
            return [
                {**dict(row), "locator": json.loads(row["locator"])}
                for row in db.execute("SELECT * FROM blocks WHERE doc = ? ORDER BY number", (doc,))
            ]

    def facts(self, doc: str | None = None) -> list[dict]:
        with self.connect() as db:
            statement = "SELECT * FROM facts"
            args = ()
            if doc is not None:
                statement += " WHERE doc = ?"
                args = (doc,)
            return [dict(row) for row in db.execute(
                statement + " ORDER BY doc, block_no, rowid", args,
            )]

    def block(self, doc: str, number: int) -> dict | None:
        with self.connect() as db:
            row = db.execute(
                "SELECT * FROM blocks WHERE doc = ? AND number = ?", (doc, number),
            ).fetchone()
            return {**dict(row), "locator": json.loads(row["locator"])} if row else None
