import json
import os
import sqlite3
from typing import Any, Dict

DB_PATH = os.path.join(os.path.dirname(__file__), 'state.db')


def _connect():
    return sqlite3.connect(DB_PATH)


def init_db() -> None:
    with _connect() as conn:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS kv (key TEXT PRIMARY KEY, value TEXT NOT NULL)"
        )
        # Persist parsed documents (one row per filename)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS docs (
                filename TEXT PRIMARY KEY,
                data TEXT NOT NULL
            )
            """
        )
        conn.commit()


DEFAULT_STATE: Dict[str, Any] = {
    "colorMap": {},
    "codeMap": {},
    "categoryColors": {},
    "catOverride": {},
    "meta": {},
    "extraCategories": ["XX_X"],
}


def get_state() -> Dict[str, Any]:
    init_db()
    data: Dict[str, Any] = {}
    with _connect() as conn:
        cur = conn.execute("SELECT key, value FROM kv")
        for k, v in cur.fetchall():
            try:
                data[k] = json.loads(v)
            except Exception:
                data[k] = v
    # merge with defaults
    out = DEFAULT_STATE.copy()
    out.update({k: v for k, v in data.items() if k in DEFAULT_STATE})
    return out


def set_state(partial: Dict[str, Any]) -> Dict[str, Any]:
    init_db()
    # Only allow known keys
    allowed = set(DEFAULT_STATE.keys())
    with _connect() as conn:
        for k, v in partial.items():
            if k in allowed:
                conn.execute(
                    "REPLACE INTO kv(key, value) VALUES (?, ?)", (k, json.dumps(v))
                )
        conn.commit()
    return get_state()


# ------------------- Docs persistence -------------------
def save_docs(items: Dict[str, Dict[str, Any]]) -> None:
    """Save multiple parsed documents. items: {filename: parsed_dict}.
    parsed_dict should have highlights, comments, paragraphs arrays.
    """
    init_db()
    with _connect() as conn:
        for fn, parsed in items.items():
            conn.execute(
                "REPLACE INTO docs(filename, data) VALUES (?, ?)", (fn, json.dumps(parsed))
            )
        conn.commit()


def delete_doc(filename: str) -> None:
    init_db()
    with _connect() as conn:
        conn.execute("DELETE FROM docs WHERE filename = ?", (filename,))
        conn.commit()


def list_docs() -> Dict[str, Any]:
    """Return aggregated data across all stored docs."""
    init_db()
    highlights: list[Dict[str, Any]] = []
    comments: list[Dict[str, Any]] = []
    paragraphs: list[Dict[str, Any]] = []
    with _connect() as conn:
        cur = conn.execute("SELECT data FROM docs")
        for (data_str,) in cur.fetchall():
            try:
                d = json.loads(data_str)
                highlights.extend(d.get("highlights", []))
                comments.extend(d.get("comments", []))
                paragraphs.extend(d.get("paragraphs", []))
            except Exception:
                continue
    return {"highlights": highlights, "comments": comments, "paragraphs": paragraphs}


def list_filenames() -> list[str]:
    init_db()
    with _connect() as conn:
        cur = conn.execute("SELECT filename FROM docs ORDER BY filename")
        return [r[0] for r in cur.fetchall()]
