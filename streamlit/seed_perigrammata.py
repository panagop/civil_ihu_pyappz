"""Load the frozen περιγράμματα exports into Postgres.

``files/perigrammata/perigrammata_<locale>_<curriculum>.csv`` is a one-off
export of the Google Sheet as it stood at handover — the archive, and the only
thing that ever reads it is this module. After the first successful seed the
database is the master and the sheet plays no part in the app.

Like :mod:`seed_external` this runs from :func:`db.bootstrap` inside Railway,
because nothing outside it can reach the database, and it is idempotent: a
(curriculum, locale) that already has rows is skipped entirely.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd
from sqlalchemy import Engine, text

from perigrammata_db import (
    CONTENT_COLUMNS,
    COURSES_TABLE,
    INTEGER_COLUMNS,
    NUMERIC_COLUMNS,
    REVISIONS_TABLE,
)

ROOT = Path(__file__).resolve().parents[1]
ARCHIVE_DIR = ROOT / "files" / "perigrammata"
FILENAME_RE = re.compile(r"perigrammata_(gr|eng)_(\d{4})$")

# Greek only for now. The English sheet writes εξάμηνο as "1st"/"2nd", which
# needs a mapping before it can go into an INTEGER column — that is part of
# doing the English version properly, not of this import. Adding "eng" here is
# the one-line half of it.
SEED_LOCALES = ("gr",)

SEED_NOTE = "Αρχική φόρτωση από το Google Sheet"


def _coerce(field: str, value: object) -> object:
    """One archive cell as the column stores it."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    raw = str(value).strip()
    if not raw:
        return None
    if field in NUMERIC_COLUMNS or field in INTEGER_COLUMNS:
        number = float(raw.replace(",", "."))
        return int(number) if field in INTEGER_COLUMNS else number
    return raw


def read_archive(path: Path, curriculum: int, locale: str) -> tuple[list[dict], int]:
    """Insert-ready rows from one archive file, plus the count of skipped ones.

    The 2018 worksheet carries a row with no ``code`` and a stray sentence in
    one cell — spreadsheet debris rather than a course. It is dropped, but
    counted, so the seed log says so instead of quietly loading 102 of 103.
    """
    frame = pd.read_csv(path, dtype=str)
    rows: list[dict] = []
    skipped = 0
    for _, record in frame.iterrows():
        code = record.get("code")
        if code is None or (isinstance(code, float) and pd.isna(code)) or not str(code).strip():
            skipped += 1
            continue
        order = record.get("sort_order")
        row: dict[str, object] = {
            "curriculum": curriculum,
            "locale": locale,
            "code": str(code).strip(),
            "sort_order": None if pd.isna(order) else int(float(order)),
        }
        for field in CONTENT_COLUMNS:
            row[field] = _coerce(field, record.get(field))
        rows.append(row)
    return rows, skipped


def _insert_sql() -> str:
    columns = ["curriculum", "locale", "code", "sort_order", *CONTENT_COLUMNS]
    names = ", ".join(columns)
    binds = ", ".join(f":{name}" for name in columns)
    return (
        f"INSERT INTO {COURSES_TABLE} ({names}) VALUES ({binds}) "
        "ON CONFLICT (curriculum, locale, code) DO NOTHING"
    )


def _seed_revisions(conn, rows: list[dict]) -> None:
    """One origin revision per course.

    Written with an empty ``changes`` object on purpose: the import is not a
    change anybody made, so it must not surface in the coordinator's "what
    moved this period" report. It exists so the earliest state of every course
    is recoverable from the history alone.
    """
    conn.execute(
        text(
            f"INSERT INTO {REVISIONS_TABLE} "
            "(curriculum, locale, code, data, changes, edited_by, note) "
            "VALUES (:curriculum, :locale, :code, CAST(:data AS jsonb), "
            "CAST('{}' AS jsonb), :editor, :note)"
        ),
        [
            {
                "curriculum": row["curriculum"],
                "locale": row["locale"],
                "code": row["code"],
                "data": json.dumps(row, ensure_ascii=False, default=str),
                "editor": "seed",
                "note": SEED_NOTE,
            }
            for row in rows
        ],
    )


def seed_perigrammata(engine: Engine) -> str:
    """Insert every archive file that has no rows yet."""
    loaded, skipped_files = [], []
    for path in sorted(ARCHIVE_DIR.glob("perigrammata_*.csv")):
        match = FILENAME_RE.match(path.stem)
        if not match:
            continue
        locale, curriculum = match.group(1), int(match.group(2))
        if locale not in SEED_LOCALES:
            continue

        with engine.connect() as conn:
            already = conn.execute(
                text(
                    f"SELECT 1 FROM {COURSES_TABLE} "
                    "WHERE curriculum = :curriculum AND locale = :locale LIMIT 1"
                ),
                {"curriculum": curriculum, "locale": locale},
            ).first()
        if already:
            skipped_files.append(f"{locale} {curriculum}")
            continue

        rows, dropped = read_archive(path, curriculum, locale)
        with engine.begin() as conn:
            conn.execute(text(_insert_sql()), rows)
            _seed_revisions(conn, rows)
        note = f"{locale} {curriculum} ({len(rows)} μαθήματα"
        note += f", {dropped} γραμμές χωρίς κωδικό)" if dropped else ")"
        loaded.append(note)

    parts = []
    if loaded:
        parts.append("φορτώθηκαν: " + ", ".join(loaded))
    if skipped_files:
        parts.append("υπήρχαν ήδη: " + ", ".join(skipped_files))
    return "Περιγράμματα — " + ("· ".join(parts) if parts else "κανένα αρχείο")
