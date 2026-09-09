"""Load the historical Εύδοξος book list into Postgres.

``files/eudoxus/eudoxus_books_<YYYY>-<YY>.xlsx`` is the department's own export
of the books offered in that academic year, and
``files/eudoxus/eudoxus_catalogue_<YYYYMMDD>.csv`` is a one-off dump of what
Eudoxus said about each of those books, fetched once so the browse tab shows
titles from the first run instead of 222 bare numeric codes.

This is the *only* import: from the next year on a list is opened as a copy of
the previous one and edited in the app (`eudoxus_db.open_year`), so this seeder
exists for the years that predate the app.

Called from :func:`db.bootstrap` inside Railway, and idempotent: a year that
already has courses is skipped, and the catalogue is loaded only while the
table is empty — after that the availability check owns it and a restart must
not overwrite fresher answers with the dump.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from sqlalchemy import Engine, text

from eudoxus_db import (
    BOOKS_TABLE,
    COURSES_TABLE,
    LOCKED,
    SELECTIONS_TABLE,
    UPSERT_BOOK_SQL,
    YEARS_TABLE,
    _clean_book,
    year_label,
)

ROOT = Path(__file__).resolve().parents[1]
EUDOXUS_DIR = ROOT / "files" / "eudoxus"
WORKBOOK_RE = re.compile(r"eudoxus_books_(\d{4})-\d{2}$")

# Workbook header -> column. The department's export is stable enough to match
# by name; anything unexpected raises rather than importing a wrong column.
COLUMNS = {
    "Τίτλος": "title",
    "Καθηγητής": "teacher",
    "Κωδικός μαθήματος": "course_code",
    "Εξάμηνο": "examino",
    "Περίοδος": "period",
    "Από που δίνονται τα βιβλία": "source",
    "Book id": "book_id",
    "Σειρά επιλογής": "priority",
    "Ενεργό": "active_row",
}

# The catalogue dump carries Eudoxus' own field names.
CATALOGUE_COLUMNS = {
    "id": "book_id",
    "title": "title",
    "subtitle": "subtitle",
    "authors": "authors",
    "isbn": "isbn",
    "publisherName": "publisher",
    "publicationYear": "publication_year",
    "editionNumber": "edition_number",
    "active": "active",
    "selectable": "selectable",
    "found": "found",
    "error": "error",
}


def read_workbook(path: Path) -> tuple[list[dict], list[dict], int]:
    """(course offerings, selections, rows skipped) from one export."""
    frame = pd.read_excel(path)
    missing = set(COLUMNS) - set(frame.columns)
    if missing:
        raise ValueError(f"{path.name}: λείπουν στήλες {sorted(missing)}")
    frame = frame.rename(columns=COLUMNS)

    # Every row of the 2025-26 export is active, but the column exists, so
    # honour it rather than assume it will stay constant.
    inactive = int((~frame["active_row"].fillna(True).astype(bool)).sum())
    frame = frame[frame["active_row"].fillna(True).astype(bool)]

    selections: list[dict] = []
    for _, row in frame.iterrows():
        selections.append(
            {
                "course_code": str(row["course_code"]).strip(),
                "examino": int(row["examino"]),
                "book_id": int(row["book_id"]),
                "priority": int(row["priority"]),
                "source": (
                    "EUDOXUS" if pd.isna(row["source"]) else str(row["source"]).strip()
                ),
            }
        )

    # One offering per (course, εξάμηνο) — ΔΟΜ022 legitimately has two.
    offerings = (
        frame.groupby(["course_code", "examino"], as_index=False)
        .agg({"title": "first", "teacher": "first", "period": "first"})
    )
    courses = [
        {
            "course_code": str(row["course_code"]).strip(),
            "examino": int(row["examino"]),
            "title": str(row["title"]).strip(),
            "teacher": None if pd.isna(row["teacher"]) else str(row["teacher"]).strip(),
            "period": None if pd.isna(row["period"]) else str(row["period"]).strip(),
        }
        for _, row in offerings.iterrows()
    ]
    return courses, selections, inactive


def _latest_catalogue() -> Path | None:
    files = sorted(EUDOXUS_DIR.glob("eudoxus_catalogue_*.csv"))
    return files[-1] if files else None


def read_catalogue(path: Path) -> list[dict]:
    frame = pd.read_csv(path)
    frame = frame.rename(columns=CATALOGUE_COLUMNS)
    rows: list[dict] = []
    for _, row in frame.iterrows():
        record = {
            target: (None if pd.isna(row.get(target)) else row.get(target))
            for target in set(CATALOGUE_COLUMNS.values())
        }
        rows.append(_clean_book(record))
    return rows


def seed_eudoxus(engine: Engine) -> str:
    """Insert every export that has no rows yet, plus the catalogue dump."""
    loaded, skipped = [], []
    for path in sorted(EUDOXUS_DIR.glob("eudoxus_books_*.xlsx")):
        match = WORKBOOK_RE.match(path.stem)
        if not match:
            continue
        year = int(match.group(1))

        with engine.connect() as conn:
            already = conn.execute(
                text(f"SELECT 1 FROM {COURSES_TABLE} WHERE year = :year LIMIT 1"),
                {"year": year},
            ).first()
        if already:
            skipped.append(year_label(year))
            continue

        courses, selections, inactive = read_workbook(path)
        with engine.begin() as conn:
            # Seeded years are history: they were submitted long ago, and only
            # one year is ever open for editing.
            conn.execute(
                text(
                    f"INSERT INTO {YEARS_TABLE} (year, status, opened_by, locked_by, locked_at) "
                    "VALUES (:year, :status, 'seed', 'seed', now()) "
                    "ON CONFLICT (year) DO NOTHING"
                ),
                {"year": year, "status": LOCKED},
            )
            conn.execute(
                text(
                    f"INSERT INTO {COURSES_TABLE} "
                    "(year, course_code, examino, title, teacher, period) "
                    "VALUES (:year, :course_code, :examino, :title, :teacher, :period) "
                    "ON CONFLICT (year, course_code, examino) DO NOTHING"
                ),
                [{**course, "year": year} for course in courses],
            )
            conn.execute(
                text(
                    f"INSERT INTO {SELECTIONS_TABLE} "
                    "(year, course_code, examino, book_id, priority, source, added_by) "
                    "VALUES (:year, :course_code, :examino, :book_id, :priority, :source, 'seed') "
                    "ON CONFLICT (year, course_code, examino, book_id) DO NOTHING"
                ),
                [{**selection, "year": year} for selection in selections],
            )
        note = f"{year_label(year)} ({len(courses)} μαθήματα, {len(selections)} επιλογές"
        note += f", {inactive} ανενεργές γραμμές)" if inactive else ")"
        loaded.append(note)

    catalogue_note = _seed_catalogue(engine)

    parts = []
    if loaded:
        parts.append("φορτώθηκαν: " + ", ".join(loaded))
    if skipped:
        parts.append("υπήρχαν ήδη: " + ", ".join(skipped))
    if catalogue_note:
        parts.append(catalogue_note)
    return "Εύδοξος — " + ("· ".join(parts) if parts else "κανένα αρχείο")


def _seed_catalogue(engine: Engine) -> str:
    """Load the dump, but only while the catalogue is empty.

    After the first load the availability check owns this table. Re-applying the
    dump on every restart would quietly replace a fresh answer with a stale one.
    """
    path = _latest_catalogue()
    if path is None:
        return "χωρίς κατάλογο βιβλίων"
    with engine.connect() as conn:
        already = conn.execute(text(f"SELECT 1 FROM {BOOKS_TABLE} LIMIT 1")).first()
    if already:
        return "κατάλογος: υπήρχε ήδη"
    rows = read_catalogue(path)
    with engine.begin() as conn:
        conn.execute(text(UPSERT_BOOK_SQL), rows)
    return f"κατάλογος: {len(rows)} βιβλία από {path.name}"
