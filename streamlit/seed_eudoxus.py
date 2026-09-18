"""Load the Εύδοξος book lists the department exported into Postgres.

``EXPORTS`` names one ``files/eudoxus`` file per academic year — the 2025-26
list arrived as ``eudoxus_books_2025-26.xlsx``, the 2026-27 one as the CSV
Εύδοξος downloads («Συγγράμματα ΕΥΔΟΞΟΣ 2026-2027.csv») — and
``eudoxus_catalogue_<YYYYMMDD>.csv`` is a one-off dump of what Eudoxus said
about those books, fetched once so the browse tab shows titles from the first
run instead of bare numeric codes.

These are the years the app did not produce: from here on a list is opened as a
copy of the previous one and edited in the app (`eudoxus_db.open_year`). A year
is seeded as history (ΚΛΕΙΔΩΜΕΝΟ) unless it is in ``OPEN_SEED_YEARS``, which is
for a list still being worked on when its file arrived.

Called from :func:`db.bootstrap` inside Railway, and idempotent: a year that
already has courses is skipped, and the catalogue is loaded only while the
table is empty — after that the availability check owns it and a restart must
not overwrite fresher answers with the dump.
"""

from __future__ import annotations

import re
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
from sqlalchemy import Engine, text

from eudoxus_db import (
    BOOKS_TABLE,
    COURSES_TABLE,
    LOCKED,
    OPEN,
    SELECTIONS_TABLE,
    UPSERT_BOOK_SQL,
    YEARS_TABLE,
    _clean_book,
    year_label,
)

ROOT = Path(__file__).resolve().parents[1]
EUDOXUS_DIR = ROOT / "files" / "eudoxus"

# The academic year is read out of the filename rather than from a fixed
# prefix: the first export was renamed by hand to ``eudoxus_books_2025-26``,
# the 2026-27 one was dropped in under the name Εύδοξος gives it
# («Συγγράμματα ΕΥΔΟΞΟΣ 2026-2027.csv»). Both forms of the second half are
# accepted, and it must be the following year — so a date stamp such as the
# catalogue's cannot be mistaken for a year range.
YEAR_RE = re.compile(r"(\d{4})-(\d{2}|\d{4})(?!\d)")
EXPORT_SUFFIXES = (".xlsx", ".csv")

# Years imported as the list *under preparation* instead of as history, with
# the year they are compared against. The department declared 2026-27 in
# Εύδοξος before this app existed, so the list arrives as a file like 2025-26
# did — but it is the year people are still working on, so it is seeded
# ΑΝΟΙΧΤΟ with a baseline, and the admin tab's «Μεταβολές» compares it against
# 2025-26 exactly as if it had been opened from inside the app.
OPEN_SEED_YEARS = {2026: 2025}

# Export header -> column. The department's export is stable enough to match
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


def read_export(path: Path) -> tuple[list[dict], list[dict], int]:
    """(course offerings, selections, rows skipped) from one export.

    Both formats the department has handed over are read here: the 2025-26
    export arrived as .xlsx, the 2026-27 one as the .csv Εύδοξος downloads.
    The columns are identical, so only the reader differs.
    """
    frame = pd.read_csv(path) if path.suffix == ".csv" else pd.read_excel(path)
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


def export_year(path: Path) -> int | None:
    """The academic year one export covers, or None if the name does not say."""
    if path.suffix not in EXPORT_SUFFIXES or path.stem.startswith("eudoxus_catalogue_"):
        return None
    match = YEAR_RE.search(path.stem)
    if not match:
        return None
    year, second = int(match.group(1)), match.group(2)
    expected = str(year + 1)
    return year if second == (expected if len(second) == 4 else expected[-2:]) else None


# The exports, one per year, named explicitly rather than globbed: working
# copies get dropped beside the real file (three of them on 2026-09-18, one
# with rows that have no εξάμηνο), and a glob would seed whichever sorts
# first. Same reason ``seed_timetable.WORKBOOKS`` is a dict.
EXPORTS = {
    2025: "eudoxus_books_2025-26.xlsx",
    2026: "Συγγράμματα ΕΥΔΟΞΟΣ 2026-2027.csv",
}


def find_exports() -> list[tuple[int, Path]]:
    """The exports, oldest year first.

    Year order matters: a year is seeded with the previous one as its
    baseline, so 2025-26 has to be in the table before 2026-27 arrives.
    """
    return [(year, EUDOXUS_DIR / EXPORTS[year]) for year in sorted(EXPORTS)]


def _seed_status(engine: Engine, year: int) -> tuple[str, int | None]:
    """(status, baseline) for a year about to be seeded.

    ``OPEN_SEED_YEARS`` asks for a year to arrive open for editing, but two
    conditions can withdraw that, and both leave the year as ordinary history
    rather than failing the whole bootstrap: the baseline it is compared
    against has to exist, and only one year may be open at a time — a
    coordinator who has already opened one in the app must not find a second
    one appearing underneath them.
    """
    baseline = OPEN_SEED_YEARS.get(year)
    if baseline is None:
        return LOCKED, None
    with engine.connect() as conn:
        has_baseline = conn.execute(
            text(f"SELECT 1 FROM {COURSES_TABLE} WHERE year = :year LIMIT 1"),
            {"year": baseline},
        ).first()
        already_open = conn.execute(
            text(f"SELECT 1 FROM {YEARS_TABLE} WHERE status = :status LIMIT 1"),
            {"status": OPEN},
        ).first()
    if not has_baseline or already_open:
        return LOCKED, None
    return OPEN, baseline


def seed_eudoxus(engine: Engine) -> str:
    """Insert every export that has no rows yet, plus the catalogue dump."""
    loaded, skipped = [], []
    for year, path in find_exports():
        with engine.connect() as conn:
            already = conn.execute(
                text(f"SELECT 1 FROM {COURSES_TABLE} WHERE year = :year LIMIT 1"),
                {"year": year},
            ).first()
        if already:
            skipped.append(year_label(year))
            continue

        courses, selections, inactive = read_export(path)
        status, baseline = _seed_status(engine, year)
        with engine.begin() as conn:
            # Seeded years are history unless OPEN_SEED_YEARS says otherwise:
            # they were submitted long ago, and only one year is ever open for
            # editing.
            conn.execute(
                text(
                    f"INSERT INTO {YEARS_TABLE} "
                    "(year, status, baseline_year, opened_by, locked_by, locked_at) "
                    "VALUES (:year, :status, :baseline, 'seed', :locked_by, :locked_at) "
                    "ON CONFLICT (year) DO NOTHING"
                ),
                {
                    "year": year,
                    "status": status,
                    "baseline": baseline,
                    # An open year has not been locked by anyone yet.
                    "locked_by": "seed" if status == LOCKED else None,
                    "locked_at": datetime.now(timezone.utc) if status == LOCKED else None,
                },
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
        note += f", {inactive} ανενεργές γραμμές" if inactive else ""
        note += f", {status}"
        note += f", βάση {year_label(baseline)})" if baseline else ")"
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
