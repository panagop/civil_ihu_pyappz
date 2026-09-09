"""Postgres access for the Εύδοξος book lists.

A year here is the academic year named by the one it starts in: ``2025`` is
«2025-26». One :data:`OPEN` year at a time is editable; every other year is
:data:`LOCKED` history.

**A new year is a copy, not a replay.** The μητρώα tables compute an open year
as baseline + accepted proposals and store nothing until it is finalised; that
machinery exists because 52 subjects are worked on by many members with no
owner. A book list has one teacher per course, so :func:`open_year` copies the
baseline outright and people edit their own courses directly. What the
coordinator reviews is the *diff* against the baseline
(:func:`changes_vs_baseline`), which cannot go stale, and
:data:`CHANGES_TABLE` records who did each edit.

Book metadata is **not** in ``eudoxus_selections``: a book appears in several
courses (222 distinct books over 291 rows) and its availability changes on
Eudoxus' side, not ours. It lives once in :data:`BOOKS_TABLE`, refreshed by the
availability check.
"""

from __future__ import annotations

import pandas as pd
from sqlalchemy import text

import db

YEARS_TABLE = "eudoxus_years"
COURSES_TABLE = "eudoxus_courses"
SELECTIONS_TABLE = "eudoxus_selections"
BOOKS_TABLE = "eudoxus_books"
CHANGES_TABLE = "eudoxus_changes"

OPEN = "ΑΝΟΙΧΤΟ"
LOCKED = "ΚΛΕΙΔΩΜΕΝΟ"

ADD = "ΠΡΟΣΘΗΚΗ"
REMOVE = "ΑΦΑΙΡΕΣΗ"
REORDER = "ΣΕΙΡΑ"

# The workbook's «Περίοδος» values, kept verbatim rather than translated: they
# are what the department's own export writes.
WINTER = "Ximerino"
SPRING = "Earino"
PERIOD_LABELS = {WINTER: "Χειμερινό", SPRING: "Εαρινό"}

SCHEMA_SQL = f"""
-- The academic year a list belongs to; 2025 means «2025-26».
CREATE TABLE IF NOT EXISTS {YEARS_TABLE} (
    year          INTEGER PRIMARY KEY,
    status        TEXT NOT NULL CHECK (status IN ('{OPEN}', '{LOCKED}')),
    baseline_year INTEGER,
    opened_by     TEXT,
    opened_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
    locked_by     TEXT,
    locked_at     TIMESTAMPTZ
);

-- One row per course *offering*. ΔΟΜ022 «Οικοδομική ΙΙ» runs in both the 7th
-- and the 9th εξάμηνο — the old programme beside the new one — with the same
-- three books in a different priority order, so the εξάμηνο is part of the key
-- and not an attribute.
CREATE TABLE IF NOT EXISTS {COURSES_TABLE} (
    year        INTEGER NOT NULL,
    course_code TEXT    NOT NULL,
    examino     INTEGER NOT NULL,
    title       TEXT    NOT NULL,
    teacher     TEXT,
    period      TEXT,
    PRIMARY KEY (year, course_code, examino)
);

CREATE TABLE IF NOT EXISTS {SELECTIONS_TABLE} (
    year        INTEGER NOT NULL,
    course_code TEXT    NOT NULL,
    examino     INTEGER NOT NULL,
    book_id     BIGINT  NOT NULL,
    priority    INTEGER NOT NULL,
    source      TEXT    NOT NULL DEFAULT 'EUDOXUS',
    added_by    TEXT,
    added_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (year, course_code, examino, book_id)
);

CREATE INDEX IF NOT EXISTS eudoxus_selections_book_idx
    ON {SELECTIONS_TABLE} (book_id);

-- Catalogue cache, refreshed by the availability check. `found` distinguishes
-- "Eudoxus says this book is withdrawn" from "this code is not in the registry
-- at all" — different problems for whoever has to fix the list.
CREATE TABLE IF NOT EXISTS {BOOKS_TABLE} (
    book_id          BIGINT PRIMARY KEY,
    title            TEXT,
    subtitle         TEXT,
    authors          TEXT,
    isbn             TEXT,
    publisher        TEXT,
    publication_year INTEGER,
    edition_number   TEXT,
    active           BOOLEAN,
    selectable       BOOLEAN,
    found            BOOLEAN NOT NULL DEFAULT TRUE,
    error            TEXT,
    checked_at       TIMESTAMPTZ
);

-- Who changed what in an open year. The coordinator's review is the diff
-- against the baseline; this is the attribution the diff cannot carry.
CREATE TABLE IF NOT EXISTS {CHANGES_TABLE} (
    id          BIGSERIAL PRIMARY KEY,
    year        INTEGER NOT NULL,
    course_code TEXT    NOT NULL,
    examino     INTEGER NOT NULL,
    book_id     BIGINT,
    action      TEXT    NOT NULL CHECK (action IN ('{ADD}', '{REMOVE}', '{REORDER}')),
    detail      TEXT,
    author      TEXT,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS eudoxus_changes_year_idx
    ON {CHANGES_TABLE} (year, created_at DESC);
"""


def year_label(year: int) -> str:
    """2025 -> «2025-26», the way the department writes it."""
    return f"{year}-{str(year + 1)[-2:]}"


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------

def stored_years() -> list[int]:
    """Years that have a book list, newest first."""
    engine = db.get_engine()
    if engine is None:
        return []
    with engine.connect() as conn:
        rows = conn.execute(
            text(f"SELECT DISTINCT year FROM {COURSES_TABLE} ORDER BY year DESC")
        )
        return [row[0] for row in rows]


def year_state(year: int) -> dict | None:
    engine = db.get_engine()
    if engine is None:
        return None
    with engine.connect() as conn:
        row = (
            conn.execute(
                text(f"SELECT * FROM {YEARS_TABLE} WHERE year = :year"), {"year": year}
            )
            .mappings()
            .first()
        )
    return dict(row) if row else None


def open_years() -> list[int]:
    engine = db.get_engine()
    if engine is None:
        return []
    with engine.connect() as conn:
        rows = conn.execute(
            text(f"SELECT year FROM {YEARS_TABLE} WHERE status = :status ORDER BY year"),
            {"status": OPEN},
        )
        return [row[0] for row in rows]


def is_editable(year: int) -> bool:
    state = year_state(year)
    return bool(state and state["status"] == OPEN)


def load_year(year: int) -> pd.DataFrame:
    """The whole list for one year: course offering, book, and what we know of it.

    A LEFT JOIN on the catalogue on purpose — a book whose metadata has never
    been fetched still has to appear, with its code and empty columns, rather
    than vanish from the list.
    """
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    query = text(
        f"""
        SELECT c.course_code, c.examino, c.title AS course_title, c.teacher, c.period,
               s.book_id, s.priority, s.source, s.added_by, s.added_at,
               b.title AS book_title, b.subtitle, b.authors, b.isbn, b.publisher,
               b.publication_year, b.edition_number, b.active, b.selectable,
               b.found, b.error, b.checked_at
        FROM {COURSES_TABLE} c
        JOIN {SELECTIONS_TABLE} s
          ON s.year = c.year AND s.course_code = c.course_code AND s.examino = c.examino
        LEFT JOIN {BOOKS_TABLE} b ON b.book_id = s.book_id
        WHERE c.year = :year
        ORDER BY c.examino, c.course_code, s.priority
        """
    )
    with engine.connect() as conn:
        return pd.read_sql(query, conn, params={"year": year})


def courses_for_year(year: int) -> pd.DataFrame:
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    query = text(
        f"""
        SELECT c.course_code, c.examino, c.title, c.teacher, c.period,
               count(s.book_id) AS books
        FROM {COURSES_TABLE} c
        LEFT JOIN {SELECTIONS_TABLE} s
          ON s.year = c.year AND s.course_code = c.course_code AND s.examino = c.examino
        WHERE c.year = :year
        GROUP BY c.course_code, c.examino, c.title, c.teacher, c.period
        ORDER BY c.examino, c.course_code
        """
    )
    with engine.connect() as conn:
        return pd.read_sql(query, conn, params={"year": year})


def book_ids_for_year(year: int) -> list[int]:
    engine = db.get_engine()
    if engine is None:
        return []
    with engine.connect() as conn:
        rows = conn.execute(
            text(
                f"SELECT DISTINCT book_id FROM {SELECTIONS_TABLE} "
                "WHERE year = :year ORDER BY book_id"
            ),
            {"year": year},
        )
        return [row[0] for row in rows]


def load_books(book_ids: list[int] | None = None) -> pd.DataFrame:
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    if book_ids is None:
        query = text(f"SELECT * FROM {BOOKS_TABLE} ORDER BY book_id")
        params: dict = {}
    else:
        if not book_ids:
            return pd.DataFrame()
        query = text(f"SELECT * FROM {BOOKS_TABLE} WHERE book_id = ANY(:ids) ORDER BY book_id")
        params = {"ids": [int(book_id) for book_id in book_ids]}
    with engine.connect() as conn:
        return pd.read_sql(query, conn, params=params)


def year_changes(year: int) -> pd.DataFrame:
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    query = text(
        f"SELECT * FROM {CHANGES_TABLE} WHERE year = :year ORDER BY created_at DESC"
    )
    with engine.connect() as conn:
        return pd.read_sql(query, conn, params={"year": year})


# --------------------------------------------------------------------------
# The open year
# --------------------------------------------------------------------------

def open_year(year: int, baseline_year: int, opened_by: str) -> str:
    """Start a year as a copy of an existing one.

    The copy is the point: next year's list starts as last year's and is edited
    down. Courses and selections are copied in one transaction with the year
    row, so there is no window where an open year exists with no list.
    """
    engine = db.get_engine()
    if engine is None:
        return "Δεν υπάρχει βάση δεδομένων."
    if year_state(year):
        return f"Το έτος {year_label(year)} έχει ήδη ανοίξει."
    if baseline_year not in stored_years():
        return f"Το έτος βάσης {year_label(baseline_year)} δεν υπάρχει στη βάση."
    already_open = open_years()
    if already_open:
        return (
            "Υπάρχει ήδη ανοιχτό έτος: "
            + ", ".join(year_label(item) for item in already_open)
        )

    with engine.begin() as conn:
        conn.execute(
            text(
                f"INSERT INTO {YEARS_TABLE} (year, status, baseline_year, opened_by) "
                "VALUES (:year, :status, :baseline, :by)"
            ),
            {"year": year, "status": OPEN, "baseline": baseline_year, "by": opened_by},
        )
        conn.execute(
            text(
                f"INSERT INTO {COURSES_TABLE} (year, course_code, examino, title, teacher, period) "
                f"SELECT :year, course_code, examino, title, teacher, period "
                f"FROM {COURSES_TABLE} WHERE year = :baseline"
            ),
            {"year": year, "baseline": baseline_year},
        )
        result = conn.execute(
            text(
                f"INSERT INTO {SELECTIONS_TABLE} "
                "(year, course_code, examino, book_id, priority, source, added_by) "
                "SELECT :year, course_code, examino, book_id, priority, source, :by "
                f"FROM {SELECTIONS_TABLE} WHERE year = :baseline"
            ),
            {"year": year, "baseline": baseline_year, "by": opened_by},
        )
    return (
        f"Το έτος {year_label(year)} άνοιξε με βάση το {year_label(baseline_year)} "
        f"({result.rowcount} επιλογές)."
    )


def lock_year(year: int, locked_by: str) -> str:
    """Close the year to further editing."""
    engine = db.get_engine()
    if engine is None:
        return "Δεν υπάρχει βάση δεδομένων."
    state = year_state(year)
    if state is None:
        return f"Το έτος {year_label(year)} δεν έχει ανοίξει."
    if state["status"] == LOCKED:
        return f"Το έτος {year_label(year)} είναι ήδη κλειδωμένο."
    with engine.begin() as conn:
        conn.execute(
            text(
                f"UPDATE {YEARS_TABLE} SET status = :status, locked_by = :by, "
                "locked_at = now() WHERE year = :year"
            ),
            {"status": LOCKED, "by": locked_by, "year": year},
        )
    return f"Το έτος {year_label(year)} κλείδωσε."


def _log(conn, year: int, course_code: str, examino: int, book_id: int | None,
         action: str, detail: str, author: str) -> None:
    conn.execute(
        text(
            f"INSERT INTO {CHANGES_TABLE} "
            "(year, course_code, examino, book_id, action, detail, author) "
            "VALUES (:year, :code, :examino, :book, :action, :detail, :author)"
        ),
        {
            "year": year,
            "code": course_code,
            "examino": examino,
            "book": book_id,
            "action": action,
            "detail": detail,
            "author": author,
        },
    )


def add_book(year: int, course_code: str, examino: int, book_id: int,
             priority: int, author: str) -> tuple[bool, str]:
    """Add one book to a course of the open year."""
    engine = db.get_engine()
    if engine is None:
        return False, "Δεν υπάρχει βάση δεδομένων."
    if not is_editable(year):
        return False, f"Το έτος {year_label(year)} δεν είναι ανοιχτό για επεξεργασία."
    with engine.begin() as conn:
        existing = conn.execute(
            text(
                f"SELECT 1 FROM {SELECTIONS_TABLE} WHERE year = :year "
                "AND course_code = :code AND examino = :examino AND book_id = :book"
            ),
            {"year": year, "code": course_code, "examino": examino, "book": book_id},
        ).first()
        if existing:
            return False, "Το βιβλίο υπάρχει ήδη στο μάθημα."
        conn.execute(
            text(
                f"INSERT INTO {SELECTIONS_TABLE} "
                "(year, course_code, examino, book_id, priority, added_by) "
                "VALUES (:year, :code, :examino, :book, :priority, :by)"
            ),
            {
                "year": year,
                "code": course_code,
                "examino": examino,
                "book": book_id,
                "priority": priority,
                "by": author,
            },
        )
        _log(conn, year, course_code, examino, book_id, ADD,
             f"σειρά επιλογής {priority}", author)
    return True, f"Προστέθηκε το βιβλίο {book_id}."


def remove_book(year: int, course_code: str, examino: int, book_id: int,
                author: str, reason: str = "") -> tuple[bool, str]:
    engine = db.get_engine()
    if engine is None:
        return False, "Δεν υπάρχει βάση δεδομένων."
    if not is_editable(year):
        return False, f"Το έτος {year_label(year)} δεν είναι ανοιχτό για επεξεργασία."
    with engine.begin() as conn:
        result = conn.execute(
            text(
                f"DELETE FROM {SELECTIONS_TABLE} WHERE year = :year "
                "AND course_code = :code AND examino = :examino AND book_id = :book"
            ),
            {"year": year, "code": course_code, "examino": examino, "book": book_id},
        )
        if result.rowcount == 0:
            return False, "Το βιβλίο δεν βρέθηκε στο μάθημα."
        _log(conn, year, course_code, examino, book_id, REMOVE, reason, author)
    return True, f"Αφαιρέθηκε το βιβλίο {book_id}."


def set_priorities(year: int, course_code: str, examino: int,
                   priorities: dict[int, int], author: str) -> tuple[bool, str]:
    """Rewrite the priority of every book of one course in a single transaction.

    Whole-course rather than per-book: the σειρά επιλογής only means anything as
    an ordering, and applying half of a reordering leaves two books sharing a
    position.
    """
    engine = db.get_engine()
    if engine is None:
        return False, "Δεν υπάρχει βάση δεδομένων."
    if not is_editable(year):
        return False, f"Το έτος {year_label(year)} δεν είναι ανοιχτό για επεξεργασία."
    if not priorities:
        return False, "Καμία αλλαγή."
    changed: list[str] = []
    with engine.begin() as conn:
        current = {
            row[0]: row[1]
            for row in conn.execute(
                text(
                    f"SELECT book_id, priority FROM {SELECTIONS_TABLE} WHERE year = :year "
                    "AND course_code = :code AND examino = :examino"
                ),
                {"year": year, "code": course_code, "examino": examino},
            )
        }
        for book_id, priority in priorities.items():
            if current.get(book_id) == priority:
                continue
            conn.execute(
                text(
                    f"UPDATE {SELECTIONS_TABLE} SET priority = :priority WHERE year = :year "
                    "AND course_code = :code AND examino = :examino AND book_id = :book"
                ),
                {
                    "priority": priority,
                    "year": year,
                    "code": course_code,
                    "examino": examino,
                    "book": book_id,
                },
            )
            changed.append(f"{book_id}: {current.get(book_id)} → {priority}")
        if changed:
            _log(conn, year, course_code, examino, None, REORDER,
                 "· ".join(changed), author)
    if not changed:
        return False, "Καμία αλλαγή στη σειρά επιλογής."
    return True, f"Ενημερώθηκε η σειρά επιλογής ({len(changed)} βιβλία)."


# --------------------------------------------------------------------------
# The catalogue
# --------------------------------------------------------------------------

UPSERT_BOOK_SQL = f"""
INSERT INTO {BOOKS_TABLE}
    (book_id, title, subtitle, authors, isbn, publisher, publication_year,
     edition_number, active, selectable, found, error, checked_at)
VALUES
    (:book_id, :title, :subtitle, :authors, :isbn, :publisher, :publication_year,
     :edition_number, :active, :selectable, :found, :error, now())
ON CONFLICT (book_id) DO UPDATE SET
    title = EXCLUDED.title,
    subtitle = EXCLUDED.subtitle,
    authors = EXCLUDED.authors,
    isbn = EXCLUDED.isbn,
    publisher = EXCLUDED.publisher,
    publication_year = EXCLUDED.publication_year,
    edition_number = EXCLUDED.edition_number,
    active = EXCLUDED.active,
    selectable = EXCLUDED.selectable,
    found = EXCLUDED.found,
    error = EXCLUDED.error,
    checked_at = EXCLUDED.checked_at
"""

_BOOK_KEYS = (
    "book_id", "title", "subtitle", "authors", "isbn", "publisher",
    "publication_year", "edition_number", "active", "selectable", "found", "error",
)


# Columns the catalogue stores as text. ISBNs in particular arrive as integers
# — both from the Eudoxus JSON and from pandas reading the dump — and Postgres
# has no assignment cast from bigint to text, so an uncoerced one fails the
# INSERT outright. numpy scalars from a DataFrame are the same class of
# problem: psycopg cannot adapt them at all.
_TEXT_KEYS = ("title", "subtitle", "authors", "isbn", "publisher", "edition_number", "error")


def _clean_book(row: dict) -> dict:
    """One catalogue row with every column present and typed for Postgres."""
    cleaned = {key: row.get(key) for key in _BOOK_KEYS}
    cleaned["book_id"] = int(cleaned["book_id"])
    # `found` defaults to True, but the dict comprehension above has already
    # put a None here, so a .get default would never fire — and a book wrongly
    # marked not-found reads as "withdrawn from the registry".
    cleaned["found"] = True if cleaned["found"] is None else bool(cleaned["found"])
    year = cleaned.get("publication_year")
    if year is None:
        cleaned["publication_year"] = None
    else:
        try:
            cleaned["publication_year"] = int(float(str(year)))
        except (TypeError, ValueError):
            cleaned["publication_year"] = None
    for key in _TEXT_KEYS:
        value = cleaned.get(key)
        cleaned[key] = None if value is None else str(value)
    for flag in ("active", "selectable"):
        value = cleaned.get(flag)
        cleaned[flag] = None if value is None else bool(value)
    return cleaned


def upsert_books(rows: list[dict]) -> int:
    """Store catalogue rows, replacing what was there. Returns rows written."""
    engine = db.get_engine()
    if engine is None or not rows:
        return 0
    payload = [_clean_book(row) for row in rows]
    with engine.begin() as conn:
        conn.execute(text(UPSERT_BOOK_SQL), payload)
    return len(payload)


def unavailable_for_year(year: int) -> pd.DataFrame:
    """The rows of a year whose book cannot be chosen again, with the reason.

    A book that has never been checked is *not* reported as a problem — absence
    of an answer is not a negative one. The page counts those separately and
    says so.
    """
    frame = load_year(year)
    if frame.empty:
        return frame
    checked = frame[frame["checked_at"].notna()]
    if checked.empty:
        return checked
    problems = checked[
        (~checked["found"].fillna(False))
        | (~checked["active"].fillna(False))
        | (~checked["selectable"].fillna(False))
    ].copy()
    problems["reason"] = problems.apply(
        lambda row: (
            row["error"] or "δεν βρέθηκε στο μητρώο"
            if not row["found"]
            else "ανενεργό (active=false)"
            if not row["active"]
            else "μη επιλέξιμο (selectable=false)"
        ),
        axis=1,
    )
    return problems


def changes_vs_baseline(year: int) -> pd.DataFrame:
    """What the open year holds that its baseline did not, and the reverse.

    Computed rather than read from the log: the log records every edit,
    including ones that cancelled each other out, and what the coordinator has
    to approve is the net difference.
    """
    state = year_state(year)
    if state is None or state.get("baseline_year") is None:
        return pd.DataFrame()
    baseline = state["baseline_year"]
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    query = text(
        f"""
        SELECT COALESCE(n.course_code, o.course_code) AS course_code,
               COALESCE(n.examino, o.examino)         AS examino,
               COALESCE(n.book_id, o.book_id)         AS book_id,
               o.priority AS priority_before,
               n.priority AS priority_after,
               CASE WHEN o.book_id IS NULL THEN '{ADD}'
                    WHEN n.book_id IS NULL THEN '{REMOVE}'
                    ELSE '{REORDER}' END AS action,
               b.title AS book_title, b.authors, b.active, b.selectable, b.found
        FROM (SELECT * FROM {SELECTIONS_TABLE} WHERE year = :year) n
        FULL OUTER JOIN
             (SELECT * FROM {SELECTIONS_TABLE} WHERE year = :baseline) o
          ON  n.course_code = o.course_code
          AND n.examino = o.examino
          AND n.book_id = o.book_id
        LEFT JOIN {BOOKS_TABLE} b ON b.book_id = COALESCE(n.book_id, o.book_id)
        WHERE o.book_id IS NULL
           OR n.book_id IS NULL
           OR n.priority IS DISTINCT FROM o.priority
        ORDER BY 2, 1, 3
        """
    )
    with engine.connect() as conn:
        return pd.read_sql(query, conn, params={"year": year, "baseline": baseline})
