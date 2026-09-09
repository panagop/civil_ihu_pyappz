"""Postgres access for the περιγράμματα μαθημάτων.

The Google Sheet that held these until now is **not** a runtime source: it was
exported once into ``files/perigrammata/*.csv`` (the archive of what it held at
handover) and the database is the master from then on. Nothing here ever
touches the network.

Unlike the μητρώα tables there is no proposals workflow: a περίγραμμα has one
natural owner per course, so anyone who passes the login gate edits directly
and the revisions table keeps the full history. That history is also what the
coordinator's "changes over a period" report reads — which is why every save
records both the new state and the fields that moved.

Everything degrades to "no database" rather than crashing, but page 6 refuses
to render stale data: with no ``DATABASE_URL`` it stops instead of falling back
to the archive CSVs.
"""

from __future__ import annotations

import json
from datetime import datetime

import pandas as pd
from sqlalchemy import text

import db

COURSES_TABLE = "perigrammata_courses"
REVISIONS_TABLE = "perigrammata_revisions"

# The 2018 πρόγραμμα σπουδών is finished: stored, viewable, printable, but not
# editable. Editing a closed curriculum only invites drift between the document
# a student was given and the one the app prints.
CURRICULA = (2018, 2025)
EDITABLE_CURRICULUM = 2025
LOCALES = ("gr", "eng")
LOCALE_LABELS = {"gr": "Ελληνικά", "eng": "Αγγλικά"}

# The identity of a περίγραμμα. `code` is unique within a (curriculum, locale),
# but 84 codes are shared between 2018 and 2025 carrying different content, so
# the curriculum has to be part of the key.
KEY_COLUMNS = ("curriculum", "locale", "code")

# Content columns, in the order the sheet had them — which is the order the
# form shows and the report prints. `lang` here is the *course* language
# («Ελληνική»), not the locale of the περίγραμμα; that is why the key column is
# called `locale` and not `lang`.
CONTENT_COLUMNS: dict[str, str] = {
    "name": "TEXT",
    "school": "TEXT",
    "department": "TEXT",
    "level": "TEXT",
    "examino": "INTEGER",
    "didactivities_name1": "TEXT",
    "didactivities_hours1": "NUMERIC",
    "didactivities_ects1": "NUMERIC",
    "didactivities_name2": "TEXT",
    "didactivities_hours2": "NUMERIC",
    "didactivities_ects2": "NUMERIC",
    "type": "TEXT",
    "prereq": "TEXT",
    "prereq_knowledge": "TEXT",
    "lang": "TEXT",
    "erasmus": "TEXT",
    "website": "TEXT",
    "learn_objectives": "TEXT",
    "skills": "TEXT",
    "subject1": "TEXT",
    "subject2": "TEXT",
    "subject3": "TEXT",
    "teaching_way": "TEXT",
    "techs_used": "TEXT",
    "task1": "TEXT",
    "hours1": "NUMERIC",
    "task2": "TEXT",
    "hours2": "NUMERIC",
    "task3": "TEXT",
    "hours3": "NUMERIC",
    "task4": "TEXT",
    "hours4": "NUMERIC",
    "task5": "TEXT",
    "hours5": "NUMERIC",
    "hours_sum": "NUMERIC",
    "grading": "TEXT",
    "refs_books": "TEXT",
    "refs_more_title": "TEXT",
    "refs_more": "TEXT",
}

NUMERIC_COLUMNS = {n for n, t in CONTENT_COLUMNS.items() if t == "NUMERIC"}
INTEGER_COLUMNS = {n for n, t in CONTENT_COLUMNS.items() if t == "INTEGER"}

# refs_more_title / refs_more are in every worksheet but in neither Word
# template. They are stored and editable, and simply do not print until a
# template gains them — dropping them on import would have been the worse
# choice.
UNPRINTED_COLUMNS = ("refs_more_title", "refs_more")

# Greek labels and the widget each field wants. Kept beside the column spec so
# a column cannot be added to the table without deciding how it is edited.
FIELD_GROUPS: list[tuple[str, list[tuple[str, str, str]]]] = [
    (
        "Γενικά",
        [
            ("name", "Τίτλος μαθήματος", "text"),
            ("school", "Σχολή", "text"),
            ("department", "Τμήμα", "text"),
            ("level", "Επίπεδο σπουδών", "text"),
            ("examino", "Εξάμηνο σπουδών", "int"),
            ("type", "Τύπος μαθήματος", "text"),
            ("lang", "Γλώσσα διδασκαλίας και εξετάσεων", "text"),
            ("erasmus", "Προσφέρεται σε φοιτητές Erasmus", "text"),
            ("website", "Ηλεκτρονική σελίδα μαθήματος", "text"),
            ("prereq", "Προαπαιτούμενα μαθήματα", "area"),
            ("prereq_knowledge", "Προαπαιτούμενες γνώσεις", "area"),
        ],
    ),
    (
        "Διδακτικές δραστηριότητες",
        [
            ("didactivities_name1", "Δραστηριότητα 1", "text"),
            ("didactivities_hours1", "Εβδομαδιαίες ώρες 1", "num"),
            ("didactivities_ects1", "Πιστωτικές μονάδες 1", "num"),
            ("didactivities_name2", "Δραστηριότητα 2", "text"),
            ("didactivities_hours2", "Εβδομαδιαίες ώρες 2", "num"),
            ("didactivities_ects2", "Πιστωτικές μονάδες 2", "num"),
        ],
    ),
    (
        "Μαθησιακά αποτελέσματα",
        [
            ("learn_objectives", "Μαθησιακά αποτελέσματα", "area"),
            ("skills", "Γενικές ικανότητες", "area"),
        ],
    ),
    (
        "Περιεχόμενο μαθήματος",
        [
            ("subject1", "Περιεχόμενο (1)", "area"),
            ("subject2", "Περιεχόμενο (2)", "area"),
            ("subject3", "Περιεχόμενο (3)", "area"),
        ],
    ),
    (
        "Διδακτικές μέθοδοι",
        [
            ("teaching_way", "Τρόπος παράδοσης", "area"),
            ("techs_used", "Χρήση ΤΠΕ", "area"),
        ],
    ),
    (
        "Οργάνωση διδασκαλίας",
        [
            ("task1", "Δραστηριότητα 1", "text"),
            ("hours1", "Φόρτος εργασίας 1", "num"),
            ("task2", "Δραστηριότητα 2", "text"),
            ("hours2", "Φόρτος εργασίας 2", "num"),
            ("task3", "Δραστηριότητα 3", "text"),
            ("hours3", "Φόρτος εργασίας 3", "num"),
            ("task4", "Δραστηριότητα 4", "text"),
            ("hours4", "Φόρτος εργασίας 4", "num"),
            ("task5", "Δραστηριότητα 5", "text"),
            ("hours5", "Φόρτος εργασίας 5", "num"),
            ("hours_sum", "Σύνολο φόρτου εργασίας", "num"),
        ],
    ),
    (
        "Αξιολόγηση και βιβλιογραφία",
        [
            ("grading", "Αξιολόγηση φοιτητών", "area"),
            ("refs_books", "Συνιστώμενη βιβλιογραφία", "area"),
            ("refs_more_title", "Συναφή περιοδικά — τίτλος", "text"),
            ("refs_more", "Συναφή περιοδικά", "area"),
        ],
    ),
]

FIELD_LABELS = {name: label for _, fields in FIELD_GROUPS for name, label, _ in fields}

# What hours_sum is supposed to be the sum of. Not a database CHECK: the
# workbooks disagree with themselves often enough that a constraint would
# refuse rows we actually have. The form warns instead.
WORKLOAD_COLUMNS = ("hours1", "hours2", "hours3", "hours4", "hours5")


def _schema_sql() -> str:
    """DDL built from CONTENT_COLUMNS, so the table and the spec cannot drift."""
    columns = ",\n    ".join(
        f"{name:<22} {sql_type}" for name, sql_type in CONTENT_COLUMNS.items()
    )
    empty_json = "'{}'::jsonb"
    return f"""
CREATE TABLE IF NOT EXISTS {COURSES_TABLE} (
    curriculum             INTEGER NOT NULL,
    locale                 TEXT    NOT NULL CHECK (locale IN ('gr', 'eng')),
    code                   TEXT    NOT NULL,
    sort_order             INTEGER,
    {columns},
    updated_at             TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by             TEXT,
    PRIMARY KEY (curriculum, locale, code)
);

-- Append-only history. `data` is the full new state (so a course can be rolled
-- back), `changes` only the fields that moved (so the coordinator's report is a
-- read rather than a diff of consecutive snapshots).
CREATE TABLE IF NOT EXISTS {REVISIONS_TABLE} (
    id         BIGSERIAL PRIMARY KEY,
    curriculum INTEGER NOT NULL,
    locale     TEXT    NOT NULL,
    code       TEXT    NOT NULL,
    data       JSONB   NOT NULL,
    changes    JSONB   NOT NULL DEFAULT {empty_json},
    edited_by  TEXT,
    edited_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    note       TEXT
);

CREATE INDEX IF NOT EXISTS perigrammata_revisions_course_idx
    ON {REVISIONS_TABLE} (curriculum, locale, code, edited_at DESC);
-- The changes report scans by date across every course.
CREATE INDEX IF NOT EXISTS perigrammata_revisions_edited_at_idx
    ON {REVISIONS_TABLE} (edited_at DESC);
"""


SCHEMA_SQL = _schema_sql()


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------

def stored_curricula(locale: str = "gr") -> list[int]:
    """Curricula that actually have rows, newest first."""
    engine = db.get_engine()
    if engine is None:
        return []
    with engine.connect() as conn:
        rows = conn.execute(
            text(
                f"SELECT DISTINCT curriculum FROM {COURSES_TABLE} "
                "WHERE locale = :locale ORDER BY curriculum DESC"
            ),
            {"locale": locale},
        )
        return [row[0] for row in rows]


def load_courses(curriculum: int, locale: str = "gr") -> pd.DataFrame:
    """Every course of one curriculum, in the order the sheet had them."""
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    query = text(
        f"SELECT * FROM {COURSES_TABLE} "
        "WHERE curriculum = :curriculum AND locale = :locale "
        "ORDER BY sort_order, code"
    )
    with engine.connect() as conn:
        return pd.read_sql(
            query, conn, params={"curriculum": curriculum, "locale": locale}
        )


def load_course(curriculum: int, locale: str, code: str) -> dict | None:
    """One course as a plain dict, or None when it does not exist."""
    engine = db.get_engine()
    if engine is None:
        return None
    with engine.connect() as conn:
        row = (
            conn.execute(
                text(
                    f"SELECT * FROM {COURSES_TABLE} "
                    "WHERE curriculum = :curriculum AND locale = :locale "
                    "AND code = :code"
                ),
                {"curriculum": curriculum, "locale": locale, "code": code},
            )
            .mappings()
            .first()
        )
    return dict(row) if row else None


def course_history(curriculum: int, locale: str, code: str) -> pd.DataFrame:
    """Revisions of one course, newest first."""
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    query = text(
        f"SELECT id, edited_at, edited_by, note, changes FROM {REVISIONS_TABLE} "
        "WHERE curriculum = :curriculum AND locale = :locale AND code = :code "
        "ORDER BY edited_at DESC"
    )
    with engine.connect() as conn:
        return pd.read_sql(
            query,
            conn,
            params={"curriculum": curriculum, "locale": locale, "code": code},
        )


def revision_data(revision_id: int) -> dict | None:
    """The full stored state of one revision — what a rollback would restore."""
    engine = db.get_engine()
    if engine is None:
        return None
    with engine.connect() as conn:
        row = conn.execute(
            text(f"SELECT data FROM {REVISIONS_TABLE} WHERE id = :id"),
            {"id": revision_id},
        ).first()
    if row is None:
        return None
    return row[0] if isinstance(row[0], dict) else json.loads(row[0])


def changes_between(
    start: datetime, end: datetime, curriculum: int | None = None, locale: str = "gr"
) -> pd.DataFrame:
    """Every revision in a period, one row per changed field.

    The seed revisions carry an empty ``changes`` object, so the initial import
    never shows up as 199 courses "changed" on the day the database was filled.
    """
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    clauses = ["edited_at >= :start", "edited_at < :end", "locale = :locale"]
    params: dict[str, object] = {"start": start, "end": end, "locale": locale}
    if curriculum is not None:
        clauses.append("curriculum = :curriculum")
        params["curriculum"] = curriculum
    query = text(
        f"SELECT id, curriculum, code, edited_at, edited_by, note, changes "
        f"FROM {REVISIONS_TABLE} WHERE {' AND '.join(clauses)} "
        "ORDER BY edited_at"
    )
    with engine.connect() as conn:
        revisions = pd.read_sql(query, conn, params=params)

    records: list[dict] = []
    for _, revision in revisions.iterrows():
        changes = revision["changes"]
        if isinstance(changes, str):
            changes = json.loads(changes)
        for field, pair in (changes or {}).items():
            before, after = (pair + [None, None])[:2] if isinstance(pair, list) else (None, None)
            records.append(
                {
                    "curriculum": revision["curriculum"],
                    "code": revision["code"],
                    "field": field,
                    "label": FIELD_LABELS.get(field, field),
                    "before": before,
                    "after": after,
                    "edited_at": revision["edited_at"],
                    "edited_by": revision["edited_by"],
                    "note": revision["note"],
                }
            )
    return pd.DataFrame(records)


# --------------------------------------------------------------------------
# Writing
# --------------------------------------------------------------------------

def _normalize(field: str, value: object) -> object:
    """Coerce a form value to what the column stores, or None when blank."""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if field in NUMERIC_COLUMNS or field in INTEGER_COLUMNS:
        if isinstance(value, str):
            value = value.strip().replace(",", ".")
            if not value:
                return None
            value = float(value)
        if value is None:
            return None
        return int(value) if field in INTEGER_COLUMNS else float(value)
    text_value = str(value).strip()
    return text_value or None


def diff_course(before: dict, after: dict) -> dict[str, list]:
    """The content fields that actually moved, as ``{field: [before, after]}``."""
    changes: dict[str, list] = {}
    for field in CONTENT_COLUMNS:
        old = _normalize(field, before.get(field))
        new = _normalize(field, after.get(field))
        # Numerics come back from Postgres as Decimal; compare as float so
        # Decimal('4') and 4.0 do not read as a change on every save.
        if field in NUMERIC_COLUMNS or field in INTEGER_COLUMNS:
            old = None if old is None else float(old)
            new = None if new is None else float(new)
        if old != new:
            changes[field] = [old, new]
    return changes


def save_course(
    curriculum: int,
    locale: str,
    code: str,
    values: dict,
    editor: str,
    loaded_at: datetime,
    note: str = "",
) -> tuple[bool, str]:
    """Write one course and append a revision, in a single transaction.

    ``loaded_at`` is the ``updated_at`` the form was rendered from. The UPDATE
    matches on it, so a second browser tab that loaded the same course earlier
    fails loudly instead of silently overwriting work: Streamlit locks nothing
    and two people on the same course is not a rare accident.

    Returns ``(saved, message)``; ``saved`` is False both for a conflict and for
    a form that changed nothing.
    """
    engine = db.get_engine()
    if engine is None:
        return False, "Δεν υπάρχει βάση δεδομένων."
    if curriculum != EDITABLE_CURRICULUM:
        return False, f"Το πρόγραμμα σπουδών {curriculum} δεν είναι επεξεργάσιμο."

    current = load_course(curriculum, locale, code)
    if current is None:
        return False, f"Το μάθημα {code} δεν βρέθηκε."

    cleaned = {field: _normalize(field, values.get(field)) for field in CONTENT_COLUMNS}
    changes = diff_course(current, cleaned)
    if not changes:
        return False, "Δεν εντοπίστηκε καμία αλλαγή."

    assignments = ", ".join(f"{field} = :{field}" for field in CONTENT_COLUMNS)
    params = dict(cleaned)
    params.update(
        {
            "curriculum": curriculum,
            "locale": locale,
            "code": code,
            "editor": editor,
            "loaded_at": loaded_at,
        }
    )

    with engine.begin() as conn:
        result = conn.execute(
            text(
                f"UPDATE {COURSES_TABLE} SET {assignments}, "
                "updated_at = now(), updated_by = :editor "
                "WHERE curriculum = :curriculum AND locale = :locale "
                "AND code = :code AND updated_at = :loaded_at"
            ),
            params,
        )
        if result.rowcount == 0:
            return False, (
                "Το μάθημα άλλαξε από κάποιον άλλον μετά το άνοιγμα της φόρμας. "
                "Φόρτωσέ το ξανά και ξαναπέρασε τις αλλαγές σου."
            )
        stored = load_course_in(conn, curriculum, locale, code)
        conn.execute(
            text(
                f"INSERT INTO {REVISIONS_TABLE} "
                "(curriculum, locale, code, data, changes, edited_by, note) "
                "VALUES (:curriculum, :locale, :code, "
                "CAST(:data AS jsonb), CAST(:changes AS jsonb), :editor, :note)"
            ),
            {
                "curriculum": curriculum,
                "locale": locale,
                "code": code,
                "data": json.dumps(stored, ensure_ascii=False, default=str),
                "changes": json.dumps(changes, ensure_ascii=False, default=str),
                "editor": editor,
                "note": note.strip() or None,
            },
        )

    labels = ", ".join(FIELD_LABELS.get(f, f) for f in changes)
    return True, f"Αποθηκεύτηκαν {len(changes)} αλλαγές: {labels}"


def load_course_in(conn, curriculum: int, locale: str, code: str) -> dict:
    """The stored row, read inside an open transaction.

    Separate from :func:`load_course` because the revision has to record what
    the database actually holds after the UPDATE — including ``updated_at`` —
    and reading it on another connection would not see the uncommitted row.
    """
    row = (
        conn.execute(
            text(
                f"SELECT * FROM {COURSES_TABLE} "
                "WHERE curriculum = :curriculum AND locale = :locale AND code = :code"
            ),
            {"curriculum": curriculum, "locale": locale, "code": code},
        )
        .mappings()
        .first()
    )
    return dict(row) if row else {}
