"""Εβδομαδιαίο πρόγραμμα — schema, terms, staff, rooms, classes and conflicts.

A *term* is one semester of one academic year: ``(2026, 'Χειμερινό')`` is the
winter of 2026-27. A term is opened as a **copy** of the same period one year
earlier (the winter timetable resembles last winter's far more than last
spring's) and then rearranged by a coordinator; it is locked when published.
Same model as Εύδοξος: there is one owner (the coordinator), so no proposals
machinery.

Tables (all ``timetable_*`` so an admin's ``\\dt`` shows them as one group):

- ``timetable_staff`` — every person who may appear in a timetable, from the
  department site plus the former members who still teach. *Staff*, not
  *instructor*: the site lists technicians who never teach, and "instructor"
  is the role a person has on one class, not what the person is.
- ``timetable_staff_terms`` — whether a person is active in a given term. One
  row per person per term rather than boolean columns on the staff row (a
  column per semester needs a migration every term) or a from/to range
  (έκτακτοι are hired per semester and come and go). No row means *unknown*.
- ``timetable_rooms`` — the rooms, with a short code the timetable prints.
- ``timetable_classes`` — one row per meeting, what a workbook row was: a
  course section (Θ, Ε1, Φ…) placed on a day and hour. ``day IS NULL`` is a
  course listed for the term and not yet placed — the arranging tab works
  through those. Instructors and rooms are link tables because both are
  already many-to-one in the data (two instructors on ΓΕΝ002, two rooms on
  ΔΟΜ011).
- ``timetable_changes`` — who did what in an open term.
- ``timetable_snapshots`` / ``timetable_snapshot_classes`` — named, frozen
  copies («Επιλογή Α», «Επιλογή Β»…) of the open term, to compare and restore.
  Save points, not parallel drafts: the term itself stays the one working
  timetable that every view shows and every form edits.

A class is a **course code**, not a (programme, code) pair. The department
gave every course new to the 2025 programme a new code and kept the code of
every course that carried over, so the code alone identifies the course.
Which programme an εξάμηνο follows in a given year is deliberately not
modelled (dropped 2026-09-16): during the transition courses may move between
winter and spring, and nobody can say in advance when. Names come from
``perigrammata_courses`` by code, the newest programme that has it winning
(``NEWEST_NAME_SQL``). The one exception is a code whose programmes disagree
on the name — ΣΥΓ017 is «Οργάνωση Εργοταξίου και Δομικές Μηχανές» in 2018 and
«Προγραμματισμός και Διαχείριση Τεχνικών Έργων» in 2025 — where a row may set
``name_curriculum`` to print another programme's title (added 2026-09-23).
Besides that, the only name-ish thing kept here is ``name_suffix``,
the «ΔΥ, ΥΕ» elective-group marker the printed timetable carries after the
course name.
"""

from __future__ import annotations

import re
from itertools import combinations

import pandas as pd
from sqlalchemy import text

import db
from perigrammata_db import COURSES_TABLE as PERIGRAMMATA_TABLE

STAFF_TABLE = "timetable_staff"
STAFF_TERMS_TABLE = "timetable_staff_terms"
ROOMS_TABLE = "timetable_rooms"
TERMS_TABLE = "timetable_terms"
CLASSES_TABLE = "timetable_classes"
CLASS_INSTRUCTORS_TABLE = "timetable_class_instructors"
CLASS_ROOMS_TABLE = "timetable_class_rooms"
CHANGES_TABLE = "timetable_changes"
SNAPSHOTS_TABLE = "timetable_snapshots"
SNAPSHOT_CLASSES_TABLE = "timetable_snapshot_classes"

WINTER = "Χειμερινό"
SPRING = "Εαρινό"
PERIODS = (WINTER, SPRING)

OPEN = "ΑΝΟΙΧΤΟ"
LOCKED = "ΚΛΕΙΔΩΜΕΝΟ"

CATEGORIES = ("ΔΕΠ", "ΕΔΙΠ", "ΕΤΕΠ", "ΕΚΤΑΚΤΟΣ", "ΑΛΛΟ")
ROOM_KINDS = ("ΑΙΘΟΥΣΑ", "ΕΡΓΑΣΤΗΡΙΟ", "ΗΥ")

# Sections: Θ theory, Ε lab (Ε1, Ε2… parallel groups), Φ φροντιστήριο.
SECTION_RE = re.compile(r"^(Θ|Ε[1-9]?|Φ)$")
SECTION_LABELS = {"Θ": "Θεωρία", "Ε": "Εργαστήριο", "Φ": "Φροντιστήριο"}
PARALLEL_LAB_RE = re.compile(r"^Ε[1-9]$")

DAYS = ("Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή")
DAY_NUMBER = {name: index + 1 for index, name in enumerate(DAYS)}
FIRST_HOUR, LAST_HOUR = 8, 21  # the calendar shows 08:00–21:00
MAX_DURATION = 6

ADD, MODIFY, DELETE = "ΠΡΟΣΘΗΚΗ", "ΜΕΤΑΒΟΛΗ", "ΔΙΑΓΡΑΦΗ"
OPENED, LOCKED_ACTION = "ΑΝΟΙΓΜΑ", "ΚΛΕΙΔΩΜΑ"
SNAPSHOT_ACTION, RESTORED_ACTION = "ΕΚΔΟΧΗ", "ΕΠΑΝΑΦΟΡΑ"
ACTIONS = (ADD, MODIFY, DELETE, OPENED, LOCKED_ACTION, SNAPSHOT_ACTION, RESTORED_ACTION)


def period_for(examino: int) -> str:
    return WINTER if examino % 2 == 1 else SPRING


def year_label(year: int) -> str:
    return f"{year}-{str(year + 1)[-2:]}"


def term_label(year: int, period: str) -> str:
    return f"{year_label(year)} {period}"


SCHEMA_SQL = f"""
CREATE TABLE IF NOT EXISTS {STAFF_TABLE} (
    id          SERIAL PRIMARY KEY,
    short_name  TEXT NOT NULL UNIQUE,
    last_name   TEXT NOT NULL,
    first_name  TEXT,
    category    TEXT NOT NULL CHECK (category IN {CATEGORIES}),
    rank        TEXT,
    subject     TEXT,
    email       TEXT,
    website_url TEXT,
    placeholder BOOLEAN NOT NULL DEFAULT FALSE,
    notes       TEXT,
    updated_by  TEXT,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS {STAFF_TERMS_TABLE} (
    staff_id INTEGER NOT NULL REFERENCES {STAFF_TABLE}(id) ON DELETE CASCADE,
    year     INTEGER NOT NULL,
    period   TEXT    NOT NULL CHECK (period IN {PERIODS}),
    active   BOOLEAN NOT NULL DEFAULT TRUE,
    category TEXT CHECK (category IN {CATEGORIES}),
    PRIMARY KEY (staff_id, year, period)
);

CREATE TABLE IF NOT EXISTS {ROOMS_TABLE} (
    code     TEXT PRIMARY KEY,
    name     TEXT NOT NULL,
    kind     TEXT NOT NULL CHECK (kind IN {ROOM_KINDS}),
    capacity INTEGER,
    active   BOOLEAN NOT NULL DEFAULT TRUE,
    notes    TEXT
);

CREATE TABLE IF NOT EXISTS {TERMS_TABLE} (
    year            INTEGER NOT NULL,
    period          TEXT    NOT NULL CHECK (period IN {PERIODS}),
    status          TEXT    NOT NULL CHECK (status IN ('{OPEN}', '{LOCKED}')),
    baseline_year   INTEGER,
    baseline_period TEXT,
    opened_by       TEXT,
    opened_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
    locked_by       TEXT,
    locked_at       TIMESTAMPTZ,
    PRIMARY KEY (year, period)
);

CREATE TABLE IF NOT EXISTS {CLASSES_TABLE} (
    id          BIGSERIAL PRIMARY KEY,
    year        INTEGER NOT NULL,
    period      TEXT    NOT NULL,
    examino     INTEGER NOT NULL CHECK (examino BETWEEN 1 AND 10),
    course_code TEXT    NOT NULL,
    section     TEXT    NOT NULL,
    name_suffix TEXT,
    name_curriculum INTEGER,
    day         INTEGER CHECK (day BETWEEN 1 AND 5),
    start_hour  INTEGER CHECK (start_hour BETWEEN {FIRST_HOUR} AND {LAST_HOUR - 1}),
    duration    INTEGER CHECK (duration BETWEEN 1 AND {MAX_DURATION}),
    notes       TEXT,
    updated_by  TEXT,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    FOREIGN KEY (year, period) REFERENCES {TERMS_TABLE}(year, period),
    -- placed entirely or not at all
    CHECK ((day IS NULL) = (start_hour IS NULL) AND (day IS NULL) = (duration IS NULL)),
    CHECK (start_hour IS NULL OR start_hour + duration <= {LAST_HOUR})
);

CREATE INDEX IF NOT EXISTS timetable_classes_term_idx
    ON {CLASSES_TABLE} (year, period, examino);

CREATE TABLE IF NOT EXISTS {CLASS_INSTRUCTORS_TABLE} (
    class_id BIGINT  NOT NULL REFERENCES {CLASSES_TABLE}(id) ON DELETE CASCADE,
    staff_id INTEGER NOT NULL REFERENCES {STAFF_TABLE}(id),
    position INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (class_id, staff_id)
);

CREATE TABLE IF NOT EXISTS {CLASS_ROOMS_TABLE} (
    class_id  BIGINT NOT NULL REFERENCES {CLASSES_TABLE}(id) ON DELETE CASCADE,
    room_code TEXT   NOT NULL REFERENCES {ROOMS_TABLE}(code),
    PRIMARY KEY (class_id, room_code)
);

CREATE TABLE IF NOT EXISTS {CHANGES_TABLE} (
    id         BIGSERIAL PRIMARY KEY,
    year       INTEGER NOT NULL,
    period     TEXT    NOT NULL,
    class_id   BIGINT,
    action     TEXT    NOT NULL CHECK (action IN {ACTIONS}),
    detail     TEXT,
    author     TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS timetable_changes_term_idx
    ON {CHANGES_TABLE} (year, period, created_at DESC);

-- A saved version of an open term. Instructors and rooms are arrays rather
-- than link tables: a frozen copy is never edited row by row.
CREATE TABLE IF NOT EXISTS {SNAPSHOTS_TABLE} (
    id         BIGSERIAL PRIMARY KEY,
    year       INTEGER NOT NULL,
    period     TEXT    NOT NULL,
    name       TEXT    NOT NULL,
    note       TEXT,
    created_by TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    FOREIGN KEY (year, period) REFERENCES {TERMS_TABLE}(year, period),
    UNIQUE (year, period, name)
);

CREATE TABLE IF NOT EXISTS {SNAPSHOT_CLASSES_TABLE} (
    snapshot_id     BIGINT    NOT NULL REFERENCES {SNAPSHOTS_TABLE}(id) ON DELETE CASCADE,
    source_class_id BIGINT    NOT NULL,  -- the working row it was copied from
    examino         INTEGER   NOT NULL,
    course_code     TEXT      NOT NULL,
    section         TEXT      NOT NULL,
    name_suffix     TEXT,
    name_curriculum INTEGER,
    day             INTEGER,
    start_hour      INTEGER,
    duration        INTEGER,
    notes           TEXT,
    instructor_ids  INTEGER[] NOT NULL DEFAULT ARRAY[]::INTEGER[],  -- in position order
    room_codes      TEXT[]    NOT NULL DEFAULT ARRAY[]::TEXT[],
    PRIMARY KEY (snapshot_id, source_class_id)
);

-- Migrations. Here rather than in db.MIGRATIONS_SQL, which runs before this
-- schema: on a fresh database the table would not exist yet.
-- 2026-09-16: a class is a course code; the programme stamp is gone.
ALTER TABLE {CLASSES_TABLE} DROP COLUMN IF EXISTS curriculum;
-- 2026-09-23: which programme's title to print, for a code whose programmes
-- disagree (ΣΥΓ017). NULL is the newest.
ALTER TABLE {CLASSES_TABLE} ADD COLUMN IF NOT EXISTS name_curriculum INTEGER;
-- 2026-09-26: the change log learns the snapshot actions.
ALTER TABLE {CHANGES_TABLE} DROP CONSTRAINT IF EXISTS timetable_changes_action_check;
ALTER TABLE {CHANGES_TABLE} ADD CONSTRAINT timetable_changes_action_check
    CHECK (action IN {ACTIONS});
"""


# --------------------------------------------------------------------------
# Terms
# --------------------------------------------------------------------------

def stored_terms() -> list[tuple[int, str]]:
    """Every term, newest first (winter before spring within a year)."""
    engine = db.get_engine()
    if engine is None:
        return []
    with engine.connect() as conn:
        rows = conn.execute(
            text(
                f"SELECT year, period FROM {TERMS_TABLE} "
                "ORDER BY year DESC, (period = :winter) DESC"
            ),
            {"winter": WINTER},
        )
        return [(row[0], row[1]) for row in rows]


def term_state(year: int, period: str) -> dict | None:
    engine = db.get_engine()
    if engine is None:
        return None
    with engine.connect() as conn:
        row = (
            conn.execute(
                text(f"SELECT * FROM {TERMS_TABLE} WHERE year = :year AND period = :period"),
                {"year": year, "period": period},
            )
            .mappings()
            .first()
        )
    return dict(row) if row else None


def open_terms() -> list[tuple[int, str]]:
    return [term for term in stored_terms() if (term_state(*term) or {}).get("status") == OPEN]


def is_editable(year: int, period: str) -> bool:
    state = term_state(year, period)
    return bool(state and state["status"] == OPEN)


def open_term(
    year: int, period: str, baseline_year: int, baseline_period: str, opened_by: str
) -> str:
    """Create a term as a copy of another one — classes, links, staff activity.

    Refuses when the term exists or when another term is already open: one
    timetable is arranged at a time, and a second open one would leave "the
    open term" ambiguous everywhere.
    """
    engine = db.get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων."
    if period not in PERIODS:
        return f"Άγνωστη περίοδος: {period}"
    if term_state(year, period):
        return f"Το {term_label(year, period)} υπάρχει ήδη."
    if not term_state(baseline_year, baseline_period):
        return f"Δεν υπάρχει το {term_label(baseline_year, baseline_period)} για αντιγραφή."
    already = open_terms()
    if already:
        return f"Υπάρχει ήδη ανοιχτό το {term_label(*already[0])}· κλειδώστε το πρώτα."

    params = {
        "year": year,
        "period": period,
        "by": opened_by,
        "b_year": baseline_year,
        "b_period": baseline_period,
    }
    with engine.begin() as conn:
        conn.execute(
            text(
                f"INSERT INTO {TERMS_TABLE} "
                "(year, period, status, baseline_year, baseline_period, opened_by) "
                "VALUES (:year, :period, :status, :b_year, :b_period, :by)"
            ),
            {**params, "status": OPEN},
        )
        # Copy each class and remember old id -> new id for the link tables.
        old_ids = [
            row[0]
            for row in conn.execute(
                text(
                    f"SELECT id FROM {CLASSES_TABLE} "
                    "WHERE year = :b_year AND period = :b_period ORDER BY id"
                ),
                params,
            )
        ]
        mapping: dict[int, int] = {}
        for old_id in old_ids:
            mapping[old_id] = conn.execute(
                text(
                    f"INSERT INTO {CLASSES_TABLE} "
                    "(year, period, examino, course_code, section, "
                    " name_suffix, name_curriculum, day, start_hour, duration, notes, updated_by) "
                    "SELECT :year, :period, examino, course_code, section, "
                    "       name_suffix, name_curriculum, day, start_hour, duration, notes, :by "
                    f"FROM {CLASSES_TABLE} WHERE id = :old_id RETURNING id"
                ),
                {**params, "old_id": old_id},
            ).scalar_one()
        for old_id, new_id in mapping.items():
            conn.execute(
                text(
                    f"INSERT INTO {CLASS_INSTRUCTORS_TABLE} (class_id, staff_id, position) "
                    f"SELECT :new_id, staff_id, position FROM {CLASS_INSTRUCTORS_TABLE} "
                    "WHERE class_id = :old_id"
                ),
                {"new_id": new_id, "old_id": old_id},
            )
            conn.execute(
                text(
                    f"INSERT INTO {CLASS_ROOMS_TABLE} (class_id, room_code) "
                    f"SELECT :new_id, room_code FROM {CLASS_ROOMS_TABLE} "
                    "WHERE class_id = :old_id"
                ),
                {"new_id": new_id, "old_id": old_id},
            )
        # Staff activity starts as last time's; the coordinator adjusts it.
        conn.execute(
            text(
                f"INSERT INTO {STAFF_TERMS_TABLE} (staff_id, year, period, active, category) "
                f"SELECT staff_id, :year, :period, active, category FROM {STAFF_TERMS_TABLE} "
                "WHERE year = :b_year AND period = :b_period "
                "ON CONFLICT DO NOTHING"
            ),
            params,
        )
        _log(
            conn, year, period, None, OPENED,
            f"Αντιγραφή από {term_label(baseline_year, baseline_period)} ({len(mapping)} γραμμές)",
            opened_by,
        )
    return ""


def lock_term(year: int, period: str, locked_by: str) -> str:
    engine = db.get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων."
    state = term_state(year, period)
    if not state:
        return f"Δεν υπάρχει το {term_label(year, period)}."
    if state["status"] != OPEN:
        return f"Το {term_label(year, period)} είναι ήδη κλειδωμένο."
    with engine.begin() as conn:
        conn.execute(
            text(
                f"UPDATE {TERMS_TABLE} SET status = :status, locked_by = :by, locked_at = now() "
                "WHERE year = :year AND period = :period"
            ),
            {"status": LOCKED, "by": locked_by, "year": year, "period": period},
        )
        # Saved versions exist to arrive at this term; once it is locked they
        # can no longer be restored, so they go with it.
        dropped = conn.execute(
            text(f"DELETE FROM {SNAPSHOTS_TABLE} WHERE year = :year AND period = :period"),
            {"year": year, "period": period},
        ).rowcount
        detail = f"Διαγράφηκαν {dropped} εκδοχές" if dropped else None
        _log(conn, year, period, None, LOCKED_ACTION, detail, locked_by)
    return ""


def _log(conn, year, period, class_id, action, detail, author) -> None:
    conn.execute(
        text(
            f"INSERT INTO {CHANGES_TABLE} (year, period, class_id, action, detail, author) "
            "VALUES (:year, :period, :class_id, :action, :detail, :author)"
        ),
        {
            "year": year,
            "period": period,
            "class_id": class_id,
            "action": action,
            "detail": detail,
            "author": author,
        },
    )


def term_changes(year: int, period: str) -> pd.DataFrame:
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    with engine.connect() as conn:
        return pd.read_sql(
            text(
                f"SELECT created_at, author, action, class_id, detail FROM {CHANGES_TABLE} "
                "WHERE year = :year AND period = :period ORDER BY created_at DESC"
            ),
            conn,
            params={"year": year, "period": period},
        )


# --------------------------------------------------------------------------
# Reading a term
# --------------------------------------------------------------------------

# A course's name, by code: the programme the row names in ``name_curriculum``
# if it has the code, otherwise the newest that does. Codes carried over from
# 2018 kept their code, and only one of 84 changed its name (ΣΥΓ017) — a term
# that still teaches the 2018 course picks that title per row.
NEWEST_NAME_SQL = f"""
SELECT p.name FROM {PERIGRAMMATA_TABLE} p
WHERE p.locale = 'gr' AND p.code = c.course_code
ORDER BY (p.curriculum = c.name_curriculum) DESC NULLS LAST, p.curriculum DESC LIMIT 1
"""

_LOAD_TERM_SQL = f"""
SELECT c.id, c.examino, c.course_code, c.section, c.name_suffix, c.name_curriculum,
       c.day, c.start_hour, c.duration, c.notes, c.updated_by, c.updated_at,
       n.name AS course_name,
       COALESCE(i.names, '') AS instructors,
       COALESCE(i.ids, ARRAY[]::INTEGER[]) AS instructor_ids,
       COALESCE(i.conflict_ids, ARRAY[]::INTEGER[]) AS conflict_ids,
       COALESCE(r.codes, ARRAY[]::TEXT[]) AS room_codes,
       COALESCE(r.names, '') AS room
FROM {CLASSES_TABLE} c
LEFT JOIN LATERAL ({NEWEST_NAME_SQL}) n ON TRUE
LEFT JOIN LATERAL (
    SELECT string_agg(s.short_name, ', ' ORDER BY ci.position, s.short_name) AS names,
           array_agg(s.id ORDER BY ci.position, s.short_name) AS ids,
           array_agg(s.id ORDER BY ci.position, s.short_name)
               FILTER (WHERE NOT s.placeholder) AS conflict_ids
    FROM {CLASS_INSTRUCTORS_TABLE} ci JOIN {STAFF_TABLE} s ON s.id = ci.staff_id
    WHERE ci.class_id = c.id
) i ON TRUE
LEFT JOIN LATERAL (
    SELECT array_agg(cr.room_code ORDER BY cr.room_code) AS codes,
           string_agg(cr.room_code, ' & ' ORDER BY cr.room_code) AS names
    FROM {CLASS_ROOMS_TABLE} cr WHERE cr.class_id = c.id
) r ON TRUE
WHERE c.year = :year AND c.period = :period
ORDER BY c.examino, c.course_code, c.section, c.day, c.start_hour
"""


def load_term(year: int, period: str) -> pd.DataFrame:
    """Every class of a term, in the shape page 4 built from the workbook.

    The column names (``course_id``, ``class_name``, ``full_class_name``,
    ``semester``, ``instructors``, ``day``, ``start_time``, ``duration``,
    ``room``…) are the workbook's, so the Word export and the views written
    for the file work unchanged on the database.
    """
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    with engine.connect() as conn:
        frame = pd.read_sql(text(_LOAD_TERM_SQL), conn, params={"year": year, "period": period})
    return shape_term(frame, period)


# Nullable TEXT columns. A NULL comes back as NaN, and `NaN or ""` keeps the
# NaN because NaN is truthy — which is how «nan» ended up inside a text input
# on 2026-09-15. They are emptied once, here, so no caller has to remember.
OPTIONAL_TEXT_COLUMNS = ("name_suffix", "notes")


def shape_term(frame: pd.DataFrame, period: str) -> pd.DataFrame:
    """Add the derived columns page 4 expects. Pure — testable without a DB."""
    frame = frame.copy()
    for column in OPTIONAL_TEXT_COLUMNS:
        frame[column] = frame[column].where(frame[column].notna(), "")
    frame["course_id"] = frame["course_code"]
    frame["course_name"] = frame["course_name"].fillna("(άγνωστο μάθημα)")
    frame["display_name"] = [
        f"{name} ({suffix})" if suffix else name
        for name, suffix in zip(frame["course_name"], frame["name_suffix"])
    ]
    frame["class_name"] = frame["section"]
    frame["full_class_name"] = frame["display_name"] + " - " + frame["section"]
    frame["semester"] = frame["examino"]
    frame["teaching_period"] = period
    frame["placed"] = frame["day"].notna()
    frame["day_number"] = frame["day"]
    frame["day"] = [DAYS[int(d) - 1] if pd.notna(d) else None for d in frame["day_number"]]
    frame["start_time"] = [
        f"{int(h):02d}:00" if pd.notna(h) else None for h in frame["start_hour"]
    ]
    frame["end_hour"] = frame["start_hour"] + frame["duration"]
    frame["end_time"] = [
        f"{int(h):02d}:00" if pd.notna(h) else None for h in frame["end_hour"]
    ]
    return frame


# --------------------------------------------------------------------------
# Conflicts — pure, on the frame load_term returns
# --------------------------------------------------------------------------

ROOM_CONFLICT, STAFF_CONFLICT, SEMESTER_CONFLICT = "Αίθουσα", "Διδάσκων", "Εξάμηνο"


def _overlap(a, b) -> bool:
    return (
        a["day_number"] == b["day_number"]
        and a["start_hour"] < b["end_hour"]
        and b["start_hour"] < a["end_hour"]
    )


def streams(name_suffix) -> frozenset[str]:
    """The κατευθύνσεις a course belongs to, from its «ΔΥ, ΓΕ» marker.

    Each token is a stream letter (Δ Γ Σ Υ) plus Υ/Ε for compulsory/elective
    within it; only the stream letter matters for who attends. A course with
    no marker is common to everyone.
    """
    if not name_suffix or (isinstance(name_suffix, float) and pd.isna(name_suffix)):
        return frozenset()
    return frozenset(token.strip()[0] for token in str(name_suffix).split(",") if token.strip())


def _same_semester_allowed(a, b) -> bool:
    """When two classes of one εξάμηνο may overlap.

    Parallel lab groups of one course (Ε1 / Ε2) are meant to run together. And
    from the 7th εξάμηνο on, courses of *different* κατευθύνσεις are followed
    by different students: the 2025-26 workbook overlaps them on purpose
    (ΓΕΩ006 «ΓΥ» beside ΥΔΡ008 «ΥΥ, ΔΕ»), so two courses whose stream markers
    share no letter are not a conflict. A course without a marker is taken by
    everyone and conflicts with anything.
    """
    if (
        a["course_code"] == b["course_code"]
        and a["section"] != b["section"]
        and PARALLEL_LAB_RE.match(a["section"])
        and PARALLEL_LAB_RE.match(b["section"])
    ):
        return True
    streams_a, streams_b = streams(a["name_suffix"]), streams(b["name_suffix"])
    return bool(streams_a and streams_b and not (streams_a & streams_b))


def conflicts(frame: pd.DataFrame) -> pd.DataFrame:
    """Pairs of placed classes that cannot both hold, and why.

    Three rules, all on overlapping hours of the same day: a room used twice,
    an instructor in two places (placeholders such as «ΔΕΠ» excluded), and two
    classes of one εξάμηνο — except parallel lab groups of the same course and
    electives of disjoint κατευθύνσεις (see ``_same_semester_allowed``). Θ and
    Ε of the same course overlapping *is* a conflict: a student attends both.
    """
    placed = frame[frame["placed"]] if "placed" in frame else frame[frame["day"].notna()]
    rows = placed.to_dict("records")
    found = []
    for a, b in combinations(rows, 2):
        if not _overlap(a, b):
            continue
        reasons = []
        rooms = set(a["room_codes"]) & set(b["room_codes"])
        if rooms:
            reasons.append((ROOM_CONFLICT, ", ".join(sorted(rooms))))
        shared_ids = set(a["conflict_ids"]) & set(b["conflict_ids"])
        if shared_ids:
            names = dict(zip(a["instructor_ids"], str(a["instructors"]).split(", ")))
            reasons.append(
                (STAFF_CONFLICT, ", ".join(names.get(i, str(i)) for i in sorted(shared_ids)))
            )
        if a["examino"] == b["examino"] and not _same_semester_allowed(a, b):
            reasons.append((SEMESTER_CONFLICT, f"Εξάμηνο {a['examino']}"))
        for kind, what in reasons:
            found.append(
                {
                    "Είδος": kind,
                    "Τι": what,
                    "Ημέρα": a["day"],
                    "Ώρες": f"{a['start_time']}–{a['end_time']} / {b['start_time']}–{b['end_time']}",
                    "Α": f"{a['course_code']} {a['section']} (εξ. {a['examino']})",
                    "Β": f"{b['course_code']} {b['section']} (εξ. {b['examino']})",
                    "id_a": a["id"],
                    "id_b": b["id"],
                }
            )
    return pd.DataFrame(
        found, columns=["Είδος", "Τι", "Ημέρα", "Ώρες", "Α", "Β", "id_a", "id_b"]
    )


# --------------------------------------------------------------------------
# Classes
# --------------------------------------------------------------------------

def _validate_class(section: str, day, start_hour, duration) -> str:
    if not SECTION_RE.match(section or ""):
        return f"Μη έγκυρο τμήμα: {section!r} (Θ, Ε, Ε1…Ε9, Φ)."
    placed = [day, start_hour, duration]
    if any(v is None for v in placed) and not all(v is None for v in placed):
        return "Ημέρα, ώρα και διάρκεια ορίζονται μαζί ή καθόλου."
    if day is not None:
        if not 1 <= int(day) <= 5:
            return "Η ημέρα πρέπει να είναι Δευτέρα–Παρασκευή."
        if not FIRST_HOUR <= int(start_hour) < LAST_HOUR:
            return f"Η ώρα έναρξης πρέπει να είναι {FIRST_HOUR}:00–{LAST_HOUR - 1}:00."
        if not 1 <= int(duration) <= MAX_DURATION:
            return f"Η διάρκεια πρέπει να είναι 1–{MAX_DURATION} ώρες."
        if int(start_hour) + int(duration) > LAST_HOUR:
            return f"Το μάθημα πρέπει να τελειώνει έως τις {LAST_HOUR}:00."
    return ""


def add_class(
    year: int,
    period: str,
    *,
    examino: int,
    course_code: str,
    section: str,
    instructor_ids: list[int],
    room_codes: list[str],
    day: int | None,
    start_hour: int | None,
    duration: int | None,
    name_suffix: str | None = None,
    name_curriculum: int | None = None,
    notes: str | None = None,
    author: str,
) -> str:
    engine = db.get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων."
    if not is_editable(year, period):
        return f"Το {term_label(year, period)} δεν είναι ανοιχτό."
    error = _validate_class(section, day, start_hour, duration)
    if error:
        return error
    with engine.begin() as conn:
        class_id = conn.execute(
            text(
                f"INSERT INTO {CLASSES_TABLE} "
                "(year, period, examino, course_code, section, name_suffix, name_curriculum, "
                " day, start_hour, duration, notes, updated_by) "
                "VALUES (:year, :period, :examino, :course_code, :section, "
                "        :name_suffix, :name_curriculum, :day, :start_hour, :duration, :notes, :author) "
                "RETURNING id"
            ),
            {
                "year": year,
                "period": period,
                "examino": examino,
                "course_code": course_code.strip(),
                "section": section,
                "name_suffix": name_suffix or None,
                "name_curriculum": name_curriculum,
                "day": day,
                "start_hour": start_hour,
                "duration": duration,
                "notes": notes or None,
                "author": author,
            },
        ).scalar_one()
        _set_links(conn, class_id, instructor_ids, room_codes)
        _log(conn, year, period, class_id, ADD, _describe(course_code, section, day, start_hour, duration), author)
    return ""


def update_class(
    class_id: int,
    *,
    examino: int,
    section: str,
    instructor_ids: list[int],
    room_codes: list[str],
    day: int | None,
    start_hour: int | None,
    duration: int | None,
    name_suffix: str | None,
    notes: str | None,
    author: str,
    name_curriculum: int | None = None,
) -> str:
    engine = db.get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων."
    current = load_class(class_id)
    if current is None:
        return "Η γραμμή δεν υπάρχει πια."
    if not is_editable(current["year"], current["period"]):
        return "Το εξάμηνο δεν είναι ανοιχτό."
    error = _validate_class(section, day, start_hour, duration)
    if error:
        return error
    with engine.begin() as conn:
        conn.execute(
            text(
                f"UPDATE {CLASSES_TABLE} SET examino = :examino, section = :section, "
                "name_suffix = :name_suffix, name_curriculum = :name_curriculum, "
                "day = :day, start_hour = :start_hour, "
                "duration = :duration, notes = :notes, updated_by = :author, "
                "updated_at = now() WHERE id = :id"
            ),
            {
                "id": class_id,
                "examino": examino,
                "section": section,
                "name_suffix": name_suffix or None,
                "name_curriculum": name_curriculum,
                "day": day,
                "start_hour": start_hour,
                "duration": duration,
                "notes": notes or None,
                "author": author,
            },
        )
        conn.execute(
            text(f"DELETE FROM {CLASS_INSTRUCTORS_TABLE} WHERE class_id = :id"), {"id": class_id}
        )
        conn.execute(text(f"DELETE FROM {CLASS_ROOMS_TABLE} WHERE class_id = :id"), {"id": class_id})
        _set_links(conn, class_id, instructor_ids, room_codes)
        _log(
            conn, current["year"], current["period"], class_id, MODIFY,
            _describe(current["course_code"], section, day, start_hour, duration), author,
        )
    return ""


def delete_class(class_id: int, author: str) -> str:
    engine = db.get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων."
    current = load_class(class_id)
    if current is None:
        return "Η γραμμή δεν υπάρχει πια."
    if not is_editable(current["year"], current["period"]):
        return "Το εξάμηνο δεν είναι ανοιχτό."
    with engine.begin() as conn:
        conn.execute(text(f"DELETE FROM {CLASSES_TABLE} WHERE id = :id"), {"id": class_id})
        _log(
            conn, current["year"], current["period"], class_id, DELETE,
            _describe(current["course_code"], current["section"], current["day"],
                      current["start_hour"], current["duration"]),
            author,
        )
    return ""


def load_class(class_id: int) -> dict | None:
    engine = db.get_engine()
    if engine is None:
        return None
    with engine.connect() as conn:
        row = (
            conn.execute(text(f"SELECT * FROM {CLASSES_TABLE} WHERE id = :id"), {"id": class_id})
            .mappings()
            .first()
        )
    return dict(row) if row else None


def _set_links(conn, class_id: int, instructor_ids: list[int], room_codes: list[str]) -> None:
    if instructor_ids:
        conn.execute(
            text(
                f"INSERT INTO {CLASS_INSTRUCTORS_TABLE} (class_id, staff_id, position) "
                "VALUES (:class_id, :staff_id, :position)"
            ),
            [
                {"class_id": class_id, "staff_id": int(staff_id), "position": position}
                for position, staff_id in enumerate(dict.fromkeys(instructor_ids))
            ],
        )
    if room_codes:
        conn.execute(
            text(
                f"INSERT INTO {CLASS_ROOMS_TABLE} (class_id, room_code) "
                "VALUES (:class_id, :room_code)"
            ),
            [{"class_id": class_id, "room_code": code} for code in dict.fromkeys(room_codes)],
        )


def _describe(course_code, section, day, start_hour, duration) -> str:
    if day is None:
        return f"{course_code} {section}: χωρίς ώρα"
    return f"{course_code} {section}: {DAYS[int(day) - 1]} {int(start_hour):02d}:00 ({int(duration)}h)"


# --------------------------------------------------------------------------
# Εκδοχές — named snapshots of the open term
# --------------------------------------------------------------------------
#
# Save points, not parallel drafts. The term stays the one working timetable;
# a snapshot freezes it under a name so another arrangement can be tried, and
# the two compared or the first one restored. Rows remember the working id they
# came from (source_class_id), which is what lets a comparison say "moved"
# rather than "removed + added", and lets a restore update rows in place.

SNAPSHOT_FIELDS = (
    "examino", "course_code", "section", "name_suffix", "name_curriculum",
    "day", "start_hour", "duration", "notes",
)

# Every class of a term with its instructors and rooms as arrays — the shape a
# snapshot row stores. Instructors keep their position order; rooms are sorted,
# as load_term prints them.
_TERM_ROWS_SQL = f"""
SELECT c.id, {", ".join(f"c.{f}" for f in SNAPSHOT_FIELDS)},
       COALESCE((SELECT array_agg(ci.staff_id ORDER BY ci.position, ci.staff_id)
                 FROM {CLASS_INSTRUCTORS_TABLE} ci WHERE ci.class_id = c.id),
                ARRAY[]::INTEGER[]) AS instructor_ids,
       COALESCE((SELECT array_agg(cr.room_code ORDER BY cr.room_code)
                 FROM {CLASS_ROOMS_TABLE} cr WHERE cr.class_id = c.id),
                ARRAY[]::TEXT[]) AS room_codes
FROM {CLASSES_TABLE} c
WHERE c.year = :year AND c.period = :period
"""

# The same columns as _LOAD_TERM_SQL, so shape_term, conflicts, the calendar
# and the Word export work on a snapshot unchanged. ``id`` is the working row
# the snapshot row was copied from.
_LOAD_SNAPSHOT_SQL = f"""
SELECT c.source_class_id AS id, c.examino, c.course_code, c.section, c.name_suffix, c.name_curriculum,
       c.day, c.start_hour, c.duration, c.notes,
       NULL::TEXT AS updated_by, NULL::TIMESTAMPTZ AS updated_at,
       n.name AS course_name,
       COALESCE(i.names, '') AS instructors,
       c.instructor_ids,
       COALESCE(i.conflict_ids, ARRAY[]::INTEGER[]) AS conflict_ids,
       c.room_codes,
       array_to_string(c.room_codes, ' & ') AS room
FROM {SNAPSHOT_CLASSES_TABLE} c
LEFT JOIN LATERAL ({NEWEST_NAME_SQL}) n ON TRUE
LEFT JOIN LATERAL (
    SELECT string_agg(s.short_name, ', ' ORDER BY u.ord) AS names,
           array_agg(s.id ORDER BY u.ord) FILTER (WHERE NOT s.placeholder) AS conflict_ids
    FROM unnest(c.instructor_ids) WITH ORDINALITY AS u(staff_id, ord)
    JOIN {STAFF_TABLE} s ON s.id = u.staff_id
) i ON TRUE
WHERE c.snapshot_id = :snapshot_id
ORDER BY c.examino, c.course_code, c.section, c.day, c.start_hour
"""


def save_snapshot(year: int, period: str, name: str, note: str | None, author: str) -> str:
    """Freeze the open term under ``name``. Refuses a locked term or a used name."""
    engine = db.get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων."
    name = (name or "").strip()
    if not name:
        return "Δώστε όνομα στην εκδοχή (π.χ. «Επιλογή Α»)."
    if not is_editable(year, period):
        return f"Το {term_label(year, period)} δεν είναι ανοιχτό."
    params = {"year": year, "period": period, "name": name}
    with engine.begin() as conn:
        taken = conn.execute(
            text(
                f"SELECT 1 FROM {SNAPSHOTS_TABLE} "
                "WHERE year = :year AND period = :period AND name = :name"
            ),
            params,
        ).first()
        if taken:
            return f"Υπάρχει ήδη εκδοχή με το όνομα «{name}»."
        snapshot_id = conn.execute(
            text(
                f"INSERT INTO {SNAPSHOTS_TABLE} (year, period, name, note, created_by) "
                "VALUES (:year, :period, :name, :note, :author) RETURNING id"
            ),
            {**params, "note": (note or "").strip() or None, "author": author},
        ).scalar_one()
        columns = ", ".join(SNAPSHOT_FIELDS)
        count = conn.execute(
            text(
                f"INSERT INTO {SNAPSHOT_CLASSES_TABLE} "
                f"(snapshot_id, source_class_id, {columns}, instructor_ids, room_codes) "
                f"SELECT :snapshot_id, t.id, {columns}, t.instructor_ids, t.room_codes "
                f"FROM ({_TERM_ROWS_SQL}) t"
            ),
            {**params, "snapshot_id": snapshot_id},
        ).rowcount
        _log(conn, year, period, None, SNAPSHOT_ACTION, f"Αποθήκευση «{name}» ({count} γραμμές)", author)
    return ""


def list_snapshots(year: int, period: str) -> pd.DataFrame:
    """The term's snapshots, oldest first, with their row counts."""
    columns = ["id", "name", "note", "created_by", "created_at", "row_count"]
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame(columns=columns)
    with engine.connect() as conn:
        return pd.read_sql(
            text(
                f"SELECT s.id, s.name, s.note, s.created_by, s.created_at, "
                f"       count(c.source_class_id) AS row_count "
                f"FROM {SNAPSHOTS_TABLE} s "
                f"LEFT JOIN {SNAPSHOT_CLASSES_TABLE} c ON c.snapshot_id = s.id "
                "WHERE s.year = :year AND s.period = :period "
                "GROUP BY s.id ORDER BY s.created_at, s.id"
            ),
            conn,
            params={"year": year, "period": period},
        )


def snapshot_state(snapshot_id: int) -> dict | None:
    engine = db.get_engine()
    if engine is None:
        return None
    with engine.connect() as conn:
        row = (
            conn.execute(
                text(f"SELECT * FROM {SNAPSHOTS_TABLE} WHERE id = :id"), {"id": snapshot_id}
            )
            .mappings()
            .first()
        )
    return dict(row) if row else None


def load_snapshot(snapshot_id: int) -> pd.DataFrame:
    """A snapshot in exactly the shape ``load_term`` returns."""
    state = snapshot_state(snapshot_id)
    if state is None:
        return pd.DataFrame()
    with db.get_engine().connect() as conn:
        frame = pd.read_sql(text(_LOAD_SNAPSHOT_SQL), conn, params={"snapshot_id": snapshot_id})
    return shape_term(frame, state["period"])


def delete_snapshot(snapshot_id: int, author: str) -> str:
    engine = db.get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων."
    state = snapshot_state(snapshot_id)
    if state is None:
        return "Η εκδοχή δεν υπάρχει πια."
    with engine.begin() as conn:
        conn.execute(text(f"DELETE FROM {SNAPSHOTS_TABLE} WHERE id = :id"), {"id": snapshot_id})
        _log(
            conn, state["year"], state["period"], None, SNAPSHOT_ACTION,
            f"Διαγραφή «{state['name']}»", author,
        )
    return ""


def restore_snapshot(snapshot_id: int, author: str) -> str:
    """Make the working term what the snapshot holds, in one transaction.

    Rows are matched on the working id they were copied from: a row still in
    the term is updated in place (only if it differs), one deleted since is
    inserted again — with a new id, which every version of the term then
    adopts — and a row the snapshot does not have is deleted. So after a
    restore the version and the term compare as identical, and so do repeated
    restores.
    """
    engine = db.get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων."
    state = snapshot_state(snapshot_id)
    if state is None:
        return "Η εκδοχή δεν υπάρχει πια."
    year, period = state["year"], state["period"]
    if not is_editable(year, period):
        return f"Το {term_label(year, period)} δεν είναι ανοιχτό."

    columns = ", ".join(SNAPSHOT_FIELDS)
    updated = inserted = 0
    with engine.begin() as conn:
        saved = conn.execute(
            text(
                f"SELECT source_class_id, {columns}, instructor_ids, room_codes "
                f"FROM {SNAPSHOT_CLASSES_TABLE} WHERE snapshot_id = :id"
            ),
            {"id": snapshot_id},
        ).mappings().all()
        working = {
            row["id"]: row
            for row in conn.execute(
                text(_TERM_ROWS_SQL), {"year": year, "period": period}
            ).mappings()
        }
        keep = set()
        for row in saved:
            values = {field: row[field] for field in SNAPSHOT_FIELDS}
            current = working.get(row["source_class_id"])
            if current is None:
                class_id = conn.execute(
                    text(
                        f"INSERT INTO {CLASSES_TABLE} "
                        f"(year, period, {columns}, updated_by) VALUES (:year, :period, "
                        + ", ".join(f":{f}" for f in SNAPSHOT_FIELDS)
                        + ", :author) RETURNING id"
                    ),
                    {**values, "year": year, "period": period, "author": author},
                ).scalar_one()
                _set_links(conn, class_id, list(row["instructor_ids"]), list(row["room_codes"]))
                # The old id is gone for good (ids are never reused), so every
                # version of the term that held it now means the new row. Without
                # this, the version just restored would differ from the term it
                # produced, and each restore would re-insert the row again.
                conn.execute(
                    text(
                        f"UPDATE {SNAPSHOT_CLASSES_TABLE} SET source_class_id = :new "
                        "WHERE source_class_id = :old AND snapshot_id IN "
                        f"(SELECT id FROM {SNAPSHOTS_TABLE} WHERE year = :year AND period = :period)"
                    ),
                    {"new": class_id, "old": row["source_class_id"], "year": year, "period": period},
                )
                inserted += 1
                continue
            keep.add(current["id"])
            if all(current[f] == values[f] for f in SNAPSHOT_FIELDS) and list(
                current["instructor_ids"]
            ) == list(row["instructor_ids"]) and list(current["room_codes"]) == list(row["room_codes"]):
                continue
            conn.execute(
                text(
                    f"UPDATE {CLASSES_TABLE} SET "
                    + ", ".join(f"{f} = :{f}" for f in SNAPSHOT_FIELDS)
                    + ", updated_by = :author, updated_at = now() WHERE id = :id"
                ),
                {**values, "id": current["id"], "author": author},
            )
            conn.execute(
                text(f"DELETE FROM {CLASS_INSTRUCTORS_TABLE} WHERE class_id = :id"), {"id": current["id"]}
            )
            conn.execute(
                text(f"DELETE FROM {CLASS_ROOMS_TABLE} WHERE class_id = :id"), {"id": current["id"]}
            )
            _set_links(conn, current["id"], list(row["instructor_ids"]), list(row["room_codes"]))
            updated += 1
        removed = [class_id for class_id in working if class_id not in keep]
        if removed:
            conn.execute(
                text(f"DELETE FROM {CLASSES_TABLE} WHERE id = ANY(:ids)"), {"ids": removed}
            )
        _log(
            conn, year, period, None, RESTORED_ACTION,
            f"Επαναφορά «{state['name']}»: {updated} άλλαξαν, {inserted} ξαναπροστέθηκαν, "
            f"{len(removed)} διαγράφηκαν",
            author,
        )
    return ""


def summarize(frame: pd.DataFrame) -> dict:
    """Rows, unplaced rows and conflicts of a term or a snapshot — what ranks
    one arrangement against another at a glance. Pure."""
    if frame.empty:
        return {"Γραμμές": 0, "Χωρίς ώρα": 0, "Συγκρούσεις": 0}
    return {
        "Γραμμές": len(frame),
        "Χωρίς ώρα": int((~frame["placed"].astype(bool)).sum()),
        "Συγκρούσεις": len(conflicts(frame)),
    }


def _when(row) -> str:
    if not row["placed"]:
        return "χωρίς ώρα"
    return f"{row['day']} {row['start_time']}–{row['end_time']}"


# What a comparison reports on, in this order: label -> how to print it.
COMPARED_ASPECTS = (
    ("Ώρα", _when),
    ("Αίθουσα", lambda row: row["room"] or "—"),
    ("Διδάσκοντες", lambda row: row["instructors"] or "—"),
    ("Εξάμηνο", lambda row: str(int(row["examino"]))),
    ("Τμήμα", lambda row: row["section"]),
    ("Τίτλος", lambda row: row["course_name"]),  # differs only through name_curriculum
    ("Ένδειξη", lambda row: row["name_suffix"] or "—"),
    ("Παρατηρήσεις", lambda row: row["notes"] or "—"),
)
ADDED, REMOVED, CHANGED = "Προστέθηκε", "Αφαιρέθηκε", "Άλλαξε"


def compare_frames(before: pd.DataFrame, after: pd.DataFrame) -> pd.DataFrame:
    """What differs between two load_term-shaped frames, matched on ``id``.

    One row per class that was added, removed or changed; for a change, only
    the aspects that moved, «πριν» and «μετά». Unchanged classes are left out.
    Pure, so testable without a database.
    """
    columns = ["Μεταβολή", "Εξάμηνο", "Μάθημα", "Τι", "Πριν", "Μετά"]
    old = {int(row["id"]): row for _, row in before.iterrows()}
    new = {int(row["id"]): row for _, row in after.iterrows()}

    def label(row) -> str:
        return f"{row['course_code']} {row['section']} · {row['course_name']}"

    def full(row) -> str:
        return " · ".join(
            show(row) for name, show in COMPARED_ASPECTS if name in ("Ώρα", "Αίθουσα", "Διδάσκοντες")
        )

    found = []
    for class_id in old.keys() | new.keys():
        a, b = old.get(class_id), new.get(class_id)
        if a is None:
            found.append((ADDED, b["examino"], label(b), "", "", full(b), b))
        elif b is None:
            found.append((REMOVED, a["examino"], label(a), "", full(a), "", a))
        else:
            moved = [(name, show(a), show(b)) for name, show in COMPARED_ASPECTS if show(a) != show(b)]
            if moved:
                found.append(
                    (
                        CHANGED, b["examino"], label(b),
                        ", ".join(m[0] for m in moved),
                        " · ".join(m[1] for m in moved),
                        " · ".join(m[2] for m in moved),
                        b,
                    )
                )
    found.sort(key=lambda f: (int(f[1]), f[6]["course_code"], f[6]["section"], f[0]))
    return pd.DataFrame([(k, int(e), m, t, p, n) for k, e, m, t, p, n, _ in found], columns=columns)


# --------------------------------------------------------------------------
# Candidate courses for the arranging tab
# --------------------------------------------------------------------------

def build_catalogue(programmes: pd.DataFrame) -> pd.DataFrame:
    """One row per course code, out of every Greek περίγραμμα of every programme.

    ``programmes`` has ``curriculum, course_code, course_name, examino``.
    ``course_name`` is the newest programme's (the rule ``NEWEST_NAME_SQL``
    applies to a term); ``examina`` maps each programme that has the code to
    its εξάμηνο there — ``{2018: 3, 2025: 4}`` for ΔΟΜ007 — and leaves out a
    programme where the εξάμηνο is blank. ``names`` maps every programme that
    has the code to its title there; see ``title_options``. Pure, so testable
    without a DB.
    """
    columns = ["course_code", "course_name", "examina", "names"]
    if programmes.empty:
        return pd.DataFrame(columns=columns)
    ordered = programmes.sort_values(["course_code", "curriculum"], ascending=[True, False])
    rows = []
    for code, group in ordered.groupby("course_code", sort=True):
        examina = {
            int(curriculum): int(examino)
            for curriculum, examino in zip(group["curriculum"], group["examino"])
            if pd.notna(examino)
        }
        names = {int(c): name for c, name in zip(group["curriculum"], group["course_name"])}
        rows.append(
            {"course_code": code, "course_name": group["course_name"].iloc[0], "examina": examina, "names": names}
        )
    return pd.DataFrame(rows, columns=columns)


def courses_for_semester(
    catalogue: pd.DataFrame, semester: int, all_semesters: bool = False
) -> pd.DataFrame:
    """What the arranging tab offers for one εξάμηνο of a term.

    By default, every code **either** programme places in that εξάμηνο: while
    the two run side by side, a course that moved (ΔΟΜ007, 3rd in 2018 and 4th
    in 2025) could be taught in either, and guessing which is exactly what the
    timetable stopped doing. ``all_semesters`` offers every code, from both
    periods, for a course taught outside its usual εξάμηνο.
    """
    if all_semesters or catalogue.empty:
        return catalogue
    return catalogue[[semester in examina.values() for examina in catalogue["examina"]]]


def title_options(names: dict[int, str]) -> dict[int | None, str]:
    """The titles a row of this code may print, keyed by ``name_curriculum``.

    One entry keyed ``None`` (the newest programme's) when every programme
    agrees; otherwise one per programme, newest first — ΣΥΓ017 offers both its
    2025 and its 2018 title.
    """
    if len(set(names.values())) <= 1:
        return {None: next(iter(names.values()))} if names else {}
    newest = max(names)
    return {
        (None if curriculum == newest else curriculum): names[curriculum]
        for curriculum in sorted(names, reverse=True)
    }


def programme_semesters(catalogue: pd.DataFrame, period: str) -> set[int]:
    """The εξάμηνα of ``period`` that any programme has courses in."""
    return {
        examino
        for examina in catalogue["examina"]
        for examino in examina.values()
        if period_for(examino) == period
    }


def examina_label(examina: dict[int, int]) -> str:
    """«εξ. 3 στο 2018 · εξ. 4 στο 2025», or «εξ. 1» when every programme agrees."""
    if not examina:
        return "χωρίς εξάμηνο"
    if len(set(examina.values())) == 1:
        return f"εξ. {next(iter(examina.values()))}"
    return " · ".join(f"εξ. {examino} στο {curriculum}" for curriculum, examino in sorted(examina.items()))


def course_catalogue(year: int, period: str) -> pd.DataFrame:
    """``build_catalogue`` from the database, with ``in_term`` for the term's codes."""
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame(columns=["course_code", "course_name", "examina", "names", "in_term"])
    with engine.connect() as conn:
        programmes = pd.read_sql(
            text(
                "SELECT curriculum, code AS course_code, name AS course_name, examino "
                f"FROM {PERIGRAMMATA_TABLE} WHERE locale = 'gr'"
            ),
            conn,
        )
        present = {
            row[0]
            for row in conn.execute(
                text(
                    f"SELECT DISTINCT course_code FROM {CLASSES_TABLE} "
                    "WHERE year = :year AND period = :period"
                ),
                {"year": year, "period": period},
            )
        }
    catalogue = build_catalogue(programmes)
    catalogue["in_term"] = catalogue["course_code"].isin(present)
    return catalogue


# --------------------------------------------------------------------------
# Staff and rooms
# --------------------------------------------------------------------------

def load_staff() -> pd.DataFrame:
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    with engine.connect() as conn:
        return pd.read_sql(
            text(f"SELECT * FROM {STAFF_TABLE} ORDER BY placeholder, category, last_name, first_name"),
            conn,
        )


def staff_for_term(year: int, period: str) -> pd.DataFrame:
    """Every staff row plus that term's ``active`` flag (None when unknown)."""
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    with engine.connect() as conn:
        return pd.read_sql(
            text(
                f"SELECT s.*, t.active, COALESCE(t.category, s.category) AS term_category "
                f"FROM {STAFF_TABLE} s "
                f"LEFT JOIN {STAFF_TERMS_TABLE} t "
                "  ON t.staff_id = s.id AND t.year = :year AND t.period = :period "
                "ORDER BY s.placeholder, s.category, s.last_name, s.first_name"
            ),
            conn,
            params={"year": year, "period": period},
        )


def set_staff_active(staff_id: int, year: int, period: str, active: bool) -> None:
    engine = db.get_engine()
    if engine is None:
        return
    with engine.begin() as conn:
        conn.execute(
            text(
                f"INSERT INTO {STAFF_TERMS_TABLE} (staff_id, year, period, active) "
                "VALUES (:staff_id, :year, :period, :active) "
                "ON CONFLICT (staff_id, year, period) DO UPDATE SET active = EXCLUDED.active"
            ),
            {"staff_id": int(staff_id), "year": year, "period": period, "active": bool(active)},
        )


STAFF_FIELDS = ("short_name", "last_name", "first_name", "category", "rank",
                "subject", "email", "website_url", "placeholder", "notes")


def save_staff(row: dict, author: str, staff_id: int | None = None) -> str:
    """Insert (``staff_id`` None) or update one person."""
    engine = db.get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων."
    if not (row.get("short_name") or "").strip() or not (row.get("last_name") or "").strip():
        return "Χρειάζονται σύντομο όνομα και επώνυμο."
    if row.get("category") not in CATEGORIES:
        return f"Άγνωστη κατηγορία: {row.get('category')}"
    values = {name: (row.get(name) or None) for name in STAFF_FIELDS}
    values["placeholder"] = bool(row.get("placeholder"))
    values["short_name"] = values["short_name"].strip()
    values["author"] = author
    assignments = ", ".join(f"{name} = :{name}" for name in STAFF_FIELDS)
    with engine.begin() as conn:
        if staff_id is None:
            conn.execute(
                text(
                    f"INSERT INTO {STAFF_TABLE} ({', '.join(STAFF_FIELDS)}, updated_by) "
                    f"VALUES ({', '.join(':' + n for n in STAFF_FIELDS)}, :author)"
                ),
                values,
            )
        else:
            conn.execute(
                text(
                    f"UPDATE {STAFF_TABLE} SET {assignments}, updated_by = :author, "
                    "updated_at = now() WHERE id = :id"
                ),
                {**values, "id": int(staff_id)},
            )
    return ""


def load_rooms(active_only: bool = False) -> pd.DataFrame:
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    where = "WHERE active" if active_only else ""
    with engine.connect() as conn:
        return pd.read_sql(text(f"SELECT * FROM {ROOMS_TABLE} {where} ORDER BY kind, code"), conn)


ROOM_FIELDS = ("code", "name", "kind", "capacity", "active", "notes")


def save_room(row: dict) -> str:
    engine = db.get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων."
    code = (row.get("code") or "").strip()
    if not code or not (row.get("name") or "").strip():
        return "Χρειάζονται κωδικός και όνομα."
    if row.get("kind") not in ROOM_KINDS:
        return f"Άγνωστο είδος: {row.get('kind')}"
    capacity = row.get("capacity")
    values = {
        "code": code,
        "name": row["name"].strip(),
        "kind": row["kind"],
        "capacity": None if capacity in (None, "") or pd.isna(capacity) else int(capacity),
        "active": bool(row.get("active", True)),
        "notes": row.get("notes") or None,
    }
    with engine.begin() as conn:
        conn.execute(
            text(
                f"INSERT INTO {ROOMS_TABLE} (code, name, kind, capacity, active, notes) "
                "VALUES (:code, :name, :kind, :capacity, :active, :notes) "
                "ON CONFLICT (code) DO UPDATE SET name = EXCLUDED.name, kind = EXCLUDED.kind, "
                "capacity = EXCLUDED.capacity, active = EXCLUDED.active, notes = EXCLUDED.notes"
            ),
            values,
        )
    return ""
