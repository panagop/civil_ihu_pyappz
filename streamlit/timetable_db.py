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

Course names are **not** stored: ``curriculum`` + ``course_code`` resolve them
from ``perigrammata_courses``. The curriculum is stored on the row rather than
looked up ("try 2025, then 2018") so a name resolves deterministically. The
only name-ish thing kept here is ``name_suffix``, the «ΔΥ, ΥΕ» elective-group
marker the printed timetable carries after the course name.
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
ACTIONS = (ADD, MODIFY, DELETE, OPENED, LOCKED_ACTION)

# Which εξάμηνα follow the 2025 programme in a given academic year; the rest
# follow 2018. The department is mid-transition, one year of study per year:
# 2025-26 ran the new programme in the first year only, 2026-27 runs it in the
# first two (εξάμηνα 1–4; corrected 2026-09-15). Extend the tuple when the
# next year moves over. Years not listed are entirely 2018.
NEW_CURRICULUM_EXAMINA: dict[int, tuple[int, ...]] = {
    2025: (1, 2),
    2026: (1, 2, 3, 4),
}
OLD_CURRICULUM, NEW_CURRICULUM = 2018, 2025


def curriculum_for(year: int, examino: int) -> int:
    return NEW_CURRICULUM if examino in NEW_CURRICULUM_EXAMINA.get(year, ()) else OLD_CURRICULUM


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
    curriculum  INTEGER NOT NULL,
    section     TEXT    NOT NULL,
    name_suffix TEXT,
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
        # A copied row follows the programme its εξάμηνο runs on in the *new*
        # year, where that programme has the code *in the same εξάμηνο*: the
        # transition moves one year of study forward each year, so last
        # winter's 3rd-εξάμηνο rows (2018) become 2025 rows in 2026-27. The
        # εξάμηνο must match because 84 codes are shared between programmes
        # with different content and sometimes a different εξάμηνο (ΔΟΜ007 is
        # 3rd in 2018, 4th in 2025). Anything else keeps its old curriculum and
        # is listed by the page as belonging to the old programme.
        known = {
            (row[0], row[1]): row[2]
            for row in conn.execute(
                text(
                    f"SELECT curriculum, code, examino FROM {PERIGRAMMATA_TABLE} "
                    "WHERE locale = 'gr'"
                )
            )
        }
        mapping: dict[int, int] = {}
        moved = 0
        for old_id in old_ids:
            new_id = conn.execute(
                text(
                    f"INSERT INTO {CLASSES_TABLE} "
                    "(year, period, examino, course_code, curriculum, section, "
                    " name_suffix, day, start_hour, duration, notes, updated_by) "
                    "SELECT :year, :period, examino, course_code, curriculum, section, "
                    "       name_suffix, day, start_hour, duration, notes, :by "
                    f"FROM {CLASSES_TABLE} WHERE id = :old_id RETURNING id"
                ),
                {**params, "old_id": old_id},
            ).scalar_one()
            mapping[old_id] = new_id
            examino, code, curriculum = conn.execute(
                text(f"SELECT examino, course_code, curriculum FROM {CLASSES_TABLE} WHERE id = :id"),
                {"id": new_id},
            ).one()
            wanted = curriculum_for(year, examino)
            if wanted != curriculum and known.get((wanted, code)) == examino:
                conn.execute(
                    text(f"UPDATE {CLASSES_TABLE} SET curriculum = :c WHERE id = :id"),
                    {"c": wanted, "id": new_id},
                )
                moved += 1
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
            f"Αντιγραφή από {term_label(baseline_year, baseline_period)} ({len(mapping)} γραμμές"
            + (f", {moved} στο πρόγραμμα 2025)" if moved else ")"),
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
        _log(conn, year, period, None, LOCKED_ACTION, None, locked_by)
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

_LOAD_TERM_SQL = f"""
SELECT c.id, c.examino, c.course_code, c.curriculum, c.section, c.name_suffix,
       c.day, c.start_hour, c.duration, c.notes, c.updated_by, c.updated_at,
       p.name AS course_name,
       COALESCE(i.names, '') AS instructors,
       COALESCE(i.ids, ARRAY[]::INTEGER[]) AS instructor_ids,
       COALESCE(i.conflict_ids, ARRAY[]::INTEGER[]) AS conflict_ids,
       COALESCE(r.codes, ARRAY[]::TEXT[]) AS room_codes,
       COALESCE(r.names, '') AS room
FROM {CLASSES_TABLE} c
LEFT JOIN {PERIGRAMMATA_TABLE} p
       ON p.curriculum = c.curriculum AND p.locale = 'gr' AND p.code = c.course_code
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
    curriculum: int,
    section: str,
    instructor_ids: list[int],
    room_codes: list[str],
    day: int | None,
    start_hour: int | None,
    duration: int | None,
    name_suffix: str | None = None,
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
                "(year, period, examino, course_code, curriculum, section, name_suffix, "
                " day, start_hour, duration, notes, updated_by) "
                "VALUES (:year, :period, :examino, :course_code, :curriculum, :section, "
                "        :name_suffix, :day, :start_hour, :duration, :notes, :author) "
                "RETURNING id"
            ),
            {
                "year": year,
                "period": period,
                "examino": examino,
                "course_code": course_code.strip(),
                "curriculum": curriculum,
                "section": section,
                "name_suffix": name_suffix or None,
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
                "name_suffix = :name_suffix, day = :day, start_hour = :start_hour, "
                "duration = :duration, notes = :notes, updated_by = :author, "
                "updated_at = now() WHERE id = :id"
            ),
            {
                "id": class_id,
                "examino": examino,
                "section": section,
                "name_suffix": name_suffix or None,
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
# Candidate courses for the arranging tab
# --------------------------------------------------------------------------

def off_programme(frame: pd.DataFrame, year: int) -> pd.DataFrame:
    """Rows whose curriculum is not the one their εξάμηνο runs on in ``year``.

    After a copy across the transition these are last year's courses of an
    εξάμηνο that has since moved to the 2025 programme — to be replaced by the
    coordinator, not silently.
    """
    if frame.empty:
        return frame
    mask = [
        int(curriculum) != curriculum_for(year, int(examino))
        for curriculum, examino in zip(frame["curriculum"], frame["examino"])
    ]
    return frame[mask]


def candidate_courses(year: int, period: str) -> pd.DataFrame:
    """Courses of the period's εξάμηνα, from the curriculum each one follows.

    Read from the περιγράμματα so the list of what *could* be timetabled is
    the programme itself, with a flag for the ones the term already has.
    """
    engine = db.get_engine()
    if engine is None:
        return pd.DataFrame()
    with engine.connect() as conn:
        courses = pd.read_sql(
            text(
                f"SELECT curriculum, code AS course_code, name AS course_name, examino "
                f"FROM {PERIGRAMMATA_TABLE} WHERE locale = 'gr' AND examino IS NOT NULL"
            ),
            conn,
        )
        present = pd.read_sql(
            text(
                f"SELECT DISTINCT course_code, curriculum FROM {CLASSES_TABLE} "
                "WHERE year = :year AND period = :period"
            ),
            conn,
            params={"year": year, "period": period},
        )
    courses["examino"] = courses["examino"].astype(int)
    courses = courses[courses["examino"].map(period_for) == period]
    courses = courses[
        [
            curriculum_for(year, int(examino)) == int(curriculum)
            for curriculum, examino in zip(courses["curriculum"], courses["examino"])
        ]
    ]
    present_keys = set(zip(present["course_code"], present["curriculum"]))
    courses["in_term"] = [
        (code, cur) in present_keys for code, cur in zip(courses["course_code"], courses["curriculum"])
    ]
    return courses.sort_values(["examino", "course_code"]).reset_index(drop=True)


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
