"""Load the pre-app timetable data into Postgres, once.

Three sources, all under ``files/timetables/``:

- ``staff.csv`` — the people, as listed on the department site on 2026-09-15
  (faculty_members, lab_staff, tech_staff, temp_staff) plus the two former
  faculty members who still teach and the «ΔΕΠ» placeholder for courses taught
  by all faculty. Loaded only while the table is empty: after that the page
  owns the list.
- ``rooms.csv`` — the rooms, one code each. The workbook wrote rooms as free
  text («Αίθ. Τεχν. Σχεδίου 1», «Εργαστήριο Τεχνικού Σχεδίου Ι» — the same
  room); ``ROOM_ALIASES`` maps every spelling the workbook used to its code
  and anything unknown *raises* rather than importing a room nobody named.
- ``<year>-<year+1>.xlsm`` — the timetable as page 4 reads it. Each workbook
  becomes two locked terms. Instructors are matched to ``short_name``; the
  bare «Σαφούρη» on ΔΟΜ007 is Γεωργία (the drawing lab; Χριστίνα is ΕΤΕΠ on
  the physics lab), and «ΔΕΠ» is the placeholder.

This is the only import: from 2026-27 on a term is opened as a copy in the
app. ``2026-2027.xlsm`` is a byte-identical copy of the 2025-26 workbook and
is deliberately **not** listed in ``WORKBOOKS`` — loading it would seed a
2026-27 that is really 2025-26 under another name.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from sqlalchemy import Engine, text

from perigrammata_db import COURSES_TABLE as PERIGRAMMATA_TABLE
from timetable_db import (
    CLASS_INSTRUCTORS_TABLE,
    CLASS_ROOMS_TABLE,
    CLASSES_TABLE,
    DAY_NUMBER,
    LOCKED,
    NEW_CURRICULUM,
    OLD_CURRICULUM,
    PERIODS,
    ROOMS_TABLE,
    STAFF_TABLE,
    STAFF_TERMS_TABLE,
    TERMS_TABLE,
    curriculum_for,
    period_for,
)

ROOT = Path(__file__).resolve().parents[1]
TIMETABLES_DIR = ROOT / "files" / "timetables"
STAFF_CSV = TIMETABLES_DIR / "staff.csv"
ROOMS_CSV = TIMETABLES_DIR / "rooms.csv"
SHEET_NAME = "timetable"

# year -> workbook. Explicit, not a glob: see the module docstring.
WORKBOOKS = {2025: TIMETABLES_DIR / "2025-2026.xlsm"}

# Every room string the workbooks used, after whitespace is collapsed and
# «A & B» split on the ampersand. Room numbers and ΗΥ1/ΗΥ2 are already codes.
ROOM_ALIASES = {
    "Αίθ. Αρχιτ.": "ΑΡΧ",
    "Αίθ. Αρχιτεκτονικής": "ΑΡΧ",
    "Αίθ. Εργαστήριου Φυσικής Ι": "ΦΥΣ",
    "Αίθ. Τεχν. Σχεδίου 1": "ΤΣ1",
    "Αίθ. Τεχν. Σχεδίου Ι": "ΤΣ1",
    "Εργαστήριο Τεχνικού Σχεδίου Ι": "ΤΣ1",
    "Αίθ. Τεχν. Σχεδίου ΙΙ": "ΤΣ2",
    "Αιθ. Εργ. Μεταφορών και Συγκοινωνιακής Υποδομής": "ΜΕΤΑΦ",
    "Εργαστήριο Γεωδαισίας": "ΓΕΩΔ",
    "Εργαστήριο Ποιοτικού Ελέγχου": "ΠΟΙΟΤ",
}

INSTRUCTOR_ALIASES = {"Σαφούρη": "Σαφούρη Γ."}
SECTION_ALIASES = {"Φροντιστηριακό": "Φ"}

# «Οικοδομική ΙΙ (ΔΥ, ΣΕ, ΥΕ)» -> suffix «ΔΥ, ΣΕ, ΥΕ». The markers are pairs of
# capitals; the separator is a comma or, on one row, a full stop.
SUFFIX_RE = re.compile(r"\s*\(([Α-ΩA-Z]{2}(?:\s*[,.]\s*[Α-ΩA-Z]{2})*)\)\s*$")


def _clean(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)) or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return re.sub(r"\s+", " ", str(value)).strip()


def split_rooms(value) -> list[str]:
    """Room codes for one workbook cell; raises on a spelling nobody mapped."""
    codes = []
    for part in _clean(value).split("&"):
        part = part.strip()
        if not part:
            continue
        if part.isdigit() or part in ("ΗΥ1", "ΗΥ2"):
            codes.append(part)
        elif part in ROOM_ALIASES:
            codes.append(ROOM_ALIASES[part])
        else:
            raise ValueError(f"Άγνωστη αίθουσα στο βιβλίο εργασίας: {part!r}")
    return codes


def split_instructors(value) -> list[str]:
    names = []
    for part in re.split(r"[,;]", _clean(value)):
        part = part.strip()
        if part:
            names.append(INSTRUCTOR_ALIASES.get(part, part))
    return names


def split_name(course_name) -> tuple[str, str | None]:
    """The course name without the elective-group marker, and the marker."""
    name = _clean(course_name)
    match = SUFFIX_RE.search(name)
    if not match:
        return name, None
    suffix = re.sub(r"\s*[,.]\s*", ", ", match.group(1))
    return name[: match.start()].strip(), suffix


def read_workbook(path: Path, year: int) -> list[dict]:
    """One dict per workbook row, with codes resolved and nothing joined yet."""
    frame = pd.read_excel(path, sheet_name=SHEET_NAME)
    rows = []
    for _, row in frame.iterrows():
        code = _clean(row["course_id"])
        if not code:
            continue
        examino = int(row["semester"])
        period = _clean(row.get("teaching_period"))
        day = _clean(row.get("day"))
        _, suffix = split_name(row.get("course_name"))
        record = {
            "examino": examino,
            "course_code": code,
            "curriculum": curriculum_for(year, examino),
            "section": SECTION_ALIASES.get(_clean(row.get("class_name")), _clean(row.get("class_name")) or "Θ"),
            "name_suffix": suffix,
            "period": period or None,
            "day": DAY_NUMBER[day] if day else None,
            "start_hour": None,
            "duration": None,
            "notes": _clean(row.get("notes")) or None,
            "instructors": split_instructors(row.get("instructors")),
            "rooms": split_rooms(row.get("room")),
        }
        if day:
            start = row["start_time"]
            record["start_hour"] = int(start.hour if hasattr(start, "hour") else str(start).split(":")[0])
            record["duration"] = int(row["duration"])
        rows.append(record)
    return rows


def read_staff(path: Path = STAFF_CSV) -> list[dict]:
    frame = pd.read_csv(path, dtype=str).fillna("")
    rows = []
    for _, row in frame.iterrows():
        record = {column: (row[column].strip() or None) for column in frame.columns}
        record["placeholder"] = bool(row["placeholder"].strip())
        rows.append(record)
    return rows


def read_rooms(path: Path = ROOMS_CSV) -> list[dict]:
    frame = pd.read_csv(path, dtype=str).fillna("")
    rows = []
    for _, row in frame.iterrows():
        rows.append(
            {
                "code": row["code"].strip(),
                "name": row["name"].strip(),
                "kind": row["kind"].strip(),
                "capacity": int(row["capacity"]) if row["capacity"].strip() else None,
                "notes": row["notes"].strip() or None,
            }
        )
    return rows


def seed_timetable(engine: Engine) -> str:
    parts = []
    with engine.begin() as conn:
        if not conn.execute(text(f"SELECT 1 FROM {STAFF_TABLE} LIMIT 1")).first():
            staff = read_staff()
            conn.execute(
                text(
                    f"INSERT INTO {STAFF_TABLE} (short_name, last_name, first_name, category, "
                    "rank, subject, website_url, placeholder, notes, updated_by) "
                    "VALUES (:short_name, :last_name, :first_name, :category, :rank, :subject, "
                    ":website_url, :placeholder, :notes, 'seed')"
                ),
                staff,
            )
            parts.append(f"προσωπικό: {len(staff)}")
        if not conn.execute(text(f"SELECT 1 FROM {ROOMS_TABLE} LIMIT 1")).first():
            rooms = read_rooms()
            conn.execute(
                text(
                    f"INSERT INTO {ROOMS_TABLE} (code, name, kind, capacity, notes) "
                    "VALUES (:code, :name, :kind, :capacity, :notes)"
                ),
                rooms,
            )
            parts.append(f"αίθουσες: {len(rooms)}")

    loaded, skipped = [], []
    for year, path in sorted(WORKBOOKS.items()):
        if not path.exists():
            continue
        with engine.connect() as conn:
            already = conn.execute(
                text(f"SELECT 1 FROM {TERMS_TABLE} WHERE year = :year LIMIT 1"), {"year": year}
            ).first()
        if already:
            skipped.append(str(year))
            continue
        rows = read_workbook(path, year)
        counts = _insert_year(engine, year, rows)
        loaded.append(f"{year}-{str(year + 1)[-2:]} ({counts})")

    if loaded:
        parts.append("φορτώθηκαν: " + ", ".join(loaded))
    if skipped:
        parts.append("υπήρχαν ήδη: " + ", ".join(skipped))
    return "Πρόγραμμα — " + ("· ".join(parts) if parts else "τίποτα νέο")


def _insert_year(engine: Engine, year: int, rows: list[dict]) -> str:
    with engine.begin() as conn:
        staff_ids = {
            row[0]: row[1]
            for row in conn.execute(text(f"SELECT short_name, id FROM {STAFF_TABLE}"))
        }
        room_codes = {row[0] for row in conn.execute(text(f"SELECT code FROM {ROOMS_TABLE}"))}
        known = {
            (row[0], row[1])
            for row in conn.execute(
                text(f"SELECT curriculum, code FROM {PERIGRAMMATA_TABLE} WHERE locale = 'gr'")
            )
        }
        for period in PERIODS:
            conn.execute(
                text(
                    f"INSERT INTO {TERMS_TABLE} (year, period, status, opened_by, locked_by, locked_at) "
                    "VALUES (:year, :period, :status, 'seed', 'seed', now())"
                ),
                {"year": year, "period": period, "status": LOCKED},
            )
        fallbacks, placed = 0, 0
        for row in rows:
            missing = [name for name in row["instructors"] if name not in staff_ids]
            if missing:
                raise ValueError(f"Άγνωστος διδάσκων στο βιβλίο εργασίας: {missing}")
            unknown_rooms = [code for code in row["rooms"] if code not in room_codes]
            if unknown_rooms:
                raise ValueError(f"Άγνωστη αίθουσα: {unknown_rooms}")
            # The rule says which programme an εξάμηνο follows; a code the
            # programme does not carry (ΔΟΜ004 in 2025) is taken from the other.
            if known and (row["curriculum"], row["course_code"]) not in known:
                other = OLD_CURRICULUM if row["curriculum"] == NEW_CURRICULUM else NEW_CURRICULUM
                if (other, row["course_code"]) in known:
                    row["curriculum"] = other
                    fallbacks += 1
            # An unplaced row has no period in the workbook; it belongs to the
            # period its εξάμηνο runs in.
            period = row["period"] or period_for(row["examino"])
            class_id = conn.execute(
                text(
                    f"INSERT INTO {CLASSES_TABLE} (year, period, examino, course_code, curriculum, "
                    "section, name_suffix, day, start_hour, duration, notes, updated_by) "
                    "VALUES (:year, :period, :examino, :course_code, :curriculum, :section, "
                    ":name_suffix, :day, :start_hour, :duration, :notes, 'seed') RETURNING id"
                ),
                {**{k: row[k] for k in ("examino", "course_code", "curriculum", "section",
                                        "name_suffix", "day", "start_hour", "duration", "notes")},
                 "year": year, "period": period},
            ).scalar_one()
            if row["day"]:
                placed += 1
            if row["instructors"]:
                conn.execute(
                    text(
                        f"INSERT INTO {CLASS_INSTRUCTORS_TABLE} (class_id, staff_id, position) "
                        "VALUES (:class_id, :staff_id, :position)"
                    ),
                    [
                        {"class_id": class_id, "staff_id": staff_ids[name], "position": i}
                        for i, name in enumerate(dict.fromkeys(row["instructors"]))
                    ],
                )
            if row["rooms"]:
                conn.execute(
                    text(
                        f"INSERT INTO {CLASS_ROOMS_TABLE} (class_id, room_code) "
                        "VALUES (:class_id, :room_code)"
                    ),
                    [{"class_id": class_id, "room_code": c} for c in dict.fromkeys(row["rooms"])],
                )
        # Whoever taught in a term was active in it; the rest stay unknown
        # until the coordinator says.
        conn.execute(
            text(
                f"INSERT INTO {STAFF_TERMS_TABLE} (staff_id, year, period, active, category) "
                f"SELECT DISTINCT ci.staff_id, c.year, c.period, TRUE, s.category "
                f"FROM {CLASS_INSTRUCTORS_TABLE} ci "
                f"JOIN {CLASSES_TABLE} c ON c.id = ci.class_id "
                f"JOIN {STAFF_TABLE} s ON s.id = ci.staff_id "
                "WHERE c.year = :year ON CONFLICT DO NOTHING"
            ),
            {"year": year},
        )
    note = f"{len(rows)} γραμμές, {placed} με ώρα"
    if fallbacks:
        note += f", {fallbacks} από το άλλο πρόγραμμα"
    return note
