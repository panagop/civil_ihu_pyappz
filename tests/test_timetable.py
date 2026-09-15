"""The timetable pipeline: workbook parsing without a database, and the schema,
seed, copy-open and conflict rules against a throwaway Postgres (``pgserver``,
skipped where it is not installed — see ``test_db_rename.py``)."""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "streamlit"))

import db  # noqa: E402
import seed_timetable as seed  # noqa: E402
import timetable_db as tdb  # noqa: E402
from settings import get_secret  # noqa: E402

WORKBOOK = seed.WORKBOOKS[2025]


# --------------------------------------------------------------------------
# No database
# --------------------------------------------------------------------------

@pytest.fixture(scope="module")
def rows():
    return seed.read_workbook(WORKBOOK, 2025)


def test_workbook_counts(rows):
    assert len(rows) == 140
    assert sum(r["day"] is not None for r in rows) == 124
    assert {r["course_code"] for r in rows} == set(pd.read_excel(WORKBOOK, sheet_name="timetable")["course_id"].dropna())


def test_every_instructor_and_room_is_known(rows):
    staff = {r["short_name"] for r in seed.read_staff()}
    rooms = {r["code"] for r in seed.read_rooms()}
    assert {n for r in rows for n in r["instructors"]} <= staff
    assert {c for r in rows for c in r["rooms"]} <= rooms


def test_room_aliases_collapse_to_one_drawing_lab():
    assert seed.split_rooms("Αίθ. Τεχν. Σχεδίου 1") == ["ΤΣ1"]
    assert seed.split_rooms("Εργαστήριο Τεχνικού Σχεδίου Ι ") == ["ΤΣ1"]
    assert seed.split_rooms("Αίθ. Αρχιτεκτονικής & Αίθ. Τεχν.\nΣχεδίου Ι") == ["ΑΡΧ", "ΤΣ1"]
    assert seed.split_rooms(301.0) == ["301"]
    with pytest.raises(ValueError):
        seed.split_rooms("Αίθουσα που δεν υπάρχει")


def test_name_suffix():
    assert seed.split_name("Οικοδομική ΙΙ (ΔΥ, ΣΕ, ΥΕ)") == ("Οικοδομική ΙΙ", "ΔΥ, ΣΕ, ΥΕ")
    assert seed.split_name("Εγγειοβελτιωτικά Έργα – Αρδεύσεις (ΥΥ. ΓΕ)")[1] == "ΥΥ, ΓΕ"
    assert seed.split_name("Φυσική για Μηχανικούς") == ("Φυσική για Μηχανικούς", None)


def test_curriculum_rule():
    assert tdb.curriculum_for(2026, 1) == 2025
    assert tdb.curriculum_for(2026, 4) == 2025
    assert tdb.curriculum_for(2026, 5) == 2018
    assert tdb.curriculum_for(2025, 3) == 2018
    assert tdb.curriculum_for(2024, 1) == 2018
    assert tdb.period_for(1) == tdb.WINTER and tdb.period_for(8) == tdb.SPRING


def _frame(specs):
    """A minimal frame in the shape load_term returns, from
    (id, examino, code, section, day, start, duration, rooms, staff_ids)."""
    records = []
    for i, examino, code, section, day, start, duration, rooms, staff, *suffix in specs:
        records.append(
            {
                "id": i, "examino": examino, "course_code": code, "curriculum": 2018,
                "section": section, "name_suffix": suffix[0] if suffix else None, "day": day, "start_hour": start,
                "duration": duration, "notes": None, "updated_by": None, "updated_at": None,
                "course_name": code, "instructors": ", ".join(f"S{s}" for s in staff),
                "instructor_ids": list(staff), "conflict_ids": [s for s in staff if s != 99],
                "room_codes": list(rooms), "room": " & ".join(rooms),
            }
        )
    return tdb.shape_term(pd.DataFrame(records), tdb.WINTER)


def test_conflicts():
    frame = _frame([
        (1, 1, "A", "Θ", 1, 9, 2, ["301"], [1]),
        (2, 1, "B", "Θ", 1, 10, 2, ["202"], [2]),      # same εξάμηνο, overlaps 1
        (3, 3, "C", "Θ", 1, 9, 2, ["301"], [3]),       # same room as 1
        (4, 5, "D", "Θ", 1, 10, 1, ["204"], [1]),      # same instructor as 1
        (5, 2, "E", "Ε1", 2, 9, 2, ["ΤΣ1"], [4]),
        (6, 2, "E", "Ε2", 2, 9, 2, ["ΤΣ2"], [5]),      # parallel lab group: allowed
        (7, 2, "E", "Θ", 2, 10, 1, ["202"], [4]),      # Θ over the lab: conflict (twice: εξάμηνο + staff)
        (8, 7, "F", "Θ", 3, 9, 4, ["205"], [99]),
        (9, 9, "G", "Θ", 3, 9, 4, ["206"], [99]),      # placeholder «ΔΕΠ» shared: no conflict
        (10, 9, "H", "Θ", None, None, None, [], [1]),  # unplaced: ignored
        (11, 7, "I", "Θ", 4, 9, 4, ["301"], [6], "ΓΥ"),
        (12, 7, "J", "Θ", 4, 9, 4, ["202"], [7], "ΥΥ, ΔΕ"),   # other streams: allowed
        (13, 7, "K", "Θ", 4, 9, 4, ["204"], [8], "ΓΕ"),       # shares Γ with 11: conflict
        (14, 7, "L", "Θ", 4, 9, 4, ["205"], [9]),             # common course: conflicts with all
    ])
    found = tdb.conflicts(frame)
    pairs = {(row["id_a"], row["id_b"], row["Είδος"]) for _, row in found.iterrows()}
    assert (1, 2, tdb.SEMESTER_CONFLICT) in pairs
    assert (1, 3, tdb.ROOM_CONFLICT) in pairs
    assert (1, 4, tdb.STAFF_CONFLICT) in pairs
    assert not any(a == 5 and b == 6 for a, b, _ in pairs)
    assert (5, 7, tdb.SEMESTER_CONFLICT) in pairs
    assert (5, 7, tdb.STAFF_CONFLICT) in pairs
    assert (6, 7, tdb.SEMESTER_CONFLICT) in pairs
    assert not any({a, b} == {8, 9} for a, b, _ in pairs)
    assert not any(10 in (a, b) for a, b, _ in pairs)
    assert (11, 12, tdb.SEMESTER_CONFLICT) not in pairs
    assert (11, 13, tdb.SEMESTER_CONFLICT) in pairs
    assert (12, 13, tdb.SEMESTER_CONFLICT) not in pairs
    assert {(11, 14), (12, 14), (13, 14)} <= {(a, b) for a, b, k in pairs if k == tdb.SEMESTER_CONFLICT}
    assert len(found) == 10


def test_streams():
    assert tdb.streams("ΥΥ, ΔΕ") == {"Υ", "Δ"}
    assert tdb.streams(None) == frozenset()
    assert tdb.streams(float("nan")) == frozenset()


# --------------------------------------------------------------------------
# With a database
# --------------------------------------------------------------------------

pgserver = pytest.importorskip("pgserver")


@pytest.fixture(scope="module")
def server(tmp_path_factory):
    instance = pgserver.get_server(tmp_path_factory.mktemp("pgdata"))
    yield instance
    instance.cleanup()


@pytest.fixture(scope="module")
def database(server, tmp_path_factory):
    """One seeded database for the module: the seed takes a while."""
    name = "timetable_" + re.sub(r"\W", "_", tmp_path_factory.mktemp("x").name).lower()[:40]
    server.psql(f"CREATE DATABASE {name}")
    url = server.get_uri(database=name)
    previous = os.environ.get("DATABASE_URL")
    os.environ["DATABASE_URL"] = url
    if get_secret("DATABASE_URL") != url:
        pytest.skip("DATABASE_URL is set in secrets.toml and shadows the test one")
    db.get_engine.clear()
    db.bootstrap.clear()
    status = db.bootstrap()
    assert "Σφάλμα" not in status, status
    yield url
    engine = db.get_engine()
    if engine is not None:
        engine.dispose()
    db.get_engine.clear()
    db.bootstrap.clear()
    if previous is None:
        del os.environ["DATABASE_URL"]
    else:
        os.environ["DATABASE_URL"] = previous


def test_seed_makes_two_locked_terms(database):
    assert tdb.stored_terms() == [(2025, tdb.WINTER), (2025, tdb.SPRING)]
    assert not tdb.open_terms()
    winter = tdb.load_term(2025, tdb.WINTER)
    spring = tdb.load_term(2025, tdb.SPRING)
    assert len(winter) + len(spring) == 140
    assert int(winter["placed"].sum()) + int(spring["placed"].sum()) == 124
    # Names resolve from the περιγράμματα for every row
    assert not (winter["course_name"] == "(άγνωστο μάθημα)").any()
    assert not (spring["course_name"] == "(άγνωστο μάθημα)").any()
    # ΔΟΜ004 is not in the 2025 programme and fell back to 2018
    dom004 = spring[spring["course_code"] == "ΔΟΜ004"].iloc[0]
    assert dom004["curriculum"] == 2018
    # ΔΟΜ011 kept both rooms, ΓΕΝ002 both instructors in workbook order
    dom011 = winter[winter["course_code"] == "ΔΟΜ011"].iloc[0]
    assert dom011["room_codes"] == ["207", "ΤΣ1"]
    gen002 = winter[winter["course_code"] == "ΓΕΝ002"].iloc[0]
    assert gen002["instructors"] == "Βοζίκης, Παπαϊωάννου"
    # Suffix printed after the name
    dom022 = winter[winter["course_code"] == "ΔΟΜ022"].iloc[0]
    assert dom022["display_name"] == "Οικοδομική ΙΙ (ΔΥ, ΣΕ, ΥΕ)"
    # Seeding is idempotent
    assert "υπήρχαν ήδη" in seed.seed_timetable(db.get_engine())


def test_seed_staff_and_rooms(database):
    staff = tdb.load_staff()
    assert len(staff) == 34
    assert staff[staff["placeholder"]]["short_name"].tolist() == ["ΔΕΠ"]
    active = tdb.staff_for_term(2025, tdb.WINTER)
    assert active[active["short_name"] == "Παπαϊωάννου"]["active"].iloc[0] is True or bool(
        active[active["short_name"] == "Παπαϊωάννου"]["active"].iloc[0]
    )
    assert pd.isna(active[active["short_name"] == "Δημητρακάκης"]["active"].iloc[0])
    assert len(tdb.load_rooms()) == 15


def test_locked_term_refuses_edits(database):
    row = tdb.load_term(2025, tdb.WINTER).iloc[0]
    assert "ανοιχτό" in tdb.delete_class(int(row["id"]), "x@ihu.gr")


def test_open_copy_edit_and_lock(database):
    assert tdb.open_term(2026, tdb.WINTER, 2025, tdb.WINTER, "c@ihu.gr") == ""
    assert tdb.open_terms() == [(2026, tdb.WINTER)]
    assert "ήδη ανοιχτό" in tdb.open_term(2026, tdb.SPRING, 2025, tdb.SPRING, "c@ihu.gr")
    copied = tdb.load_term(2026, tdb.WINTER)
    original = tdb.load_term(2025, tdb.WINTER)
    assert len(copied) == len(original)
    assert copied["instructors"].tolist() == original["instructors"].tolist()
    assert copied["room_codes"].tolist() == original["room_codes"].tolist()
    # 2026-27 runs the 2025 programme in εξάμηνα 1–4: the copied 3rd-εξάμηνο
    # rows move to it where the code exists there, and their names resolve.
    third = copied[copied["examino"] == 3]
    assert (third["curriculum"] == 2025).any()
    assert not (third["course_name"] == "(άγνωστο μάθημα)").any()
    assert (copied[copied["examino"] >= 5]["curriculum"] == 2018).all()
    # ΔΟΜ007 is 3rd in 2018 but 4th in 2025: it must not be re-resolved, and
    # it is reported as off-programme together with the codes 2025 lacks.
    dom007 = third[third["course_code"] == "ΔΟΜ007"]
    assert (dom007["curriculum"] == 2018).all()
    stale = tdb.off_programme(copied, 2026)
    assert set(stale["course_code"]) >= {"ΔΟΜ007", "ΔΟΜ006", "ΔΟΜ008", "ΥΔΡ001"}
    assert (stale["examino"] == 3).all()
    assert tdb.off_programme(tdb.load_term(2025, tdb.WINTER), 2025).empty
    # The copy carries last winter's activity
    active = tdb.staff_for_term(2026, tdb.WINTER)
    assert bool(active[active["short_name"] == "Κίρτας"]["active"].iloc[0])

    staff = tdb.load_staff()
    kirtas = int(staff[staff["short_name"] == "Κίρτας"]["id"].iloc[0])
    assert tdb.add_class(
        2026, tdb.WINTER, examino=1, course_code="ΓΕΝ001", curriculum=2025, section="Φ",
        instructor_ids=[kirtas], room_codes=["301"], day=1, start_hour=9, duration=1,
        author="c@ihu.gr",
    ) == ""
    assert "τελειώνει" in tdb.add_class(
        2026, tdb.WINTER, examino=1, course_code="ΓΕΝ001", curriculum=2025, section="Θ",
        instructor_ids=[], room_codes=[], day=1, start_hour=20, duration=3, author="c@ihu.gr",
    )
    assert "τμήμα" in tdb.add_class(
        2026, tdb.WINTER, examino=1, course_code="ΓΕΝ001", curriculum=2025, section="Χ",
        instructor_ids=[], room_codes=[], day=None, start_hour=None, duration=None, author="c@ihu.gr",
    )
    term = tdb.load_term(2026, tdb.WINTER)
    added = term[(term["course_code"] == "ΓΕΝ001") & (term["section"] == "Φ")].iloc[0]
    assert added["instructors"] == "Κίρτας" and added["room_codes"] == ["301"]
    # Κίρτας is on ΓΕΩ003 Δευτέρα 11:00; move the new row onto it -> staff conflict
    assert tdb.update_class(
        int(added["id"]), examino=1, section="Φ", instructor_ids=[kirtas], room_codes=["301"],
        day=1, start_hour=11, duration=1, name_suffix=None, notes=None, author="c@ihu.gr",
    ) == ""
    problems = tdb.conflicts(tdb.load_term(2026, tdb.WINTER))
    assert (problems["Είδος"] == tdb.STAFF_CONFLICT).any()
    assert tdb.delete_class(int(added["id"]), "c@ihu.gr") == ""

    candidates = tdb.candidate_courses(2026, tdb.WINTER)
    assert set(candidates["examino"].unique()) <= {1, 3, 5, 7, 9}
    assert (candidates[candidates["examino"] <= 4]["curriculum"] == 2025).all()
    assert (candidates[candidates["examino"] >= 5]["curriculum"] == 2018).all()

    assert tdb.lock_term(2026, tdb.WINTER, "c@ihu.gr") == ""
    assert not tdb.open_terms()
    changes = tdb.term_changes(2026, tdb.WINTER)
    assert changes["action"].tolist()[0] == tdb.LOCKED_ACTION
    assert set(changes["action"]) >= {tdb.OPENED, tdb.ADD, tdb.MODIFY, tdb.DELETE, tdb.LOCKED_ACTION}


def test_page_renders(database):
    """The page script runs end to end against the seeded database.

    ``st.user`` is logged out under AppTest, so this exercises the public
    views and the read-only side of Προετοιμασία; the writes are covered above.
    """
    from streamlit.testing.v1 import AppTest

    page = ROOT / "streamlit" / "app_pages" / "8_🗓_timetable_v2.py"
    app = AppTest.from_file(str(page), default_timeout=300)
    app.run()
    assert not app.exception, [str(e.value) for e in app.exception]
    assert not app.error, [e.value for e in app.error]
    assert len(app.tabs) == 8
    assert app.selectbox[0].label == "Εξάμηνο:"
