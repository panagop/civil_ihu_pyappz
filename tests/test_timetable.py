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
from sqlalchemy import text as sa_text

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
    return seed.read_workbook(WORKBOOK)


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


def test_period_for():
    assert tdb.period_for(1) == tdb.WINTER and tdb.period_for(8) == tdb.SPRING


def _programmes():
    """(curriculum, code, name, εξάμηνο) rows shaped like perigrammata_courses."""
    return pd.DataFrame(
        [
            (2018, "ΔΟΜ007", "Οικοδομική Ι", 3),
            (2025, "ΔΟΜ007", "Οικοδομική Ι", 4),          # moved to spring
            (2018, "ΣΥΓ017", "Οργάνωση Εργοταξίου", 9),
            (2025, "ΣΥΓ017", "Προγραμματισμός Έργων", 9),  # renamed
            (2018, "ΔΟΜ004", "Τεχνική Μηχανική Ι", 2),     # 2018 only
            (2025, "ΓΕΝ099", "Νέο μάθημα", 3),             # 2025 only
            (2018, "ΧΧΧ001", "Χωρίς εξάμηνο", None),
        ],
        columns=["curriculum", "course_code", "course_name", "examino"],
    )


def test_catalogue_is_one_row_per_code_with_the_newest_name():
    catalogue = tdb.build_catalogue(_programmes()).set_index("course_code")
    assert catalogue.index.is_unique and len(catalogue) == 5
    assert catalogue.loc["ΣΥΓ017", "course_name"] == "Προγραμματισμός Έργων"
    assert catalogue.loc["ΔΟΜ004", "course_name"] == "Τεχνική Μηχανική Ι"
    assert catalogue.loc["ΔΟΜ007", "examina"] == {2018: 3, 2025: 4}
    assert catalogue.loc["ΧΧΧ001", "examina"] == {}
    assert catalogue.loc["ΣΥΓ017", "names"] == {2018: "Οργάνωση Εργοταξίου", 2025: "Προγραμματισμός Έργων"}


def test_title_options():
    # Programmes agree: one title, the default.
    assert tdb.title_options({2018: "Οικοδομική Ι", 2025: "Οικοδομική Ι"}) == {None: "Οικοδομική Ι"}
    # They disagree (ΣΥΓ017): the newest is the default, the older one is picked by year.
    assert tdb.title_options({2018: "Οργάνωση Εργοταξίου", 2025: "Προγραμματισμός Έργων"}) == {
        None: "Προγραμματισμός Έργων",
        2018: "Οργάνωση Εργοταξίου",
    }
    assert tdb.title_options({}) == {}


def test_courses_for_semester():
    catalogue = tdb.build_catalogue(_programmes())

    def codes(semester, all_semesters=False):
        return set(tdb.courses_for_semester(catalogue, semester, all_semesters)["course_code"])

    # Either programme's εξάμηνο counts: ΔΟΜ007 is offered in the 3rd and the 4th.
    assert codes(3) == {"ΔΟΜ007", "ΓΕΝ099"}
    assert codes(4) == {"ΔΟΜ007"}
    assert codes(9) == {"ΣΥΓ017"}
    # The toggle offers everything, both periods, even a course with no εξάμηνο.
    assert codes(3, all_semesters=True) == set(catalogue["course_code"])
    assert tdb.programme_semesters(catalogue, tdb.WINTER) == {3, 9}
    assert tdb.programme_semesters(catalogue, tdb.SPRING) == {2, 4}


def test_examina_label():
    assert tdb.examina_label({2018: 3, 2025: 4}) == "εξ. 3 στο 2018 · εξ. 4 στο 2025"
    assert tdb.examina_label({2018: 9, 2025: 9}) == "εξ. 9"
    assert tdb.examina_label({}) == "χωρίς εξάμηνο"


def _frame(specs):
    """A minimal frame in the shape load_term returns, from
    (id, examino, code, section, day, start, duration, rooms, staff_ids)."""
    records = []
    for i, examino, code, section, day, start, duration, rooms, staff, *suffix in specs:
        records.append(
            {
                "id": i, "examino": examino, "course_code": code,
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


def test_optional_text_columns_are_never_nan():
    """A NULL must reach the edit form as "", not as NaN.

    `NaN or ""` keeps the NaN — NaN is truthy — so an unfilled «Ένδειξη μετά
    τον τίτλο» rendered the literal «nan» in the text input.
    """
    frame = _frame([(1, 1, "A", "Θ", 1, 9, 2, ["301"], [1])])
    for column in tdb.OPTIONAL_TEXT_COLUMNS:
        assert frame[column].iloc[0] == ""
        assert not pd.isna(frame[column].iloc[0])
    assert frame["display_name"].iloc[0] == "A"

    with_suffix = _frame([(1, 1, "A", "Θ", 1, 9, 2, ["301"], [1], "ΔΥ, ΥΕ")])
    assert with_suffix["display_name"].iloc[0] == "A (ΔΥ, ΥΕ)"


def test_compare_frames():
    before = _frame([
        (1, 1, "A", "Θ", 1, 9, 2, ["301"], [1]),
        (2, 1, "B", "Θ", 2, 9, 2, ["301"], [2]),
        (3, 3, "C", "Θ", 3, 9, 2, ["202"], [3]),
        (4, 3, "D", "Θ", None, None, None, [], []),
    ])
    after = _frame([
        (1, 1, "A", "Θ", 1, 9, 2, ["301"], [1]),         # unchanged
        (2, 1, "B", "Θ", 4, 11, 2, ["301"], [2]),        # moved
        (3, 3, "C", "Θ", 3, 9, 2, ["204"], [3, 4]),      # other room, extra instructor
        (5, 5, "E", "Ε1", 1, 9, 2, ["ΤΣ1"], [5]),        # new
    ])                                                   # 4 removed
    diff = tdb.compare_frames(before, after)
    assert len(diff) == 4
    by_code = {row["Μάθημα"].split()[0]: row for _, row in diff.iterrows()}
    assert set(by_code) == {"B", "C", "D", "E"}
    assert by_code["B"]["Μεταβολή"] == tdb.CHANGED and by_code["B"]["Τι"] == "Ώρα"
    assert by_code["B"]["Πριν"] == "Τρίτη 09:00–11:00" and by_code["B"]["Μετά"] == "Πέμπτη 11:00–13:00"
    assert by_code["C"]["Τι"] == "Αίθουσα, Διδάσκοντες"
    assert by_code["D"]["Μεταβολή"] == tdb.REMOVED and by_code["D"]["Πριν"].startswith("χωρίς ώρα")
    assert by_code["E"]["Μεταβολή"] == tdb.ADDED and by_code["E"]["Εξάμηνο"] == 5
    assert diff["Εξάμηνο"].tolist() == sorted(diff["Εξάμηνο"])
    assert tdb.compare_frames(before, before).empty

    summary = tdb.summarize(before)
    assert summary == {"Γραμμές": 4, "Χωρίς ώρα": 1, "Συγκρούσεις": 0}


def test_streams():
    assert tdb.streams("ΥΥ, ΔΕ") == {"Υ", "Δ"}
    assert tdb.streams(None) == frozenset()
    assert tdb.streams(float("nan")) == frozenset()


# --------------------------------------------------------------------------
# With a database
# --------------------------------------------------------------------------

pgserver = pytest.importorskip("pixeltable_pgserver")


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
    # Names by code, newest programme first: ΔΟΜ004 exists only in 2018, and
    # ΣΥΓ017 prints its 2025 name even in a 2025-26 term.
    assert "curriculum" not in winter.columns
    dom004 = spring[spring["course_code"] == "ΔΟΜ004"].iloc[0]
    assert dom004["course_name"] == "Τεχνική Μηχανική Ι"
    syg017 = winter[winter["course_code"] == "ΣΥΓ017"].iloc[0]
    assert syg017["course_name"] == "Προγραμματισμός και Διαχείριση Τεχνικών Έργων"
    # ΔΟΜ011 kept both rooms, ΓΕΝ002 both instructors in workbook order
    dom011 = winter[winter["course_code"] == "ΔΟΜ011"].iloc[0]
    assert dom011["room_codes"] == ["207", "ΤΣ1"]
    gen002 = winter[winter["course_code"] == "ΓΕΝ002"].iloc[0]
    assert gen002["instructors"] == "Βοζίκης, Παπαϊωάννου"
    # Suffix printed after the name
    dom022 = winter[winter["course_code"] == "ΔΟΜ022"].iloc[0]
    assert dom022["display_name"] == "Οικοδομική ΙΙ (ΔΥ, ΣΕ, ΥΕ)"
    # A course with no marker and no note reads back as empty strings, not NaN
    gen001 = winter[winter["course_code"] == "ΓΕΝ001"].iloc[0]
    assert gen001["name_suffix"] == "" and gen001["notes"] == ""
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
    # A copy is the same courses, same εξάμηνα, same names.
    assert copied["course_code"].tolist() == original["course_code"].tolist()
    assert copied["examino"].tolist() == original["examino"].tolist()
    assert copied["course_name"].tolist() == original["course_name"].tolist()
    # The copy carries last winter's activity
    active = tdb.staff_for_term(2026, tdb.WINTER)
    assert bool(active[active["short_name"] == "Κίρτας"]["active"].iloc[0])

    staff = tdb.load_staff()
    kirtas = int(staff[staff["short_name"] == "Κίρτας"]["id"].iloc[0])
    assert tdb.add_class(
        2026, tdb.WINTER, examino=1, course_code="ΓΕΝ001", section="Φ",
        instructor_ids=[kirtas], room_codes=["301"], day=1, start_hour=9, duration=1,
        author="c@ihu.gr",
    ) == ""
    assert "τελειώνει" in tdb.add_class(
        2026, tdb.WINTER, examino=1, course_code="ΓΕΝ001", section="Θ",
        instructor_ids=[], room_codes=[], day=1, start_hour=20, duration=3, author="c@ihu.gr",
    )
    assert "τμήμα" in tdb.add_class(
        2026, tdb.WINTER, examino=1, course_code="ΓΕΝ001", section="Χ",
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

    # What the arranging tab offers: by code, from both programmes. ΔΟΜ007
    # (3rd in 2018, 4th in 2025) is in the term already; ΥΔΡ002 (4th in 2018,
    # 3rd in 2025) is a spring course last year and is offered for the 3rd now.
    catalogue = tdb.course_catalogue(2026, tdb.WINTER)
    assert catalogue["course_code"].is_unique
    third = tdb.courses_for_semester(catalogue, 3).set_index("course_code")
    assert bool(third.loc["ΔΟΜ007", "in_term"])
    assert not bool(third.loc["ΥΔΡ002", "in_term"])
    assert tdb.programme_semesters(catalogue, tdb.WINTER) <= {1, 3, 5, 7, 9}
    # Any code can go into any εξάμηνο of the term, and its name resolves.
    assert tdb.add_class(
        2026, tdb.WINTER, examino=3, course_code="ΥΔΡ002", section="Θ",
        instructor_ids=[], room_codes=[], day=None, start_hour=None, duration=None, author="c@ihu.gr",
    ) == ""
    term = tdb.load_term(2026, tdb.WINTER)
    moved = term[term["course_code"] == "ΥΔΡ002"].iloc[0]
    assert moved["examino"] == 3 and moved["course_name"] == "Μηχανική των Ρευστών"

    # ΣΥΓ017 taught as the 2018 course: the row prints the 2018 title, and
    # editing it back to the default restores the 2025 one.
    assert tdb.add_class(
        2026, tdb.WINTER, examino=9, course_code="ΣΥΓ017", section="Ε", name_curriculum=2018,
        instructor_ids=[], room_codes=[], day=None, start_hour=None, duration=None, author="c@ihu.gr",
    ) == ""
    term = tdb.load_term(2026, tdb.WINTER)
    old_title = term[(term["course_code"] == "ΣΥΓ017") & (term["section"] == "Ε")].iloc[0]
    assert old_title["course_name"] == "Οργάνωση Εργοταξίου και Δομικές Μηχανές"
    assert term[(term["course_code"] == "ΣΥΓ017") & (term["section"] == "Θ")].iloc[0][
        "course_name"
    ] == "Προγραμματισμός και Διαχείριση Τεχνικών Έργων"
    assert tdb.update_class(
        int(old_title["id"]), examino=9, section="Ε", instructor_ids=[], room_codes=[],
        day=None, start_hour=None, duration=None, name_suffix=None, notes=None, author="c@ihu.gr",
        name_curriculum=None,
    ) == ""
    term = tdb.load_term(2026, tdb.WINTER)
    assert term[term["id"] == old_title["id"]].iloc[0]["course_name"] == "Προγραμματισμός και Διαχείριση Τεχνικών Έργων"

    assert tdb.lock_term(2026, tdb.WINTER, "c@ihu.gr") == ""
    assert not tdb.open_terms()
    changes = tdb.term_changes(2026, tdb.WINTER)
    assert changes["action"].tolist()[0] == tdb.LOCKED_ACTION
    assert set(changes["action"]) >= {tdb.OPENED, tdb.ADD, tdb.MODIFY, tdb.DELETE, tdb.LOCKED_ACTION}


COMPARED = ["id", "examino", "course_code", "section", "name_suffix", "day", "start_hour",
            "duration", "notes", "instructors", "instructor_ids", "room_codes", "room"]


def _same(a: pd.DataFrame, b: pd.DataFrame, *, ids: bool = True) -> bool:
    """Same classes; with ``ids=False`` a re-inserted row may carry a new id."""
    columns = COMPARED if ids else COMPARED[1:]

    def rows(frame):
        return sorted(map(str, frame[columns].values.tolist()))

    return rows(a) == rows(b)


def test_snapshots_save_compare_restore_and_lock(database):
    """Runs on a term of its own, opened and locked here, after the flow above
    has locked its own: only one term may be open at a time."""
    year, period = 2027, tdb.WINTER
    assert tdb.open_term(year, period, 2025, tdb.WINTER, "c@ihu.gr") == ""
    assert "όνομα" in tdb.save_snapshot(year, period, "  ", None, "c@ihu.gr")
    assert tdb.save_snapshot(year, period, "Επιλογή Α", "όπως πέρυσι", "c@ihu.gr") == ""
    assert "ήδη" in tdb.save_snapshot(year, period, "Επιλογή Α", None, "c@ihu.gr")
    assert "ανοιχτό" in tdb.save_snapshot(2025, tdb.WINTER, "Χ", None, "c@ihu.gr")

    snapshots = tdb.list_snapshots(year, period)
    assert snapshots["name"].tolist() == ["Επιλογή Α"]
    a_id = int(snapshots["id"].iloc[0])
    original = tdb.load_term(year, period)
    assert int(snapshots["row_count"].iloc[0]) == len(original)
    snapshot_a = tdb.load_snapshot(a_id)
    assert _same(snapshot_a, original)
    # The snapshot feeds the same conflict check as the working term.
    assert len(tdb.conflicts(snapshot_a)) == len(tdb.conflicts(original))

    # Rearrange: move one row, delete one, add one.
    placed = original[original["placed"]]
    moved, gone = placed.iloc[0], placed.iloc[1]
    new_day = 5 if int(moved["day_number"]) != 5 else 4
    assert tdb.update_class(
        int(moved["id"]), examino=int(moved["examino"]), section=moved["section"],
        instructor_ids=list(moved["instructor_ids"]), room_codes=["ΤΣ1"],
        day=new_day, start_hour=int(moved["start_hour"]), duration=int(moved["duration"]),
        name_suffix=moved["name_suffix"], notes=moved["notes"], author="c@ihu.gr",
    ) == ""
    assert tdb.delete_class(int(gone["id"]), "c@ihu.gr") == ""
    assert tdb.add_class(
        year, period, examino=1, course_code="ΓΕΝ001", section="Φ",
        instructor_ids=[], room_codes=["301"], day=2, start_hour=18, duration=1, author="c@ihu.gr",
    ) == ""
    # ΣΥΓ017 taught as the 2018 course: a version keeps the row's title choice.
    syg = original[(original["course_code"] == "ΣΥΓ017") & (original["section"] == "Θ")].iloc[0]
    assert tdb.update_class(
        int(syg["id"]), examino=int(syg["examino"]), section="Θ",
        instructor_ids=list(syg["instructor_ids"]), room_codes=list(syg["room_codes"]),
        day=int(syg["day_number"]), start_hour=int(syg["start_hour"]), duration=int(syg["duration"]),
        name_suffix=syg["name_suffix"], notes=syg["notes"], author="c@ihu.gr", name_curriculum=2018,
    ) == ""
    assert tdb.save_snapshot(year, period, "Επιλογή Β", None, "c@ihu.gr") == ""
    b_id = int(tdb.list_snapshots(year, period)["id"].iloc[-1])
    snapshot_b = tdb.load_snapshot(b_id)
    assert snapshot_b[snapshot_b["id"] == syg["id"]].iloc[0]["course_name"] == (
        "Οργάνωση Εργοταξίου και Δομικές Μηχανές"
    )

    diff = tdb.compare_frames(tdb.load_snapshot(a_id), snapshot_b)
    assert sorted(diff["Μεταβολή"]) == sorted([tdb.CHANGED, tdb.CHANGED, tdb.REMOVED, tdb.ADDED])
    changes = diff[diff["Μεταβολή"] == tdb.CHANGED]
    changes = changes.set_index(changes["Μάθημα"].str.split().str[0])  # by course code
    assert changes.loc["ΣΥΓ017", "Τι"] == "Τίτλος"
    change = changes.loc[moved["course_code"]]
    assert change["Τι"].startswith("Ώρα") and "Αίθουσα" in change["Τι"]

    # Restoring Α brings the term back. The deleted row returns with a new id,
    # which Α adopts, so the two compare as identical — not as removed + added.
    assert tdb.restore_snapshot(a_id, "c@ihu.gr") == ""
    restored = tdb.load_term(year, period)
    assert len(restored) == len(original)
    assert _same(restored, snapshot_a, ids=False)
    assert restored[restored["id"] == syg["id"]].iloc[0]["course_name"] == (
        "Προγραμματισμός και Διαχείριση Τεχνικών Έργων"
    )
    assert int(gone["id"]) not in set(restored["id"])
    assert tdb.compare_frames(tdb.load_snapshot(a_id), restored).empty
    assert _same(tdb.load_snapshot(a_id), restored)
    # Restoring again changes nothing.
    assert tdb.restore_snapshot(a_id, "c@ihu.gr") == ""
    assert tdb.load_term(year, period)["id"].tolist() == restored["id"].tolist()

    assert tdb.delete_snapshot(b_id, "c@ihu.gr") == ""
    assert tdb.list_snapshots(year, period)["name"].tolist() == ["Επιλογή Α"]

    assert tdb.lock_term(year, period, "c@ihu.gr") == ""
    assert tdb.list_snapshots(year, period).empty
    assert "δεν υπάρχει" in tdb.restore_snapshot(a_id, "c@ihu.gr")
    changes = tdb.term_changes(year, period)
    assert set(changes["action"]) >= {tdb.SNAPSHOT_ACTION, tdb.RESTORED_ACTION}
    assert "εκδοχές" in changes["detail"].iloc[0]


def test_action_check_accepts_snapshot_actions_after_migration(database):
    """An installation from before 2026-09-26 has the old CHECK; starting replaces it."""
    engine = db.get_engine()
    old = tuple(a for a in tdb.ACTIONS if a not in (tdb.SNAPSHOT_ACTION, tdb.RESTORED_ACTION))
    with engine.begin() as conn:
        # Rows the old CHECK would reject, or re-adding it fails.
        conn.execute(
            sa_text(f"DELETE FROM {tdb.CHANGES_TABLE} WHERE action = ANY(:new)"),
            {"new": [tdb.SNAPSHOT_ACTION, tdb.RESTORED_ACTION]},
        )
        conn.execute(sa_text(f"ALTER TABLE {tdb.CHANGES_TABLE} DROP CONSTRAINT timetable_changes_action_check"))
        conn.execute(sa_text(
            f"ALTER TABLE {tdb.CHANGES_TABLE} ADD CONSTRAINT timetable_changes_action_check "
            f"CHECK (action IN {old})"
        ))
    engine.dispose()
    db.get_engine.clear()
    engine = db.get_engine()
    with engine.begin() as conn:
        conn.execute(
            sa_text(f"INSERT INTO {tdb.CHANGES_TABLE} (year, period, action) VALUES (2025, :p, :a)"),
            {"p": tdb.WINTER, "a": tdb.RESTORED_ACTION},
        )
        names = [row[0] for row in conn.execute(sa_text(
            "SELECT conname FROM pg_constraint WHERE conrelid = CAST(:t AS regclass) AND contype = 'c'"
        ), {"t": tdb.CHANGES_TABLE})]
    assert names == ["timetable_changes_action_check"]


def test_page_renders(database):
    """The page script runs end to end against the seeded database.

    ``st.user`` is logged out under AppTest, so this exercises the public
    views and the read-only side of Προετοιμασία; the writes are covered above.
    """
    from streamlit.testing.v1 import AppTest

    page = ROOT / "streamlit" / "app_pages" / "8_📅_weekly_timetable.py"
    app = AppTest.from_file(str(page), default_timeout=300)
    app.run()
    assert not app.exception, [str(e.value) for e in app.exception]
    assert not app.error, [e.value for e in app.error]
    assert len(app.tabs) == 8
    assert app.selectbox[0].label == "Εξάμηνο:"


def test_curriculum_column_is_dropped_on_start(database):
    """An installation from before 2026-09-16 still has the column; starting drops it."""
    engine = db.get_engine()
    with engine.begin() as conn:
        conn.execute(sa_text(f"ALTER TABLE {tdb.CLASSES_TABLE} ADD COLUMN IF NOT EXISTS curriculum INTEGER"))
    engine.dispose()
    db.get_engine.clear()
    engine = db.get_engine()
    with engine.connect() as conn:
        columns = {
            row[0]
            for row in conn.execute(
                sa_text("SELECT column_name FROM information_schema.columns WHERE table_name = :t"),
                {"t": tdb.CLASSES_TABLE},
            )
        }
    assert "course_code" in columns and "curriculum" not in columns
    assert not tdb.load_term(2025, tdb.WINTER).empty


def test_page_toggle_offers_every_semester(database, monkeypatch):
    """A coordinator on an open term: the toggle widens the course list."""
    import auth
    from streamlit.testing.v1 import AppTest

    if not tdb.open_terms():
        assert tdb.open_term(2026, tdb.SPRING, 2025, tdb.SPRING, "c@ihu.gr") == ""
    monkeypatch.setattr(auth, "is_authorized", lambda: True)
    monkeypatch.setattr(db, "is_coordinator", lambda email: True)

    page = ROOT / "streamlit" / "app_pages" / "8_📅_weekly_timetable.py"
    app = AppTest.from_file(str(page), default_timeout=300)
    app.run()
    assert not app.exception, [str(e.value) for e in app.exception]
    default = list(app.selectbox(key="add_course").options)
    assert default
    app.toggle(key="add_all_semesters").set_value(True).run()
    assert not app.exception, [str(e.value) for e in app.exception]
    widened = list(app.selectbox(key="add_course").options)
    assert len(widened) > len(default)
    # Winter courses are offered in a spring term once the toggle is on.
    assert any(option.startswith("ΔΟΜ007 ") for option in widened)


def test_page_versions_section(database, monkeypatch):
    """A coordinator on an open term with two versions: the section renders,
    compares, and switching the comparison does not raise."""
    import auth
    from streamlit.testing.v1 import AppTest

    if not tdb.open_terms():
        assert tdb.open_term(2026, tdb.SPRING, 2025, tdb.SPRING, "c@ihu.gr") == ""
    year, period = tdb.open_terms()[0]
    for name in ("Επιλογή Α", "Επιλογή Β"):
        if name not in set(tdb.list_snapshots(year, period)["name"]):
            assert tdb.save_snapshot(year, period, name, None, "c@ihu.gr") == ""
    monkeypatch.setattr(auth, "is_authorized", lambda: True)
    monkeypatch.setattr(db, "is_coordinator", lambda email: True)

    page = ROOT / "streamlit" / "app_pages" / "8_📅_weekly_timetable.py"
    app = AppTest.from_file(str(page), default_timeout=300)
    app.run()
    assert not app.exception, [str(e.value) for e in app.exception]
    left = app.selectbox(key="compare_left")
    assert len(left.options) == 3  # the current timetable and the two versions
    app.selectbox(key="compare_right").set_value(left.value).run()
    assert not app.exception, [str(e.value) for e in app.exception]
    assert any("διαφορετικές" in i.value for i in app.info)
