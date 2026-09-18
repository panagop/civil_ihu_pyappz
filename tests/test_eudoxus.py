"""The Εύδοξος pipeline: the imports and type coercion without a database, then
the seed itself against a throwaway Postgres (``pgserver``, skipped where it is
not installed — see ``test_db_rename.py``).

service.eudoxus.gr is a third party, so nothing here touches the network.
"""

from __future__ import annotations

import io
import os
import re
import sys
from pathlib import Path

import pandas as pd
import pytest
from sqlalchemy import text as sa_text

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "streamlit"))

import eudoxus_db as edb  # noqa: E402
import seed_eudoxus as seed  # noqa: E402
import db  # noqa: E402
from eudoxus_client import availability_reason  # noqa: E402
from settings import get_secret  # noqa: E402

WORKBOOK = seed.EUDOXUS_DIR / "eudoxus_books_2025-26.xlsx"
CSV_EXPORT = seed.EUDOXUS_DIR / "Συγγράμματα ΕΥΔΟΞΟΣ 2026-2027.csv"


@pytest.fixture(scope="module")
def workbook():
    return seed.read_export(WORKBOOK)


@pytest.fixture(scope="module")
def csv_export():
    return seed.read_export(CSV_EXPORT)


def test_year_label():
    assert edb.year_label(2025) == "2025-26"
    assert edb.year_label(2029) == "2029-30"


def test_workbook_counts(workbook):
    courses, selections, inactive = workbook
    assert len(courses) == 102
    assert len(selections) == 291
    assert inactive == 0


def test_course_offering_key_is_unique(workbook):
    """(year, code, εξάμηνο) is the primary key — it must not collide."""
    courses, _, _ = workbook
    keys = {(course["course_code"], course["examino"]) for course in courses}
    assert len(keys) == len(courses)


def test_dom022_is_kept_twice(workbook):
    """ΔΟΜ022 runs in both the 7th and the 9th εξάμηνο, with its own book order."""
    courses, selections, _ = workbook
    offerings = sorted(
        course["examino"] for course in courses if course["course_code"] == "ΔΟΜ022"
    )
    assert offerings == [7, 9]

    def order(examino: int) -> dict[int, int]:
        return {
            item["book_id"]: item["priority"]
            for item in selections
            if item["course_code"] == "ΔΟΜ022" and item["examino"] == examino
        }

    seventh, ninth = order(7), order(9)
    assert set(seventh) == set(ninth)      # same three books
    assert seventh != ninth                # different priority order


def test_selection_key_is_unique(workbook):
    _, selections, _ = workbook
    keys = {
        (item["course_code"], item["examino"], item["book_id"]) for item in selections
    }
    assert len(keys) == len(selections)


def test_catalogue_rows_are_postgres_safe():
    """Nothing but str/int/bool/None may reach psycopg."""
    rows = seed.read_catalogue(seed._latest_catalogue())
    assert len(rows) == 222
    for row in rows:
        for key, value in row.items():
            assert value is None or isinstance(value, (str, int, bool)), (key, value)


def test_isbn_is_coerced_to_text():
    """Postgres has no assignment cast from bigint to text."""
    cleaned = edb._clean_book({"book_id": 1, "isbn": 9789606796142, "found": True})
    assert cleaned["isbn"] == "9789606796142"
    assert isinstance(cleaned["book_id"], int)


def test_clean_book_fills_every_bound_column():
    cleaned = edb._clean_book({"book_id": 7})
    assert set(cleaned) == set(edb._BOOK_KEYS)
    assert cleaned["found"] is True
    assert cleaned["active"] is None


def test_clean_book_survives_a_junk_year():
    assert edb._clean_book({"book_id": 1, "publication_year": "χ.χ."})["publication_year"] is None
    assert edb._clean_book({"book_id": 1, "publication_year": "2009"})["publication_year"] == 2009


@pytest.mark.parametrize(
    "row, expected",
    [
        ({"found": True, "active": True, "selectable": True}, ""),
        ({"found": True, "active": False, "selectable": True}, "ανενεργό (active=false)"),
        (
            {"found": True, "active": True, "selectable": False},
            "μη επιλέξιμο (selectable=false)",
        ),
        ({"found": False, "error": "boom"}, "boom"),
    ],
)
def test_availability_reason(row, expected):
    assert availability_reason(row) == expected


def test_catalogue_matches_the_workbook():
    """Every book in the list has a catalogue row, or the browse tab shows a gap."""

    _, selections, _ = seed.read_export(WORKBOOK)
    wanted = {item["book_id"] for item in selections}
    have = {row["book_id"] for row in seed.read_catalogue(seed._latest_catalogue())}
    assert wanted <= have
    assert len(wanted) == 222


# --------------------------------------------------------------------------
# Which file covers which year
# --------------------------------------------------------------------------

@pytest.mark.parametrize(
    "name, expected",
    [
        ("eudoxus_books_2025-26.xlsx", 2025),
        ("Συγγράμματα ΕΥΔΟΞΟΣ 2026-2027.csv", 2026),
        # The catalogue's date stamp is not a year range, and neither is a
        # second half that does not follow the first.
        ("eudoxus_catalogue_20260909.csv", None),
        ("eudoxus_books_2025-28.xlsx", None),
        ("eudoxus.py", None),
    ],
)
def test_export_year(name, expected):
    assert seed.export_year(Path(name)) == expected


def test_exports_are_found_in_year_order():
    """2026-27 is seeded with 2025-26 as its baseline, so order matters."""
    found = seed.find_exports()
    assert [year for year, _ in found] == sorted(year for year, _ in found)
    assert dict(found).keys() >= {2025, 2026}


def test_exports_are_named_not_globbed():
    """Working copies land beside the real file; only the listed one is read."""
    for year, path in seed.find_exports():
        assert path.exists(), path
        assert seed.export_year(path) == year
    assert seed.EXPORTS[2026] == "Συγγράμματα ΕΥΔΟΞΟΣ 2026-2027.csv"


# --------------------------------------------------------------------------
# The 2026-27 list
# --------------------------------------------------------------------------

def test_csv_export_counts(csv_export):
    courses, selections, inactive = csv_export
    assert len(courses) == 105
    assert len(selections) == 305
    assert inactive == 0


def test_csv_export_keys_are_unique(csv_export):
    courses, selections, _ = csv_export
    assert len({(c["course_code"], c["examino"]) for c in courses}) == len(courses)
    assert len(
        {(s["course_code"], s["examino"], s["book_id"]) for s in selections}
    ) == len(selections)


def test_csv_export_is_typed_for_postgres(csv_export):
    """The CSV reader must coerce exactly as the Excel one does."""
    courses, selections, _ = csv_export
    for course in courses:
        assert isinstance(course["examino"], int)
        assert course["teacher"] is None or isinstance(course["teacher"], str)
    for selection in selections:
        assert isinstance(selection["book_id"], int)
        assert isinstance(selection["priority"], int)


def test_2026_is_seeded_as_the_open_year():
    assert seed.OPEN_SEED_YEARS == {2026: 2025}


def test_export_headers_are_what_the_seed_reads():
    """What the app writes, the seed must read back — same headers, same order."""
    assert list(edb.EXPORT_COLUMNS) == list(seed.COLUMNS)


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
    name = "eudoxus_" + re.sub(r"\W", "_", tmp_path_factory.mktemp("x").name).lower()[:40]
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


def test_seed_loads_both_years(database):
    assert edb.stored_years() == [2026, 2025]
    assert len(edb.courses_for_year(2025)) == 102
    assert len(edb.courses_for_year(2026)) == 105
    assert len(edb.load_year(2026)) == 305
    assert len(edb.book_ids_for_year(2026)) == 228


def test_2026_arrives_open_against_2025(database):
    """The year people are still working on, not history."""
    state = edb.year_state(2026)
    assert state["status"] == edb.OPEN
    assert state["baseline_year"] == 2025
    assert state["locked_at"] is None
    assert edb.is_editable(2026)
    assert edb.open_years() == [2026]

    past = edb.year_state(2025)
    assert past["status"] == edb.LOCKED
    assert past["baseline_year"] is None


def test_the_coordinator_sees_the_net_change(database):
    diff = edb.changes_vs_baseline(2026)
    counts = diff["action"].value_counts().to_dict()
    assert counts.get(edb.ADD) == 35
    assert counts.get(edb.REMOVE) == 21
    # Three course offerings exist only in 2026-27.
    new_courses = set(map(tuple, edb.courses_for_year(2026)[["course_code", "examino"]].values))
    old_courses = set(map(tuple, edb.courses_for_year(2025)[["course_code", "examino"]].values))
    assert new_courses - old_courses == {("ΔΟΜ037", 2), ("ΔΟΜ038", 3), ("ΣΥΓ008", 8)}


def test_the_new_books_are_simply_unchecked(database):
    """18 book codes are new: unknown to the catalogue is not a problem report."""
    frame = edb.load_year(2026)
    assert frame[frame["checked_at"].isna()]["book_id"].nunique() == 18
    # The catalogue dump found 9 books that cannot be chosen again, all of
    # them in 2025-26: the 2026-27 list already dropped every one. A year
    # carrying unknown codes must not inflate that count — reading `~` off an
    # object column once turned these 9 into 210.
    assert edb.unavailable_for_year(2025)["book_id"].nunique() == 9
    assert edb.unavailable_for_year(2026).empty


def test_unusable_mask_survives_an_unknown_book():
    """An object-dtype column is what a LEFT JOIN with a miss gives pandas."""
    frame = pd.DataFrame(
        {
            "found": pd.Series([True, True, False, None], dtype=object),
            "active": pd.Series([True, False, True, None], dtype=object),
            "selectable": pd.Series([True, True, True, None], dtype=object),
        }
    )
    assert list(edb.unusable_mask(frame)) == [False, True, True, True]


def test_seeding_is_idempotent(database):
    status = seed.seed_eudoxus(db.get_engine())
    assert "2025-26" in status and "2026-27" in status
    assert "φορτώθηκαν" not in status
    assert edb.year_state(2026)["status"] == edb.OPEN


def test_an_open_year_is_not_seeded_twice(database):
    """A coordinator's open year must not be joined by a seeded one."""
    assert seed._seed_status(db.get_engine(), 2026) == (edb.LOCKED, None)


def test_csv_export_round_trips_the_official_file(database):
    """The 2026-27 list exported from the app is the file the department uploaded."""
    data = edb.export_csv(edb.load_year(2026))
    assert not data.startswith(b"\xef\xbb\xbf")
    text = data.decode("utf-8")
    assert "\r\n" in text and "\n" not in text.replace("\r\n", "")
    ours = pd.read_csv(io.BytesIO(data))
    theirs = pd.read_csv(CSV_EXPORT)
    assert list(ours.columns) == list(theirs.columns)
    key = ["Κωδικός μαθήματος", "Εξάμηνο", "Book id"]
    ours = ours.sort_values(key).reset_index(drop=True)
    theirs = theirs.sort_values(key).reset_index(drop=True)
    pd.testing.assert_frame_equal(ours, theirs, check_dtype=False)
    assert len(ours) == 305


def test_a_locked_year_cannot_be_deleted(database):
    message = edb.delete_year(2025, "test@ihu.gr")
    assert "κλειδωμένο" in message
    assert len(edb.courses_for_year(2025)) == 102


def test_delete_open_year_then_reseed_from_file(database):
    """A year opened by mistake as a copy is replaced by the submitted list.

    Last in the module: it empties and refills the open year.
    """
    engine = db.get_engine()
    with engine.begin() as conn:
        conn.execute(
            sa_text(
                f"INSERT INTO {edb.CHANGES_TABLE} "
                "(year, course_code, examino, book_id, action, author) "
                f"VALUES (2026, 'ΔΟΜ037', 2, 1, '{edb.ADD}', 'test@ihu.gr')"
            )
        )
    message = edb.delete_year(2026, "test@ihu.gr")
    assert message.startswith("Διαγράφηκε")
    assert "305 επιλογές" in message and "1 καταγραφές" in message
    assert edb.year_state(2026) is None
    assert edb.stored_years() == [2025]
    assert edb.year_changes(2026).empty
    assert edb.delete_year(2026, "test@ihu.gr").endswith("δεν υπάρχει στη βάση.")

    status = seed.seed_eudoxus(engine)
    assert "φορτώθηκαν: 2026-27" in status
    assert len(edb.courses_for_year(2026)) == 105
    assert len(edb.load_year(2026)) == 305
    state = edb.year_state(2026)
    assert state["status"] == edb.OPEN and state["baseline_year"] == 2025
