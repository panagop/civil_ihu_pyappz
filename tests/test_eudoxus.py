"""Tests for the Εύδοξος pipeline that need neither a database nor the network.

The Postgres instance is Railway-only and service.eudoxus.gr is a third party,
so what is pinned here is the import and the type coercion — the two places a
mistake would only show up in production.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "streamlit"))

import eudoxus_db as edb  # noqa: E402
import seed_eudoxus as seed  # noqa: E402
from eudoxus_client import availability_reason  # noqa: E402

WORKBOOK = seed.EUDOXUS_DIR / "eudoxus_books_2025-26.xlsx"


@pytest.fixture(scope="module")
def workbook():
    return seed.read_workbook(WORKBOOK)


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

    _, selections, _ = seed.read_workbook(WORKBOOK)
    wanted = {item["book_id"] for item in selections}
    have = {row["book_id"] for row in seed.read_catalogue(seed._latest_catalogue())}
    assert wanted <= have
    assert len(wanted) == 222
