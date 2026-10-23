"""Tests for the per-subject CSV export of the external electors (page 5)."""

from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "streamlit"))

from mitroa import table  # noqa: E402

SAMPLE = ROOT / "files" / "mitroa" / "csv_test" / "Θεμελιώσεις - Αντιστηρίξεις.csv"


def _entry(code, field, rows):
    df = pd.DataFrame(rows, columns=[table.CHARAKTIRISMOS_COL, table.ID_COL, "Επώνυμο"])
    return {"code": code, "field": field, "domain": "", "df": df}


def test_subject_csv_matches_the_sample_layout():
    entry = _entry(575, "Θεμελιώσεις - Αντιστηρίξεις", [
        ("ΙΔΙΟΥ", 11665, "ΓΑΡΙΝΗ"),
        ("ΣΥΝΑΦΟΥΣ", "8808", " ASIMAKI "),
    ])
    assert table.subject_csv(entry["df"]) == (
        "ΚΩΔΙΚΟΣ ΧΡΗΣΤΗ,ΕΠΩΝΥΜΟ,ΣΥΝΑΦΕΙΑ\r\n11665,ΓΑΡΙΝΗ,1\r\n8808,ASIMAKI,0\r\n"
    ).encode()
    # Same header, line ending and encoding (no BOM) as the file given as the model
    assert SAMPLE.read_bytes().startswith(table.subject_csv(entry["df"])[:40])


def test_unknown_characterisation_raises_with_the_subject():
    workbook = {"575": _entry(575, "Θεμελιώσεις", [("", 1, "Χ")])}
    with pytest.raises(ValueError, match="575 — Θεμελιώσεις"):
        table.csv_archive(workbook)


def test_archive_holds_one_file_per_subject_with_safe_names():
    workbook = {
        "606": _entry(606, "Προγραμματισμός Η/Υ", [("ΙΔΙΟΥ", 1, "Α")]),
        "570": _entry(570, "Σχεδιασμός: Τεχνικές.", [("ΣΥΝΑΦΟΥΣ", 2, "Β")]),
    }
    names = zipfile.ZipFile(io.BytesIO(table.csv_archive(workbook))).namelist()
    assert names == ["Σχεδιασμός- Τεχνικές.csv", "Προγραμματισμός Η-Υ.csv"]
