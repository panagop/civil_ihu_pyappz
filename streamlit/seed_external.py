"""Load the historical ``external_<year>.xlsx`` workbooks into Postgres.

The database is only reachable from inside Railway, so this cannot be run from
a developer machine — it is called by :func:`db.bootstrap` when the app starts
there. It is idempotent: a year that already has rows is skipped, and the
inserts themselves are ``ON CONFLICT DO NOTHING``.

Only the three columns that are *decisions* are stored (characterisation,
reasoning, and who/where). Everything else in the workbook is a copy of that
year's ΑΠΕΛΛΑ export and is joined back in at display time.
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd
from sqlalchemy import Engine, text

ROOT = Path(__file__).resolve().parents[1]
BY_YEAR_DIR = ROOT / "files" / "mitroa" / "mitroa_by_year"

HEADER_ROW = 8  # 0-based row holding the α/α ... Αιτιολόγηση συνάφειας header

# Canonical name -> the header text we look for. Headers are matched after
# stripping accents, case and every non-letter, because the workbooks are
# inconsistent: the same column appears as both "Χαρακτηρισμός" and the
# soft-hyphenated "Χαρακτη-ρισμός" depending on the year.
COLUMNS = {
    "elector_id": "Κωδικός Χρήστη",
    "characterization": "Χαρακτηρισμός",
    "reasoning": "Αιτιολόγηση συνάφειας",
}
# The workbooks spell the two characterisations inconsistently
CHARACTERIZATION_ALIASES = {"ΙΔΙΟ": "ΙΔΙΟΥ", "ΣΥΝΑΦΕΣ": "ΣΥΝΑΦΟΥΣ"}
VALID_CHARACTERIZATIONS = {"ΙΔΙΟΥ", "ΣΥΝΑΦΟΥΣ"}

INSERT_SQL = text(
    """
    INSERT INTO external_electors
        (year, field_code, elector_id, characterization, reasoning)
    VALUES (:year, :field_code, :elector_id, :characterization, :reasoning)
    ON CONFLICT (year, field_code, elector_id) DO NOTHING
    """
)


def _norm_header(value: object) -> str:
    """Accent-free, case-free, letters-only form of a header cell."""
    decomposed = unicodedata.normalize("NFD", str(value))
    stripped = "".join(c for c in decomposed if not unicodedata.combining(c))
    return re.sub(r"[^0-9a-zα-ω]", "", stripped.casefold().replace("ς", "σ"))


def _column_map(header: list) -> dict[str, int]:
    """Map canonical name -> position, for the columns we need."""
    positions = {_norm_header(cell): i for i, cell in enumerate(header)}
    return {
        key: positions[_norm_header(label)]
        for key, label in COLUMNS.items()
        if _norm_header(label) in positions
    }


def read_workbook(path: Path, year: int) -> list[dict]:
    """Every stored decision in one workbook, as insert-ready dicts."""
    sheets = pd.read_excel(path, sheet_name=None, header=None)
    rows: list[dict] = []
    for name, raw in sheets.items():
        if not str(name).strip().isdigit():
            continue  # not a γνωστικό αντικείμενο sheet
        field_code = int(str(name).strip())
        columns = _column_map(raw.iloc[HEADER_ROW].tolist())
        missing = set(COLUMNS) - set(columns)
        if missing:
            raise ValueError(f"{path.name}, φύλλο {name}: λείπουν στήλες {missing}")

        body = raw.iloc[HEADER_ROW + 1:]
        for _, record in body.iterrows():
            elector = record.iat[columns["elector_id"]]
            if pd.isna(elector):
                continue
            characterization = str(record.iat[columns["characterization"]]).strip()
            characterization = CHARACTERIZATION_ALIASES.get(
                characterization, characterization
            )
            if characterization not in VALID_CHARACTERIZATIONS:
                raise ValueError(
                    f"{path.name}, φύλλο {name}, εκλέκτορας {elector}: "
                    f"άγνωστος χαρακτηρισμός {characterization!r}"
                )
            reasoning = record.iat[columns["reasoning"]]
            rows.append(
                {
                    "year": year,
                    "field_code": field_code,
                    "elector_id": int(elector),
                    "characterization": characterization,
                    "reasoning": "" if pd.isna(reasoning) else str(reasoning).strip(),
                }
            )
    return rows


def seed_historical_years(engine: Engine) -> str:
    """Insert every ``external_<year>.xlsx`` that has no rows yet."""
    loaded, skipped = [], []
    for path in sorted(BY_YEAR_DIR.glob("external_*.xlsx")):
        match = re.search(r"external_(\d{4})", path.stem)
        if not match:
            continue
        year = int(match.group(1))

        with engine.connect() as conn:
            already = conn.execute(
                text("SELECT 1 FROM external_electors WHERE year = :year LIMIT 1"),
                {"year": year},
            ).first()
        if already:
            skipped.append(str(year))
            continue

        rows = read_workbook(path, year)
        with engine.begin() as conn:
            conn.execute(INSERT_SQL, rows)
        loaded.append(f"{year} ({len(rows)} εγγραφές)")

    parts = []
    if loaded:
        parts.append("φορτώθηκαν: " + ", ".join(loaded))
    if skipped:
        parts.append("υπήρχαν ήδη: " + ", ".join(skipped))
    return "Βάση δεδομένων — " + ("· ".join(parts) if parts else "κανένα αρχείο")
