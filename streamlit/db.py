"""Postgres access for the μητρώα tables.

The database lives on Railway and is reachable **only from inside Railway**
(no public TCP proxy — see CLAUDE.md, "Database"). Locally and on Streamlit
Cloud there is no ``DATABASE_URL``, so :func:`get_engine` returns ``None`` and
every caller must degrade gracefully rather than crash: the file-backed tabs of
page 5 keep working everywhere, only the database-backed ones go quiet.

Because nothing outside Railway can reach the database, the schema and the
historical 2025 data are installed by the app itself on first start — see
:func:`bootstrap`. Both steps are idempotent.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st
from sqlalchemy import Engine, create_engine, text

from settings import get_secret

CHARACTERIZATIONS = ("ΙΔΙΟΥ", "ΣΥΝΑΦΟΥΣ")

ROOT = Path(__file__).resolve().parents[1]
PROFESSORS_DIR = ROOT / "files" / "mitroa" / "professors_tables"
# Which ΑΠΕΛΛΑ export a stored year joins against. Kept as a file rather than a
# table because it must also resolve where there is no database (locally, on
# Streamlit Cloud), and because the export filenames are date stamps, not years.
SNAPSHOTS_CSV = ROOT / "files" / "mitroa" / "registry_snapshots.csv"

# One row per (year, γνωστικό αντικείμενο, elector). Everything else about the
# elector — name, φορέας, βαθμίδα, ΦΕΚ — is joined in from that year's ΑΠΕΛΛΑ
# export (a parquet file in the repo), so it is never duplicated here.
SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS external_electors (
    year             INTEGER NOT NULL,
    field_code       INTEGER NOT NULL,
    elector_id       INTEGER NOT NULL,
    characterization TEXT    NOT NULL CHECK (characterization IN ('ΙΔΙΟΥ', 'ΣΥΝΑΦΟΥΣ')),
    reasoning        TEXT    NOT NULL,
    created_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (year, field_code, elector_id)
);

-- "in which αντικείμενα does this person appear?" — the primary key only
-- covers the (year, field_code) prefix.
CREATE INDEX IF NOT EXISTS external_electors_year_elector_idx
    ON external_electors (year, elector_id);
"""


@st.cache_resource
def get_engine() -> Engine | None:
    """SQLAlchemy engine, or None when no database is configured.

    ``pool_pre_ping`` matters here: Railway recycles idle connections, and a
    Streamlit process can sit untouched for hours between visitors.
    """
    url = get_secret("DATABASE_URL")
    if not url:
        return None
    # Railway hands out a plain postgresql:// URL; SQLAlchemy 2 needs the
    # driver spelled out to pick psycopg 3 over the (absent) psycopg2.
    if url.startswith("postgresql://"):
        url = url.replace("postgresql://", "postgresql+psycopg://", 1)
    elif url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql+psycopg://", 1)
    return create_engine(url, pool_pre_ping=True)


def is_available() -> bool:
    return get_engine() is not None


@st.cache_resource
def bootstrap() -> str:
    """Create the schema and load the historical years. Runs once per process.

    Returns a short human-readable status line. Never raises: a database
    problem must not take down the file-backed pages.
    """
    engine = get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων (δεν έχει οριστεί DATABASE_URL)."
    try:
        with engine.begin() as conn:
            conn.execute(text(SCHEMA_SQL))
        from seed_external import seed_historical_years

        status = seed_historical_years(engine)
    except Exception as exc:  # noqa: BLE001 - reported, not raised
        status = f"Σφάλμα βάσης: {exc}"
    # Also to stdout: the database is unreachable from outside Railway, so the
    # deployment logs are the only place this can be checked.
    print(f"[db.bootstrap] {status}", flush=True)
    return status


def load_external_electors(year: int) -> pd.DataFrame:
    """The stored rows for one year, ordered ΙΔΙΟΥ first then by elector id."""
    engine = get_engine()
    if engine is None:
        return pd.DataFrame()
    query = text(
        """
        SELECT year, field_code, elector_id, characterization, reasoning
        FROM external_electors
        WHERE year = :year
        ORDER BY field_code,
                 characterization <> 'ΙΔΙΟΥ',
                 elector_id
        """
    )
    with engine.connect() as conn:
        return pd.read_sql(query, conn, params={"year": year})


def registry_file_for_year(year: int) -> Path | None:
    """The ΑΠΕΛΛΑ export a given year's decisions should be joined against."""
    mapping = pd.read_csv(SNAPSHOTS_CSV).set_index("year")["filename"]
    if year not in mapping.index:
        return None
    path = PROFESSORS_DIR / mapping.loc[year]
    return path if path.exists() else None


def stored_years() -> list[int]:
    """Years that have rows in the database, newest first."""
    engine = get_engine()
    if engine is None:
        return []
    with engine.connect() as conn:
        rows = conn.execute(
            text("SELECT DISTINCT year FROM external_electors ORDER BY year DESC")
        )
        return [row[0] for row in rows]
