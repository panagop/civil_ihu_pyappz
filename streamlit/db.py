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

from settings import get_secret, get_secret_list

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

-- A year under preparation. Its table is never stored while open: it is the
-- baseline year plus the accepted proposals, computed on read. Only
-- finalisation writes rows into external_electors.
CREATE TABLE IF NOT EXISTS year_status (
    year          INTEGER PRIMARY KEY,
    status        TEXT    NOT NULL CHECK (status IN ('ΑΝΟΙΧΤΟ', 'ΚΛΕΙΔΩΜΕΝΟ')),
    baseline_year INTEGER,
    opened_by     TEXT,
    opened_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
    locked_by     TEXT,
    locked_at     TIMESTAMPTZ
);

-- Proposed changes to the open year. Two members proposing on the same elector
-- produce two rows rather than overwriting each other; the coordinator resolves
-- the conflict when deciding.
CREATE TABLE IF NOT EXISTS proposals (
    id               BIGSERIAL PRIMARY KEY,
    year             INTEGER NOT NULL,
    field_code       INTEGER NOT NULL,
    elector_id       INTEGER NOT NULL,
    action           TEXT NOT NULL CHECK (action IN
                         ('ΠΡΟΣΘΗΚΗ', 'ΑΦΑΙΡΕΣΗ', 'ΧΑΡΑΚΤΗΡΙΣΜΟΣ', 'ΑΙΤΙΟΛΟΓΗΣΗ')),
    characterization TEXT CHECK (characterization IN ('ΙΔΙΟΥ', 'ΣΥΝΑΦΟΥΣ')),
    reasoning        TEXT,
    note             TEXT NOT NULL,
    author           TEXT NOT NULL,
    created_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
    status           TEXT NOT NULL DEFAULT 'ΕΚΚΡΕΜΕΙ' CHECK (status IN
                         ('ΕΚΚΡΕΜΕΙ', 'ΕΓΚΡΙΘΗΚΕ', 'ΑΠΟΡΡΙΦΘΗΚΕ', 'ΑΠΟΣΥΡΘΗΚΕ')),
    decided_by       TEXT,
    decided_at       TIMESTAMPTZ,
    decision_note    TEXT,
    -- An addition or a re-characterisation is meaningless without the value it
    -- proposes; the database refuses one rather than trusting the UI.
    CONSTRAINT proposals_needs_characterization CHECK (
        action NOT IN ('ΠΡΟΣΘΗΚΗ', 'ΧΑΡΑΚΤΗΡΙΣΜΟΣ') OR characterization IS NOT NULL),
    CONSTRAINT proposals_needs_reasoning CHECK (
        action NOT IN ('ΠΡΟΣΘΗΚΗ', 'ΑΙΤΙΟΛΟΓΗΣΗ') OR reasoning IS NOT NULL),
    CONSTRAINT proposals_note_not_blank CHECK (btrim(note) <> '')
);

CREATE INDEX IF NOT EXISTS proposals_year_field_idx ON proposals (year, field_code);
CREATE INDEX IF NOT EXISTS proposals_year_status_idx ON proposals (year, status);
"""

PENDING = "ΕΚΚΡΕΜΕΙ"
ACCEPTED = "ΕΓΚΡΙΘΗΚΕ"
REJECTED = "ΑΠΟΡΡΙΦΘΗΚΕ"
WITHDRAWN = "ΑΠΟΣΥΡΘΗΚΕ"

OPEN = "ΑΝΟΙΧΤΟ"
LOCKED = "ΚΛΕΙΔΩΜΕΝΟ"

ADD = "ΠΡΟΣΘΗΚΗ"
REMOVE = "ΑΦΑΙΡΕΣΗ"
RECHARACTERIZE = "ΧΑΡΑΚΤΗΡΙΣΜΟΣ"
REJUSTIFY = "ΑΙΤΙΟΛΟΓΗΣΗ"


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


# --------------------------------------------------------------------------
# Roles
# --------------------------------------------------------------------------

def is_coordinator(email: str | None) -> bool:
    """True for the people who may decide proposals and lock a year.

    Kept in a `coordinator_emails` setting rather than a users table: with two
    roles and a handful of people a table would need an admin screen to manage
    and would still need a way to appoint the first admin.
    """
    if not email:
        return False
    return email.strip().lower() in {
        item.lower() for item in get_secret_list("coordinator_emails")
    }


# --------------------------------------------------------------------------
# The year under preparation
# --------------------------------------------------------------------------

def year_state(year: int) -> dict | None:
    """The year_status row, or None if the year was never opened."""
    engine = get_engine()
    if engine is None:
        return None
    with engine.connect() as conn:
        row = conn.execute(
            text("SELECT * FROM year_status WHERE year = :year"), {"year": year}
        ).mappings().first()
    return dict(row) if row else None


def open_year(year: int, baseline_year: int, opened_by: str) -> str:
    """Start a working year based on an already finalised one.

    No rows are copied. The table is the baseline plus accepted proposals until
    :func:`finalize_year` writes it down — so a carried-forward elector always
    reflects the baseline, and there is one place where the year becomes real.
    """
    engine = get_engine()
    if engine is None:
        return "Δεν υπάρχει βάση δεδομένων."
    if year_state(year):
        return f"Το έτος {year} έχει ήδη ανοίξει."
    if baseline_year not in stored_years():
        return f"Το έτος βάσης {baseline_year} δεν υπάρχει στη βάση."
    with engine.begin() as conn:
        conn.execute(
            text(
                """
                INSERT INTO year_status (year, status, baseline_year, opened_by)
                VALUES (:year, :status, :baseline, :by)
                """
            ),
            {"year": year, "status": OPEN, "baseline": baseline_year, "by": opened_by},
        )
    return f"Το έτος {year} άνοιξε με βάση το {baseline_year}."


def working_electors(year: int) -> pd.DataFrame:
    """The current state of an open year: baseline + accepted proposals.

    Proposals are replayed in the order they were decided, so a later accepted
    proposal wins over an earlier one on the same elector.
    """
    state = year_state(year)
    if state is None:
        return pd.DataFrame()
    if state["status"] == LOCKED:
        return load_external_electors(year)

    table = load_external_electors(state["baseline_year"])
    if not table.empty:
        table = table.assign(year=year)
    rows = {
        (int(record.field_code), int(record.elector_id)): {
            "year": year,
            "field_code": int(record.field_code),
            "elector_id": int(record.elector_id),
            "characterization": record.characterization,
            "reasoning": record.reasoning,
        }
        for record in table.itertuples(index=False)
    }

    for change in list_proposals(year, status=ACCEPTED).itertuples(index=False):
        key = (int(change.field_code), int(change.elector_id))
        if change.action == ADD:
            rows[key] = {
                "year": year,
                "field_code": key[0],
                "elector_id": key[1],
                "characterization": change.characterization,
                "reasoning": change.reasoning,
            }
        elif change.action == REMOVE:
            rows.pop(key, None)
        elif key in rows:
            # A change to somebody who is no longer in the table is a no-op
            # rather than an error: the removal may have been accepted later.
            if change.action == RECHARACTERIZE:
                rows[key]["characterization"] = change.characterization
            elif change.action == REJUSTIFY:
                rows[key]["reasoning"] = change.reasoning

    if not rows:
        return pd.DataFrame(
            columns=["year", "field_code", "elector_id", "characterization", "reasoning"]
        )
    result = pd.DataFrame(list(rows.values()))
    return result.sort_values(
        ["field_code", "characterization", "elector_id"],
        key=lambda col: col.ne("ΙΔΙΟΥ") if col.name == "characterization" else col,
    ).reset_index(drop=True)


def finalize_year(year: int, locked_by: str) -> str:
    """Write the computed table into external_electors and freeze the year."""
    engine = get_engine()
    if engine is None:
        return "Δεν υπάρχει βάση δεδομένων."
    state = year_state(year)
    if state is None:
        return f"Το έτος {year} δεν έχει ανοίξει."
    if state["status"] == LOCKED:
        return f"Το έτος {year} είναι ήδη κλειδωμένο."

    pending = list_proposals(year, status=PENDING)
    if not pending.empty:
        return (
            f"Εκκρεμούν {len(pending)} προτάσεις — αποφασίστε τις πρώτα."
        )

    table = working_electors(year)
    if table.empty:
        return "Ο πίνακας είναι κενός — δεν κλειδώνεται."

    records = table.to_dict("records")
    with engine.begin() as conn:
        conn.execute(
            text(
                """
                INSERT INTO external_electors
                    (year, field_code, elector_id, characterization, reasoning)
                VALUES (:year, :field_code, :elector_id, :characterization, :reasoning)
                ON CONFLICT (year, field_code, elector_id) DO NOTHING
                """
            ),
            records,
        )
        conn.execute(
            text(
                """
                UPDATE year_status
                   SET status = :status, locked_by = :by, locked_at = now()
                 WHERE year = :year
                """
            ),
            {"status": LOCKED, "by": locked_by, "year": year},
        )
    return f"Το έτος {year} οριστικοποιήθηκε με {len(records)} εγγραφές."


# --------------------------------------------------------------------------
# Proposals
# --------------------------------------------------------------------------

def list_proposals(
    year: int, field_code: int | None = None, status: str | None = None
) -> pd.DataFrame:
    """Proposals for a year, oldest first (the order they are replayed in)."""
    engine = get_engine()
    if engine is None:
        return pd.DataFrame()
    clauses = ["year = :year"]
    params: dict = {"year": year}
    if field_code is not None:
        clauses.append("field_code = :field_code")
        params["field_code"] = field_code
    if status is not None:
        clauses.append("status = :status")
        params["status"] = status
    query = text(
        f"SELECT * FROM proposals WHERE {' AND '.join(clauses)} "
        "ORDER BY decided_at NULLS LAST, created_at, id"
    )
    with engine.connect() as conn:
        return pd.read_sql(query, conn, params=params)


def add_proposal(
    year: int,
    field_code: int,
    elector_id: int,
    action: str,
    note: str,
    author: str,
    characterization: str | None = None,
    reasoning: str | None = None,
) -> str:
    """Record one proposed change. Refused once the year is locked."""
    engine = get_engine()
    if engine is None:
        return "Δεν υπάρχει βάση δεδομένων."
    state = year_state(year)
    if state is None:
        return f"Το έτος {year} δεν έχει ανοίξει."
    if state["status"] == LOCKED:
        return f"Το έτος {year} είναι κλειδωμένο — δεν δέχεται προτάσεις."
    if not note or not note.strip():
        return "Η αιτιολόγηση της μεταβολής είναι υποχρεωτική."

    with engine.begin() as conn:
        conn.execute(
            text(
                """
                INSERT INTO proposals
                    (year, field_code, elector_id, action, characterization,
                     reasoning, note, author)
                VALUES (:year, :field_code, :elector_id, :action, :characterization,
                        :reasoning, :note, :author)
                """
            ),
            {
                "year": year,
                "field_code": field_code,
                "elector_id": elector_id,
                "action": action,
                "characterization": characterization,
                "reasoning": reasoning,
                "note": note.strip(),
                "author": author,
            },
        )
    return "Η πρόταση καταχωρήθηκε."


def decide_proposal(
    proposal_id: int, status: str, decided_by: str, decision_note: str | None = None
) -> str:
    """Accept or reject a pending proposal. Coordinator only — check first."""
    engine = get_engine()
    if engine is None:
        return "Δεν υπάρχει βάση δεδομένων."
    with engine.begin() as conn:
        updated = conn.execute(
            text(
                """
                UPDATE proposals
                   SET status = :status, decided_by = :by, decided_at = now(),
                       decision_note = :decision_note
                 WHERE id = :id AND status = :pending
                """
            ),
            {
                "status": status,
                "by": decided_by,
                "decision_note": decision_note,
                "id": proposal_id,
                "pending": PENDING,
            },
        ).rowcount
    if not updated:
        return "Η πρόταση δεν εκκρεμεί πλέον."
    return f"Η πρόταση {proposal_id}: {status.lower()}."


def withdraw_proposal(proposal_id: int, author: str) -> str:
    """Withdraw your own pending proposal."""
    engine = get_engine()
    if engine is None:
        return "Δεν υπάρχει βάση δεδομένων."
    with engine.begin() as conn:
        updated = conn.execute(
            text(
                """
                UPDATE proposals
                   SET status = :withdrawn, decided_by = :author, decided_at = now()
                 WHERE id = :id AND status = :pending AND author = :author
                """
            ),
            {
                "withdrawn": WITHDRAWN,
                "id": proposal_id,
                "pending": PENDING,
                "author": author,
            },
        ).rowcount
    return "Η πρόταση αποσύρθηκε." if updated else "Δεν βρέθηκε εκκρεμής πρότασή σας."
