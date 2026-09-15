"""The μητρώα tables against a real, throwaway Postgres.

The production database is reachable only from inside Railway, so until now a
schema change could be checked only by pushing and reading the deployment log.
``pgserver`` (a dev extra) bundles Postgres binaries and starts one in a temp
directory; these tests are skipped where it is not installed.

Two installations are exercised: one built under the pre-2026-09-15 names,
which must be renamed on start without losing a row, a constraint name or the
sequence position — and a fresh one, which must come up under the new names
with no rename and no backup copies.
"""

from __future__ import annotations

import os
import re
import sys
import zipfile
from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest
from sqlalchemy import create_engine, text

pgserver = pytest.importorskip("pgserver")

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "streamlit"))

import db  # noqa: E402
import seed_external  # noqa: E402
from settings import get_secret  # noqa: E402

# What the tables looked like before the rename is, by construction, the
# current schema without the prefix — that equivalence is what the rename
# asserts, so the test builds the legacy installation from it.
LEGACY_SCHEMA_SQL = db.SCHEMA_SQL.replace(db.TABLE_PREFIX, "")
LEGACY_MIGRATIONS_SQL = db.MIGRATIONS_SQL.replace(db.TABLE_PREFIX, "")
LEGACY_ROWS = {
    "external_electors": [
        (2025, 555, 1, "ΙΔΙΟΥ", "α"),
        (2025, 555, 2, "ΣΥΝΑΦΟΥΣ", "β"),
        (2025, 556, 1, "ΙΔΙΟΥ", "γ"),
    ],
    "proposals": [
        # (field, elector, action, characterization, reasoning, status)
        (555, 3, db.ADD, "ΙΔΙΟΥ", "δ", db.ACCEPTED),
        (555, 2, db.REMOVE, None, None, db.ACCEPTED),
        (556, 1, db.MODIFY, "ΣΥΝΑΦΟΥΣ", "ε", db.PENDING),
    ],
}


@pytest.fixture(scope="module")
def server(tmp_path_factory):
    instance = pgserver.get_server(tmp_path_factory.mktemp("pgdata"))
    yield instance
    instance.cleanup()


@pytest.fixture
def database(server, request):
    """A new, empty database wired into ``db`` through ``DATABASE_URL``."""
    name = re.sub(r"\W", "_", request.node.name).lower()[:60]
    server.psql(f"CREATE DATABASE {name}")
    url = server.get_uri(database=name)
    previous = os.environ.get("DATABASE_URL")
    os.environ["DATABASE_URL"] = url
    if get_secret("DATABASE_URL") != url:
        pytest.skip("DATABASE_URL is set in secrets.toml and shadows the test one")
    db.get_engine.clear()
    db.bootstrap.clear()
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


def install_legacy(url: str) -> None:
    """A database as the app left it before the rename, with a few rows."""
    engine = create_engine(url.replace("postgresql://", "postgresql+psycopg://"))
    with engine.begin() as conn:
        conn.execute(text(LEGACY_SCHEMA_SQL))
        conn.execute(text(LEGACY_MIGRATIONS_SQL))
        conn.execute(
            text(
                "INSERT INTO external_electors "
                "(year, field_code, elector_id, characterization, reasoning) "
                "VALUES (:y, :f, :e, :c, :r)"
            ),
            [
                {"y": y, "f": f, "e": e, "c": c, "r": r}
                for y, f, e, c, r in LEGACY_ROWS["external_electors"]
            ],
        )
        conn.execute(
            text(
                "INSERT INTO year_status (year, status, baseline_year, opened_by) "
                "VALUES (2026, :open, 2025, 'a@ihu.gr')"
            ),
            {"open": db.OPEN},
        )
        conn.execute(
            text(
                "INSERT INTO proposals (year, field_code, elector_id, action, "
                "characterization, reasoning, note, author, status, decided_by, "
                "decided_at) VALUES (2026, :f, :e, :a, :c, :r, 'σημείωση', "
                "'m@ihu.gr', :s, CASE WHEN :s = :pending THEN NULL ELSE 'c@ihu.gr' "
                "END, CASE WHEN :s = :pending THEN NULL ELSE now() END)"
            ),
            [
                {"f": f, "e": e, "a": a, "c": c, "r": r, "s": s, "pending": db.PENDING}
                for f, e, a, c, r, s in LEGACY_ROWS["proposals"]
            ],
        )
    engine.dispose()


def catalogue_names(engine) -> set[str]:
    """Every constraint, index and sequence name in the public schema."""
    with engine.connect() as conn:
        constraints = conn.execute(
            text(
                "SELECT conname FROM pg_constraint c JOIN pg_namespace n "
                "ON n.oid = c.connamespace WHERE n.nspname = 'public'"
            )
        ).scalars().all()
        indexes = conn.execute(
            text("SELECT indexname FROM pg_indexes WHERE schemaname = 'public'")
        ).scalars().all()
        sequences = conn.execute(
            text(
                "SELECT sequencename FROM pg_sequences WHERE schemaname = 'public'"
            )
        ).scalars().all()
    return set(constraints) | set(indexes) | set(sequences)


def test_legacy_installation_is_renamed_on_start(database, capsys):
    install_legacy(database)

    engine = db.get_engine()
    out = capsys.readouterr().out
    assert "Αποτυχία" not in out
    assert "Μετονομάστηκαν πίνακες" in out
    for old, new in db.LEGACY_TABLES.items():
        assert f"{old} → {new}" in out

    # The tables moved, and a verbatim copy of each was left behind.
    backups = {db.BACKUP_PREFIX + old for old in db.LEGACY_TABLES}
    assert set(db.mitroa_tables()) == set(db.LEGACY_TABLES.values()) | backups
    with engine.connect() as conn:
        for old, new in db.LEGACY_TABLES.items():
            assert conn.execute(text(f"SELECT to_regclass('{old}')")).scalar() is None
            live = conn.execute(text(f"SELECT count(*) FROM {new}")).scalar()
            copy = conn.execute(
                text(f"SELECT count(*) FROM {db.BACKUP_PREFIX}{old}")
            ).scalar()
            assert live == copy

    # Nothing in the catalogue still carries an old name: the μητρώα objects
    # are all mitroa_*, and the only mitroa_* objects are those.
    names = catalogue_names(engine)
    legacy_prefixes = tuple(f"{old}_" for old in db.LEGACY_TABLES)
    assert not [n for n in names if n.startswith(legacy_prefixes)]
    assert {
        "mitroa_external_electors_pkey",
        "mitroa_external_electors_characterization_check",
        "mitroa_external_electors_year_elector_idx",
        "mitroa_year_status_pkey",
        "mitroa_year_status_status_check",
        "mitroa_proposals_pkey",
        "mitroa_proposals_action_check",
        "mitroa_proposals_needs_characterization",
        "mitroa_proposals_needs_reasoning",
        "mitroa_proposals_note_not_blank",
        "mitroa_proposals_year_field_idx",
        "mitroa_proposals_year_status_idx",
        "mitroa_proposals_id_seq",
    } <= names
    with engine.connect() as conn:
        owner = conn.execute(
            text("SELECT pg_get_serial_sequence('mitroa_proposals', 'id')")
        ).scalar()
    assert owner == "public.mitroa_proposals_id_seq"

    # The rows came through, and the module reads them under the new names.
    assert db.stored_years() == [2025]
    assert len(db.load_external_electors(2025)) == 3
    assert db.year_state(2026)["baseline_year"] == 2025
    proposals = db.list_proposals(2026)
    assert sorted(proposals["id"]) == [1, 2, 3]

    working = db.working_electors(2026)
    assert set(zip(working["field_code"], working["elector_id"])) == {
        (555, 1), (555, 3), (556, 1)
    }
    projected = db.working_electors(2026, include_pending=True)
    row = projected[(projected["field_code"] == 556) & (projected["elector_id"] == 1)]
    assert row["characterization"].iat[0] == "ΣΥΝΑΦΟΥΣ"


def test_writes_after_the_rename(database, capsys):
    install_legacy(database)
    db.get_engine()
    capsys.readouterr()

    # The sequence kept its position: the next id follows the three old rows.
    assert db.add_proposal(
        2026, 556, 2, db.ADD, "νέος", "y@ihu.gr", "ΙΔΙΟΥ", "στ"
    ) == "Η πρόταση καταχωρήθηκε."
    assert sorted(db.list_proposals(2026)["id"]) == [1, 2, 3, 4]

    # A rejected write names the renamed constraint, so a future traceback
    # points at the right object.
    message = db.add_proposal(2026, 556, 2, "ΛΑΘΟΣ", "x", "y@ihu.gr", "ΙΔΙΟΥ", "z")
    assert "απορρίφθηκε" in message
    assert "mitroa_proposals_action_check" in capsys.readouterr().out

    assert db.withdraw_proposal(4, "y@ihu.gr") == "Η πρόταση αποσύρθηκε."
    assert db.decide_proposal(3, db.ACCEPTED, "c@ihu.gr").endswith(
        db.ACCEPTED.lower() + "."
    )
    assert db.pending_by_field(2026).empty

    # Finalisation writes into the renamed table and locks the renamed year.
    message = db.finalize_year(2026, "c@ihu.gr", blocked_ids={3})
    assert "οριστικοποιήθηκε με 2 εγγραφές" in message
    assert db.stored_years() == [2026, 2025]
    assert db.year_state(2026)["status"] == db.LOCKED
    stored = db.load_external_electors(2026).set_index(["field_code", "elector_id"])
    assert stored.loc[(556, 1), "characterization"] == "ΣΥΝΑΦΟΥΣ"


def test_second_start_is_a_no_op(database, capsys):
    install_legacy(database)
    db.get_engine()
    before = db.mitroa_tables()
    capsys.readouterr()

    db.get_engine.clear()
    db.get_engine()
    out = capsys.readouterr().out
    assert "Αποτυχία" not in out
    assert "Μετονομάστηκαν" not in out
    assert db.mitroa_tables() == before


def test_backup_archive_holds_every_table(database):
    install_legacy(database)
    db.get_engine()

    filename, payload = db.backup_archive()
    assert re.fullmatch(r"mitroa_db_\d{8}-\d{4}\.zip", filename)
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        assert set(archive.namelist()) == {f"{t}.csv" for t in db.mitroa_tables()}
        live = archive.read("mitroa_external_electors.csv")
        assert live.startswith("﻿".encode("utf-8"))
        assert len(pd.read_csv(BytesIO(live))) == 3
        copy = pd.read_csv(BytesIO(archive.read(
            f"{db.BACKUP_PREFIX}proposals.csv"
        )))
        assert list(copy["id"]) == [1, 2, 3]


def test_fresh_installation_uses_the_new_names(database, capsys):
    engine = db.get_engine()
    out = capsys.readouterr().out
    assert "Αποτυχία" not in out
    assert "Μετονομάστηκαν" not in out
    assert db.mitroa_tables() == sorted(db.LEGACY_TABLES.values())

    # The seeder writes into the renamed table, and skips it the second time.
    assert "φορτώθηκαν: 2025 (1476 εγγραφές)" in seed_external.seed_historical_years(
        engine
    )
    assert "υπήρχαν ήδη: 2025" in seed_external.seed_historical_years(engine)
    assert db.stored_years() == [2025]

    # The full production start path, including the other two schemas, and
    # the log line that is the only way to confirm a change landed on Railway.
    status = db.bootstrap()
    assert "Σφάλμα" not in status
    tables = status.split("πίνακες: ")[1].split(", ")
    assert set(db.LEGACY_TABLES.values()) <= set(tables)
    assert not (set(db.LEGACY_TABLES) & set(tables))
