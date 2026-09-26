"""The μητρώα tables against a real, throwaway Postgres.

The production database is reachable only from inside Railway, so until now a
schema change could be checked only by pushing and reading the deployment log.
``pixeltable_pgserver`` (a dev extra) bundles Postgres binaries and starts one in a temp
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

pgserver = pytest.importorskip("pixeltable_pgserver")

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "streamlit"))

from mitroa import data as mdb  # noqa: E402
from mitroa import seed as seed_external  # noqa: E402
from shared import database as db  # noqa: E402
from shared.settings import get_secret  # noqa: E402

# What the tables looked like before the rename is, by construction, the
# current schema without the prefix — that equivalence is what the rename
# asserts, so the test builds the legacy installation from it.
LEGACY_SCHEMA_SQL = mdb.SCHEMA_SQL.replace(mdb.TABLE_PREFIX, "")
LEGACY_MIGRATIONS_SQL = mdb.MIGRATIONS_SQL.replace(mdb.TABLE_PREFIX, "")
LEGACY_ROWS = {
    "external_electors": [
        (2025, 555, 1, "ΙΔΙΟΥ", "α"),
        (2025, 555, 2, "ΣΥΝΑΦΟΥΣ", "β"),
        (2025, 556, 1, "ΙΔΙΟΥ", "γ"),
    ],
    "proposals": [
        # (field, elector, action, characterization, reasoning, status)
        (555, 3, mdb.ADD, "ΙΔΙΟΥ", "δ", mdb.ACCEPTED),
        (555, 2, mdb.REMOVE, None, None, mdb.ACCEPTED),
        (556, 1, mdb.MODIFY, "ΣΥΝΑΦΟΥΣ", "ε", mdb.PENDING),
    ],
}


@pytest.fixture(scope="module")
def server(tmp_path_factory):
    instance = pgserver.get_server(tmp_path_factory.mktemp("pgdata"))
    yield instance
    instance.cleanup()


@pytest.fixture(autouse=True)
def no_committed_archives(tmp_path, monkeypatch):
    """Keep the repository's real backup zips out of the tests' way.

    A committed archive makes the app drop the in-database copies on start;
    the tests decide themselves when that should happen.
    """
    monkeypatch.setattr(mdb, "BACKUPS_DIR", tmp_path / "db_backups")
    (tmp_path / "db_backups").mkdir()


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
            {"open": mdb.OPEN},
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
                {"f": f, "e": e, "a": a, "c": c, "r": r, "s": s, "pending": mdb.PENDING}
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
    for old, new in mdb.LEGACY_TABLES.items():
        assert f"{old} → {new}" in out

    # The tables moved, and a verbatim copy of each was left behind.
    backups = {mdb.BACKUP_PREFIX + old for old in mdb.LEGACY_TABLES}
    assert set(mdb.mitroa_tables()) == set(mdb.LEGACY_TABLES.values()) | backups
    with engine.connect() as conn:
        for old, new in mdb.LEGACY_TABLES.items():
            assert conn.execute(text(f"SELECT to_regclass('{old}')")).scalar() is None
            live = conn.execute(text(f"SELECT count(*) FROM {new}")).scalar()
            copy = conn.execute(
                text(f"SELECT count(*) FROM {mdb.BACKUP_PREFIX}{old}")
            ).scalar()
            assert live == copy

    # Nothing in the catalogue still carries an old name: the μητρώα objects
    # are all mitroa_*, and the only mitroa_* objects are those.
    names = catalogue_names(engine)
    legacy_prefixes = tuple(f"{old}_" for old in mdb.LEGACY_TABLES)
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
    assert mdb.stored_years() == [2025]
    assert len(mdb.load_external_electors(2025)) == 3
    assert mdb.year_state(2026)["baseline_year"] == 2025
    proposals = mdb.list_proposals(2026)
    assert sorted(proposals["id"]) == [1, 2, 3]

    working = mdb.working_electors(2026)
    assert set(zip(working["field_code"], working["elector_id"])) == {
        (555, 1), (555, 3), (556, 1)
    }
    projected = mdb.working_electors(2026, include_pending=True)
    row = projected[(projected["field_code"] == 556) & (projected["elector_id"] == 1)]
    assert row["characterization"].iat[0] == "ΣΥΝΑΦΟΥΣ"


def test_writes_after_the_rename(database, capsys):
    install_legacy(database)
    db.get_engine()
    capsys.readouterr()

    # The sequence kept its position: the next id follows the three old rows.
    assert mdb.add_proposal(
        2026, 556, 2, mdb.ADD, "νέος", "y@ihu.gr", "ΙΔΙΟΥ", "στ"
    ) == "Η πρόταση καταχωρήθηκε."
    assert sorted(mdb.list_proposals(2026)["id"]) == [1, 2, 3, 4]

    # A rejected write names the renamed constraint, so a future traceback
    # points at the right object.
    message = mdb.add_proposal(2026, 556, 2, "ΛΑΘΟΣ", "x", "y@ihu.gr", "ΙΔΙΟΥ", "z")
    assert "απορρίφθηκε" in message
    assert "mitroa_proposals_action_check" in capsys.readouterr().out

    assert mdb.withdraw_proposal(4, "y@ihu.gr") == "Η πρόταση αποσύρθηκε."
    assert mdb.decide_proposal(3, mdb.ACCEPTED, "c@ihu.gr").endswith(
        mdb.ACCEPTED.lower() + "."
    )
    assert mdb.pending_by_field(2026).empty

    # Finalisation writes into the renamed table and locks the renamed year.
    message = mdb.finalize_year(2026, "c@ihu.gr", blocked_ids={3})
    assert "οριστικοποιήθηκε με 2 εγγραφές" in message
    assert mdb.stored_years() == [2026, 2025]
    assert mdb.year_state(2026)["status"] == mdb.LOCKED
    stored = mdb.load_external_electors(2026).set_index(["field_code", "elector_id"])
    assert stored.loc[(556, 1), "characterization"] == "ΣΥΝΑΦΟΥΣ"


def test_second_start_is_a_no_op(database, capsys):
    install_legacy(database)
    db.get_engine()
    before = mdb.mitroa_tables()
    capsys.readouterr()

    db.get_engine.clear()
    db.get_engine()
    out = capsys.readouterr().out
    assert "Αποτυχία" not in out
    assert "Μετονομάστηκαν" not in out
    assert mdb.mitroa_tables() == before


def test_copies_are_dropped_once_their_archive_is_committed(database, capsys):
    install_legacy(database)
    db.get_engine()
    copies = [name for name in mdb.mitroa_tables() if name.startswith(mdb.BACKUP_PREFIX)]
    assert len(copies) == 3

    # An archive from before the copies were taken does not cover them.
    (mdb.BACKUPS_DIR / "mitroa_db_20260101-0900.zip").write_bytes(b"")
    db.get_engine.clear()
    db.get_engine()
    assert [n for n in mdb.mitroa_tables() if n.startswith(mdb.BACKUP_PREFIX)] == copies

    # One from that day or later does, and the copies go — with a log line.
    (mdb.BACKUPS_DIR / "mitroa_db_20260915-1343.zip").write_bytes(b"")
    db.get_engine.clear()
    capsys.readouterr()
    db.get_engine()
    out = capsys.readouterr().out
    assert "Διαγράφηκαν αντίγραφα" in out
    assert all(name in out for name in copies)
    assert "Αποτυχία" not in out
    assert mdb.mitroa_tables() == sorted(mdb.LEGACY_TABLES.values())
    # The live tables were not touched.
    assert mdb.stored_years() == [2025]
    assert sorted(mdb.list_proposals(2026)["id"]) == [1, 2, 3]


def test_backup_archive_holds_every_table(database):
    install_legacy(database)
    db.get_engine()

    filename, payload = mdb.backup_archive()
    assert re.fullmatch(r"mitroa_db_\d{8}-\d{4}\.zip", filename)
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        assert set(archive.namelist()) == {f"{t}.csv" for t in mdb.mitroa_tables()}
        live = archive.read("mitroa_external_electors.csv")
        assert live.startswith("﻿".encode("utf-8"))
        assert len(pd.read_csv(BytesIO(live))) == 3
        copy = pd.read_csv(BytesIO(archive.read(
            f"{mdb.BACKUP_PREFIX}proposals.csv"
        )))
        assert list(copy["id"]) == [1, 2, 3]


def test_fresh_installation_uses_the_new_names(database, capsys):
    engine = db.get_engine()
    out = capsys.readouterr().out
    assert "Αποτυχία" not in out
    assert "Μετονομάστηκαν" not in out
    assert mdb.mitroa_tables() == sorted(mdb.LEGACY_TABLES.values())

    # The seeder writes into the renamed table, and skips it the second time.
    assert "φορτώθηκαν: 2025 (1476 εγγραφές)" in seed_external.seed_historical_years(
        engine
    )
    assert "υπήρχαν ήδη: 2025" in seed_external.seed_historical_years(engine)
    assert mdb.stored_years() == [2025]

    # The full production start path, including the other two schemas, and
    # the log line that is the only way to confirm a change landed on Railway.
    status = db.bootstrap()
    assert "Σφάλμα" not in status
    tables = status.split("πίνακες: ")[1].split(", ")
    assert set(mdb.LEGACY_TABLES.values()) <= set(tables)
    assert not (set(mdb.LEGACY_TABLES) & set(tables))
