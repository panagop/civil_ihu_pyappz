"""The Postgres engine: the single place a connection is made.

It also installs every feature's schema (:func:`get_engine`) and seeds the
historical data (:func:`bootstrap`); the tables themselves belong to the
feature packages — :mod:`mitroa.data`, :mod:`perigrammata.data`,
:mod:`eudoxus.data`, :mod:`timetable.data`.

The database lives on Railway and is reachable **only from inside Railway**
(no public TCP proxy — see CLAUDE.md, "Database"). Locally and on Streamlit
Cloud there is no ``DATABASE_URL``, so :func:`get_engine` returns ``None`` and
every caller must degrade gracefully rather than crash: the file-backed tabs of
page 5 keep working everywhere, only the database-backed ones go quiet.

Because nothing outside Railway can reach the database, the schema and the
historical data are installed by the app itself on first start. Both steps are
idempotent.
"""

from __future__ import annotations

from sqlalchemy import Engine, create_engine, text

import streamlit as st
from shared.settings import get_secret


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
    engine = create_engine(url, pool_pre_ping=True)

    # The schema is installed here rather than in bootstrap() because Streamlit
    # runs only the page you actually open: land straight on page 5 and home.py
    # never executes. A migration that depends on the landing page is a
    # migration that silently does not happen. This function is cached, so the
    # DDL runs once per process, and it is cheap and idempotent.
    try:
        # Imported here, not at module level: every feature module imports this
        # one for the engine, so a top-level import would be circular.
        from eudoxus.data import SCHEMA_SQL as EUDOXUS_SCHEMA_SQL
        from mitroa import data as mitroa
        from perigrammata.data import SCHEMA_SQL as PERIGRAMMATA_SCHEMA_SQL
        from timetable.data import SCHEMA_SQL as TIMETABLE_SCHEMA_SQL

        # The order is load-bearing. The μητρώα rename must run before its
        # CREATE TABLE IF NOT EXISTS, or those would create three empty tables
        # under the new names; the timetable references perigrammata_courses.
        with engine.begin() as conn:
            dropped = mitroa._drop_committed_backup_copies(conn)
            renamed = mitroa._rename_legacy_tables(conn)
            conn.execute(text(mitroa.SCHEMA_SQL))
            conn.execute(text(mitroa.MIGRATIONS_SQL))
            conn.execute(text(PERIGRAMMATA_SCHEMA_SQL))
            conn.execute(text(EUDOXUS_SCHEMA_SQL))
            conn.execute(text(TIMETABLE_SCHEMA_SQL))
        if dropped:
            print(
                "[db.get_engine] Διαγράφηκαν αντίγραφα ασφαλείας που υπάρχουν "
                "πλέον στο αποθετήριο: " + ", ".join(dropped),
                flush=True,
            )
        if renamed:
            print(
                "[db.get_engine] Μετονομάστηκαν πίνακες: " + ", ".join(renamed),
                flush=True,
            )
    except Exception as exc:  # noqa: BLE001 - reported, not raised
        print(f"[db.get_engine] Αποτυχία εφαρμογής σχήματος: {exc}", flush=True)
    return engine



def is_available() -> bool:
    return get_engine() is not None


@st.cache_resource
def bootstrap() -> str:
    """Load the historical years. Runs once per process, from home.py.

    The schema itself is installed by :func:`get_engine`; only the data seeding
    lives here, because it is slow and needed by nothing until a page asks for a
    stored year. Never raises: a database problem must not take down the
    file-backed pages.
    """
    engine = get_engine()
    if engine is None:
        return "Χωρίς βάση δεδομένων (δεν έχει οριστεί DATABASE_URL)."
    try:
        from eudoxus.seed import seed_eudoxus
        from mitroa.seed import seed_historical_years
        from perigrammata.seed import seed_perigrammata
        from timetable.seed import seed_timetable

        status = seed_historical_years(engine)
        status = f"{status} · {seed_perigrammata(engine)}"
        status = f"{status} · {seed_eudoxus(engine)}"
        # After the περιγράμματα: the timetable resolves its course codes there.
        status = f"{status} · {seed_timetable(engine)}"
        # Report the tables too: without SSH into the container, and with no
        # public proxy to the database, the log line is the only way to confirm
        # a schema change actually landed.
        with engine.connect() as conn:
            tables = sorted(
                row[0]
                for row in conn.execute(
                    text(
                        "SELECT tablename FROM pg_tables WHERE schemaname = 'public'"
                    )
                )
            )
        status = f"{status} · πίνακες: {', '.join(tables) or '(κανένας)'}"
    except Exception as exc:  # noqa: BLE001 - reported, not raised
        status = f"Σφάλμα βάσης: {exc}"
    # Also to stdout: the database is unreachable from outside Railway, so the
    # deployment logs are the only place this can be checked.
    print(f"[db.bootstrap] {status}", flush=True)
    return status
