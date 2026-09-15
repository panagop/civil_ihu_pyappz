# Rename the three μητρώα tables

**Agreed 2026-09-06 · executed 2026-09-15**, once the 2026 μητρώο had been
finalised and locked — members were working in the app on those tables, and
this was not a change to make under them.

Referenced from CLAUDE.md, "Database (Postgres on Railway)".

## The renames

The names chosen at execution differ from the ones first pencilled in
(`mitroo_electors` / `mitroo_years` / `mitroo_proposals`): the existing names
were kept and only prefixed, so a reader of old commits, logs or plans still
recognises them.

| before | after |
| ------ | ----- |
| `external_electors` | `mitroa_external_electors` |
| `year_status` | `mitroa_year_status` |
| `proposals` | `mitroa_proposals` |

The common prefix is the point: with no SSH and no public proxy, a future
admin session's `\dt` shows the app's tables as three groups — `mitroa_*`,
`perigrammata_*`, `eudoxus_*`.

## How it was done

The names appear only in `db.py` and two statements of `seed_external.py`;
those were changed outright. The database side is `db._rename_legacy_tables`,
called from `get_engine()` **before** `SCHEMA_SQL`, in the same transaction:

- For each table still present under its old name (and whose new name is
  free) it first copies the table verbatim to `mitroa_backup_20260915_<old>`
  — a snapshot of the exact pre-rename state, kept *inside* the database
  because nothing outside Railway can take one — then renames the table, and
  then its constraints, indexes and the owned sequence (`ALTER TABLE … RENAME`
  touches none of those). It prints
  `[db.get_engine] Μετονομάστηκαν πίνακες: …` so the deployment log shows it
  happened. A later start finds nothing to do.
- The order is the trap: run after the schema, `CREATE TABLE IF NOT EXISTS
  mitroa_…` would have created three empty tables, the rename would have
  failed on "relation already exists", and the app would have started against
  empty tables while the real rows sat orphaned under the old names.
- There are no foreign keys between the three tables, so nothing broke
  structurally.

Verified before deploying by `tests/test_db_rename.py` against an embedded
Postgres (`pgserver`, in the dev extras): a database built under the old
names is renamed on start with every row, constraint name and the sequence
position intact, every write path still works, a second start is a no-op,
and a fresh database comes up under the new names with no backup copies.

## Follow-up

1. In the app, as coordinator, open «Προετοιμασία 2026» → «Αντίγραφο
   ασφαλείας της βάσης (CSV)», download the zip and commit it under
   `files/mitroa/db_backups/`. It contains the three live tables and the
   three `mitroa_backup_20260915_*` copies.
2. Once that commit exists, the copies can be dropped. There is no UI for
   it on purpose (a delete button next to a backup button invites the wrong
   click); add a one-off `DROP TABLE` to `MIGRATIONS_SQL` guarded with
   `IF EXISTS`, deploy, and remove it again.
