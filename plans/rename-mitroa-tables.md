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

## Aftermath

Deployed 2026-09-15 10:41 UTC (commit `a4e7ef7`); the log showed the rename
line and the bootstrap table list with the six `mitroa_*` tables. The
coordinator downloaded `mitroa_db_20260915-1343.zip` (2.937 electors rows,
107 proposals, 1 year — each twice, live and copy) and committed it under
`files/mitroa/db_backups/` (commit `44bb9ef`).

The copies are not dropped by hand — there is deliberately no delete button
next to a backup button. `db._drop_committed_backup_copies` drops a
`mitroa_backup_<date>_*` table on start once a zip dated that day or later is
committed in the folder, and logs it — so the first deploy carrying that
function removes the 2026-09-15 copies. Confirm it in the deployment log:
`[db.get_engine] Διαγράφηκαν αντίγραφα ασφαλείας …` followed by a bootstrap
table list with three `mitroa_*` tables.
