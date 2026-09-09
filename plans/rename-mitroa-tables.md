# Rename the three μητρώα tables

**Agreed 2026-09-06 · deliberately deferred · not started.**

Referenced from CLAUDE.md, "Database (Postgres on Railway)". Execute this
once the 2026 μητρώο is finalised and locked — members are working in the app
on those tables, and this is not a change to make under them.

## The renames

| now | to |
| --- | -- |
| `external_electors` | `mitroo_electors` |
| `year_status` | `mitroo_years` |
| `proposals` | `mitroo_proposals` |

The common prefix is the point: with no SSH and no public proxy, a future
admin session's `\dt` should show the app's tables as a group.

## Doing it

Cheap in the code — the names appear only in `db.py` and two lines of
`seed_external.py`. The database side has one trap:

- **The renames must run *before* `SCHEMA_SQL`, not in `MIGRATIONS_SQL`.**
  Migrations run after the schema, so with the new names in `SCHEMA_SQL` the
  order would be: create empty `mitroo_electors` → rename `external_electors`
  to `mitroo_electors` → *relation already exists*. `get_engine()` catches and
  prints that, so the app would start against three empty tables while the real
  rows sat orphaned under the old names — indistinguishable from data loss,
  diagnosable only from the deployment log. Add a third block executed first,
  guarded so later starts are no-ops:

  ```sql
  DO $$
  BEGIN
    IF to_regclass('public.external_electors') IS NOT NULL
       AND to_regclass('public.mitroo_electors') IS NULL THEN
      ALTER TABLE external_electors RENAME TO mitroo_electors;
    END IF;
  END $$;
  ```

- Renaming a table renames **neither its indexes, nor its constraints, nor
  `proposals_id_seq`**. Rename them in the same pass: a `CheckViolation` still
  naming `proposals_action_check` is exactly how the 2026-09-05 production bug
  was identified, and a stale name would misdirect that.
- There are no foreign keys between the three tables, so nothing breaks
  structurally.
- **Export the data first.** The rename cannot be undone from outside Railway.
