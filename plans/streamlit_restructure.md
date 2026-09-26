# Restructure `streamlit/` into feature packages

Status: **done on branch `restructure-streamlit`** (2026-09-27). Nothing changes for users.

## Why

`streamlit/` holds 17 flat modules (~7.000 lines) beside `home.py`. They
already group themselves by name prefix — `timetable_*`, `eudoxus_*`,
`perigrammata_*`, `seed_*` — and one file, `db.py` (929 lines), mixes the
shared engine with the whole μητρώα domain. Work happens one feature at a
time, so the folder should be organised the same way.

The bundled Streamlit skill (`developing-with-streamlit`, Streamlit 1.64)
agrees with the direction: pages in `app_pages/` as direct scripts, business
logic in importable modules. It prescribes no layout for those modules beyond
a `utils/` example; grouping by feature is our choice.

## Target layout

```text
streamlit/
├── home.py                     unchanged (router)
├── app_pages/                  unchanged — filenames are URLs
├── utils/                      unchanged (workbook readers + Word export)
├── .streamlit/                 unchanged
├── shared/
│   ├── __init__.py
│   ├── database.py             ← db.py: get_engine, bootstrap, is_available
│   ├── settings.py             ← settings.py
│   ├── auth.py                 ← auth.py  (+ is_coordinator, from db.py)
│   └── branding.py             ← branding.py
├── mitroa/
│   ├── __init__.py
│   ├── data.py                 ← db.py: schema, migrations, rename/backups,
│   │                              years, proposals, working_electors …
│   ├── table.py                ← external_table.py
│   ├── report.py               ← external_report.py
│   ├── ui.py                   ← proposals_ui.py
│   └── seed.py                 ← seed_external.py
├── perigrammata/
│   ├── data.py  report.py  seed.py      ← perigrammata_db / _report / seed_perigrammata
├── eudoxus/
│   ├── data.py  client.py  seed.py      ← eudoxus_db / eudoxus_client / seed_eudoxus
└── timetable/
    ├── data.py  seed.py                 ← timetable_db / seed_timetable
```

`_ooo_exams-schedule_old.py` is deleted (it is in git history; Streamlit never
loads it).

### Naming decisions

- **Feature packages at top level, no umbrella package.** `streamlit/` is
  already on `sys.path` (`streamlit run` adds the entry script's folder, the
  pages and tests insert it), so `from timetable import data` works with no
  extra setup. Checked: none of `shared`, `mitroa`, `perigrammata`, `eudoxus`,
  `timetable` is an installed distribution in the venv. `shared` rather than
  `core`/`common`, which are generic enough to collide with a future
  dependency.
- **Modules named by role (`data`, `seed`, `report`)** — the package already
  says which feature.
- **Call sites keep their existing aliases**, which is what keeps this
  refactor mechanical:

  | today | after |
  | --- | --- |
  | `import db` (engine) | `from shared import database as db` |
  | `import timetable_db as tdb` | `from timetable import data as tdb` |
  | `import perigrammata_db as pdb` | `from perigrammata import data as pdb` |
  | `import eudoxus_db as edb` | `from eudoxus import data as edb` |
  | `db.working_electors(...)` etc. | `mdb.working_electors(...)` — `from mitroa import data as mdb` |

  So `db.get_engine()` (81 call sites) and every `tdb.` / `pdb.` / `edb.` call
  stay byte-identical; only the μητρώα calls on page 5, `proposals_ui` and the
  tests change prefix from `db.` to `mdb.`.

## The `db.py` split — the only non-mechanical part

`shared/database.py` keeps: `get_engine`, `is_available`, `bootstrap`.

`mitroa/data.py` takes everything μητρώα: `CHARACTERIZATIONS`, the status and
action constants, `SCHEMA_SQL`, `MIGRATIONS_SQL`, `LEGACY_TABLES`,
`TABLE_PREFIX`, `BACKUP_PREFIX`, `BACKUPS_DIR`, `PROFESSORS_DIR`,
`SNAPSHOTS_CSV`, `_rename_legacy_tables`, `_drop_committed_backup_copies`,
`mitroa_tables`, `backup_archive`, `load_external_electors`,
`registry_file_for_year`, `stored_years`, and the year/proposal functions.

`is_coordinator` moves to `shared/auth.py`: it is a role check used by page 5
**and** page 8, not μητρώα logic.

**Schema order must survive the move exactly.** `get_engine` runs, in one
transaction:

1. `_drop_committed_backup_copies` (μητρώα)
2. `_rename_legacy_tables` (μητρώα) — must precede step 3, or
   `CREATE TABLE IF NOT EXISTS` creates empty tables under the new names
3. μητρώα `SCHEMA_SQL`, then `MIGRATIONS_SQL`
4. περιγράμματα, Εύδοξος, timetable `SCHEMA_SQL` — in that order (the
   timetable references `perigrammata_courses`)

`get_engine` keeps importing the feature modules **inside the function**, as
today: every feature module imports `shared.database` for the engine, so a
top-level import back would be circular. A registry/hook mechanism would remove
the cycle but is a design change; not in scope.

The log prefixes `[db.get_engine]` / `[db.bootstrap]` **stay as they are** —
the Railway log is the only way to confirm a schema change, and the text
people search for should not change.

## Things that will break unless handled

- **`ROOT = Path(__file__).resolve().parents[1]`** in `db.py`,
  `branding.py`, `perigrammata_report.py` and the four seeds becomes
  `parents[2]` one level deeper. Every one is a path into `files/`; a wrong
  one fails the seed or the logo at runtime, and the tests catch the seeds and
  the report. `branding.py` is only exercised by AppTest pages — check the logo
  explicitly.
- **`monkeypatch` targets follow the name, not the file:**
  - `test_db_rename.py`: `monkeypatch.setattr(db, "BACKUPS_DIR", …)` →
    `mdb`, since `_drop_committed_backup_copies` reads the global of the module
    it lives in. Patching the old name would silently patch nothing and the
    test would read the real `files/mitroa/db_backups/`.
  - `test_timetable.py`: `monkeypatch.setattr(db, "is_coordinator", …)` →
    `auth`, where page 8 will look it up. A `from shared.auth import
    is_coordinator` in the *page* is fine — AppTest re-executes the page's
    imports on every run, after the patch — but a helper module imported
    earlier (e.g. `mitroa/ui.py`) must call it as `auth.is_coordinator(...)`,
    or it keeps the unpatched function.
- **`importorskip`-style silent skips:** the database tests skip when
  `pixeltable_pgserver` is missing. After the move, confirm the run says
  **77 passed, 0 skipped**; any skip means an import broke.
- **Pages' `sys.path.insert(0, …parent.parent)`** stays — it already points at
  `streamlit/`, which is still the import root.
- **`@st.cache_resource` keys** include the function's module, so the caches
  simply start fresh under the new names. Harmless: they are per process.
- **No URL changes.** `app_pages/` is untouched, and page URLs come from those
  filenames only.

## Steps (one commit, tests between steps locally)

0. CLAUDE.md, "Navigation": correct "never reaches the `st.navigation` call".
   Per `runtime/pages_manager.py:61` the legacy mode is switched on by the
   mere *existence* of `pages/`; with no `.py` files in it, `_mpa_v1` runs
   `home.py` as its only page, which then calls `st.navigation` and switches
   the mode off — it works, by detour. (The stray empty `pages/` folder was
   deleted on 2026-09-27.)
1. `git mv` every module to its new place, adding `__init__.py` files — so
   git keeps each file's history (`git log --follow`).
2. Split `db.py` into `shared/database.py` + `mitroa/data.py`; move
   `is_coordinator` to `shared/auth.py`.
3. Rewrite imports: the modules, the pages, the tests (using the alias table
   above). Fix every `ROOT = …parents[…]`.
4. Fix the two monkeypatch targets.
5. `uv run --extra dev pytest` → 77 passed, 0 skipped.
   `uv run ruff check streamlit tests` → no new findings beyond the existing
   ones.
6. Run the app locally and open every page once (pages 1, 3, 4 need no
   database; 5–8 stop or degrade without one, which is itself the check).
   Check the logo renders.
7. Update CLAUDE.md: the project-structure tree and every
   `streamlit/<module>.py` link (~40 of them).
8. Push; in the Railway deploy log, confirm `[db.bootstrap]` lists all 20
   tables with no `Αποτυχία εφαρμογής σχήματος` line.

## Out of scope

- `utils/` stays as it is. Its timetable and exam modules could move into the
  feature packages later, but page 4 (legacy, due for deletion) uses them; do
  that when page 4 goes.
- Breaking the `database` ↔ feature-schema import cycle.
- Renaming `home.py` to `streamlit_app.py` (the skill's default): it would
  change the Railway start command for no gain.
- The bogus `[tool.hatch.build.targets.wheel] packages = ["myproject"]` in
  `pyproject.toml` — unrelated; worth its own small fix.

## Effort and risk

About an hour. Risk is low and front-loaded: nearly every mistake is an
`ImportError` at test or page load, not a silent wrong result. The two
exceptions — a monkeypatch hitting the wrong module, and a `ROOT` one level
off — are called out above with how to check them.
