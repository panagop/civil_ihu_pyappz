# civil-ihu-pyappz

Streamlit multi-page application for the Civil Engineering department at the International Hellenic University (IHU). Manages course syllabi, course registries (μητρώα), exam scheduling, and weekly timetables.

## Running the app

```bash
uv run streamlit run streamlit/home.py
```

## Project structure

```
civil_ihu_pyappz/
├── streamlit/                        # Streamlit app (entry point + pages)
│   ├── home.py                       # Entry point — st.navigation only (see "Navigation")
│   ├── shared/                       # Used by every feature
│   │   ├── database.py               # THE Postgres engine: get_engine (installs every schema) + bootstrap (seeds)
│   │   ├── shared/settings.py               # get_secret / require_secret — secrets.toml OR env vars
│   │   ├── shared/auth.py                   # OIDC gate helpers + is_coordinator (the coordinator_emails role)
│   │   └── shared/branding.py               # st.logo + department/university names, used by every page
│   ├── mitroa/                       # Μητρώα (page 5)
│   │   ├── CLAUDE.md                 # Feature notes — each package has one (see "Feature notes")
│   │   ├── data.py                   # Schema, migrations, years, proposals, backups (imported as mdb)
│   │   ├── table.py                  # THE submitted table's layout + registry join
│   │   ├── report.py                 # Consolidated Word report (landscape A4)
│   │   ├── ui.py                     # "Προετοιμασία <έτους>" tab — the only writing UI
│   │   └── seed.py                   # Loads external_<year>.xlsx into mitroa_external_electors
│   ├── perigrammata/                 # Περιγράμματα (page 6)
│   │   ├── data.py                   # Column spec, schema, load/save/history (imported as pdb)
│   │   ├── report.py                 # Word output (one / all / changes)
│   │   └── seed.py                   # Loads files/perigrammata/*.csv into the DB (once)
│   ├── eudoxus/                      # Εύδοξος (page 7)
│   │   ├── client.py                 # Unofficial client for service.eudoxus.gr
│   │   ├── data.py                   # Schema, year copy/lock, catalogue (imported as edb)
│   │   └── seed.py                   # Loads files/eudoxus/* into the DB (once)
│   ├── timetable/                    # Εβδομαδιαίο πρόγραμμα (page 8)
│   │   ├── data.py                   # Schema, terms, staff, rooms, classes, conflicts, εκδοχές (imported as tdb)
│   │   └── seed.py                   # Loads files/timetables/{staff,rooms}.csv + 2025-2026.xlsm (once)
│   ├── utils/                        # Workbook readers + Word export for pages 3, 4 and 8
│   ├── app_pages/                    # NOT "pages/" — see "Navigation" below
│   │   ├── 0_home.py                 # Landing page + Microsoft login/logout UI
│   │   ├── 1_📇_perigrammata (legacy).py  # Syllabi v1 (Google Sheets) — superseded by page 6
│   │   ├── 3_⛱_exams-schedule.py    # Exam schedule (public) — reads files/exams/*.xlsm
│   │   ├── 4_📅_weekly_timetable (legacy).py  # Timetable v1 (Excel) — superseded by page 8
│   │   ├── 5_📊_mitroa_v2.py         # Registries v2 (5 tabs) — login gate ACTIVE
│   │   ├── 6_📇_perigrammata_v2.py   # Syllabi v2 (Postgres, editable) — login gate ACTIVE
│   │   ├── 7_📚_eudoxus.py           # Εύδοξος book lists (Postgres) — login gate ACTIVE
│   │   └── 8_📅_weekly_timetable.py  # THE weekly timetable (Postgres) — public view, coordinator edits
│   └── .streamlit/
│       ├── config.toml               # Theme — IS committed (see "Branding and theme")
│       └── secrets.toml              # Google Sheets IDs + auth credentials (NOT in git — create locally)
├── scripts/
│   └── write_secrets_toml.py         # Writes [auth] to secrets.toml at container start
├── civil_ihu_pyappz/                 # Python package (legacy; perigrammata.py not used by the app)
├── files/
│   ├── logos/                        # Department + university marks (committed, not hot-linked)
│   ├── exams/                        # Exam Excel files (.xlsm); active: exams-2026-06.xlsm
│   ├── timetables/                   # Timetable Excel files (.xlsm); active: 2025-2026.xlsm
│   │   ├── staff.csv                 # Seed: the people (site 2026-09-15 + former ΔΕΠ + «ΔΕΠ»)
│   │   └── rooms.csv                 # Seed: the rooms, one code each
│   ├── perigrammata/                 # Frozen Google Sheets export — seed input, then archive
│   │   ├── perigrammata_gr_2018.csv  # 102 courses (+1 debris row without a code)
│   │   ├── perigrammata_gr_2025.csv  # 96 courses
│   │   └── perigrammata_eng_2018.csv # captured, not seeded yet
│   ├── eudoxus/                      # Εύδοξος — seed input, then archive
│   │   ├── eudoxus_books_2025-26.xlsx        # the department's export: 291 rows
│   │   ├── Συγγράμματα ΕΥΔΟΞΟΣ 2026-2027.csv # the 2026-27 export: 305 rows (seeded ΑΝΟΙΧΤΟ)
│   │   ├── eudoxus_catalogue_20260909.csv    # what Eudoxus said about those 222 books
│   │   └── eudoxus.py                        # the original standalone script
│   └── mitroa/                       # Registries — see streamlit/mitroa/CLAUDE.md
│       ├── db_backups/           # CSV zips downloaded from the app (see streamlit/mitroa/CLAUDE.md)
│       ├── professors_tables/        # Annual ΑΠΕΛΛΑ exports (.parquet/.feather/.xlsx)
│       ├── mitroa_by_year/           # Submitted external-elector workbooks (external_<year>.xlsx)
│       ├── antikeimena.csv           # The 52 γνωστικά αντικείμενα (Code, field, domain)
│       ├── registry_snapshots.csv    # year -> which professors_export_*.parquet to join
│       └── json2024/, json2025/      # Older JSON exports
├── jupyter/                          # Exploration notebooks (not part of app)
├── plans/                            # Implementation plans (markdown)
├── tests/                            # pytest; test_db_rename.py runs an embedded Postgres
└── pyproject.toml                    # Dependencies — managed with uv
```

### Code layout (feature packages, since 2026-09-27)

The modules beside `home.py` are grouped **by feature**, not by kind — work
happens one feature at a time. `plans/streamlit_restructure.md` has the
reasoning; the git tag `pre-restructure` is the last flat version.

- **`streamlit/` is the import root**, not the repository: `streamlit run`
  puts the entry script's folder on `sys.path`, and the pages and tests insert
  it. So imports are `from timetable import data`, never
  `from streamlit.timetable …` — the folder named `streamlit` would shadow the
  library.
- **Call sites keep short aliases**: `from shared import database as db`,
  `from mitroa import data as mdb`, `from perigrammata import data as pdb`,
  `from eudoxus import data as edb`, `from timetable import data as tdb`. So
  `db.get_engine()` is the engine everywhere, and anything μητρώα is `mdb.`.
- **A module inside a package reaches `files/` with
  `Path(__file__).resolve().parents[2]`** — one level deeper than before.
- **`monkeypatch` targets the module a name lives in now**, not where it used
  to: `mdb.BACKUPS_DIR`, `auth.is_coordinator`. Patching a name that does not
  exist raises; patching one that exists in the wrong module silently does
  nothing.
- Log prefixes such as `[db.get_engine]` / `[db.bootstrap]` kept their old
  text on purpose: it is what gets searched for in the Railway log.

## Dependencies

Uses `uv` as the package manager.

```bash
uv sync           # install all dependencies
uv sync --extra dev   # include dev tools (pytest, ruff, black, jupyter, pixeltable-pgserver)
```

Key libraries: `streamlit[auth]` (>=1.42 for OIDC), `httpx` (transitive auth dep), `pandas`, `openpyxl`, `python-docx`, `docxtpl`, `docxcompose`, `streamlit-calendar`, `pydantic`, `sqlalchemy` + `psycopg[binary]` (Postgres), `requests` (the Eudoxus client — it was always pulled in by Streamlit, now declared).

- **Always sync with `--extra dev`.** A plain `uv sync` *removes* the dev
  tools, including the embedded Postgres the database tests need — and those
  tests then skip rather than fail.
- **The VS Code Ruff extension locks `.venv\Scripts\ruff.exe`** while its
  server runs, so a sync that replaces ruff fails with "Access is denied"
  (2026-09-26). Set `"ruff.importStrategy": "useBundled"` in the VS Code
  settings, or disable the extension for the sync. The Python Environments
  extension may also pop "Error refreshing packages" while `.venv` is being
  rebuilt; it is transient — the log shows it recovering once the sync ends.
- **Upgrading dependencies**: `uv lock --upgrade --dry-run` first to see what
  moves, then `uv lock --upgrade`, `uv sync --extra dev`, and the full test
  run — it must say **0 skipped**. Push the upgrade as its own commit, apart
  from code changes, so a regression points at one or the other. Last done
  2026-09-26 (SQLAlchemy 2.0 → 2.1 and Streamlit 1.63 → 1.64 needed no code
  change).

## Secrets / credentials

Settings are read through [streamlit/shared/settings.py](streamlit/shared/settings.py), **never
`st.secrets` directly** — see "Settings lookup" below for why.

`streamlit/.streamlit/secrets.toml` is gitignored — but `config.toml` beside it
is **not**, deliberately; see "Branding and theme". On a new machine, create the
secrets file manually with:

```toml
gsheet_perigrammata_id = "..."
gsheet_exams_schedule_id = "..."

# Optional: restrict access to specific @ihu.gr emails. If omitted or empty,
# ANY @ihu.gr account is allowed. See "Authentication" section below.
# allowed_emails = ["someone@ihu.gr", "another@ihu.gr"]

# Microsoft Entra ID OIDC — required for the gated pages (1, 5, 6, 7).
# Restricted to @ihu.gr accounts by streamlit/shared/auth.py.
[auth]
redirect_uri = "http://localhost:8501/oauth2callback"
cookie_secret = "<generate: python -c \"import secrets; print(secrets.token_hex(32))\">"
client_id = "<Application (client) ID from Azure>"
client_secret = "<client secret value from Azure>"
server_metadata_url = "https://login.microsoftonline.com/<TENANT_ID>/v2.0/.well-known/openid-configuration"
```

The Google Sheets are accessed as public CSV exports (no OAuth needed, just the sheet IDs).

### Settings lookup (works locally, on Streamlit Cloud AND on Railway)

`st.secrets` reads **only TOML files** — there is no environment-variable
fallback, so plain env vars on a host like Railway are invisible to it. Worse,
when no secrets file exists at all it **raises `StreamlitSecretNotFoundError`**
rather than reporting a missing key: `st.secrets.get(key, default)` and
`"key" in st.secrets` raise too, so neither can be used to probe safely.

[streamlit/shared/settings.py](streamlit/shared/settings.py) papers over this. It tries
`st.secrets` first (so behaviour is unchanged wherever a file exists) and falls
back to `os.environ`:

| Environment | Source |
|-------------|--------|
| Local | `streamlit/.streamlit/secrets.toml` |
| Streamlit Cloud | Cloud dashboard secrets |
| Railway | service environment variables (same key names) |

- `get_secret(key, default=None)` — value or default.
- `get_secret_list(key)` — list; a comma-separated string (env) is split.
- `require_secret(key)` — value, or `st.error` + `st.stop()` with a clear message.

Use these in new pages. Nested TOML (like `[auth]`) has **no env-var
equivalent**, and `st.login()` reads `[auth]` out of the TOML itself rather than
through `shared/settings.py` — so on a host without a secrets file, OIDC needs a real
file. [scripts/write_secrets_toml.py](scripts/write_secrets_toml.py) writes one
from flat `AUTH_*` variables and is chained into the Railway start command. It
is a no-op without those variables and never overwrites an existing file.

Streamlit searches three locations, last wins: `~/.streamlit/secrets.toml`,
`<cwd>/.streamlit/secrets.toml`, `<entry script dir>/.streamlit/secrets.toml`.
The third is why the file lives in `streamlit/.streamlit/` here, and is where
the script writes.

## Authentication

> **Active since 2026-09-06.** Microsoft login works in all three environments.
> Gated: pages 1, 5, 6 and 7 (`require_ihu_login()`). Public: pages 3, 4 (legacy) and 8 —
> timetables and exam schedules carry no personal data.
>
> `st.login()` raises `StreamlitAuthError` where `[auth]` is missing, so never
> call it unguarded — `auth.is_configured()` exists for that, and
> `render_login_block()` shows a warning instead of a traceback.

The gated pages sit behind Microsoft Entra ID OIDC via Streamlit's native `st.login()`. The gate lives in [streamlit/shared/auth.py](streamlit/shared/auth.py):

- `render_login_block()` — called from `home.py`. Shows the login button when logged out; user info + logout when logged in.
- `require_ihu_login()` — called at the top of each protected page (after `st.set_page_config(...)`). Checks state and `st.stop()`s if unauthorized.

By default, **any** `@ihu.gr` account is accepted. To restrict to a specific set of people, add a top-level `allowed_emails = [...]` list in `secrets.toml` (see commented example above). Behavior:

- `allowed_emails` set → only those exact emails (case-insensitive) work, even if `@ihu.gr`.
- `allowed_emails` missing/empty → any `@ihu.gr` account works.

Non-`@ihu.gr` emails are **always** rejected, listed or not — the suffix check
runs before the allowlist. Granting an external collaborator access would need a
change in `_email_allowed`.

It is one list covering every gated page; there is no per-page group. Set it as
a TOML list locally and on Streamlit Cloud, and as a comma-separated
`allowed_emails` environment variable on Railway (`get_secret_list` splits it).

### Azure / Streamlit Cloud setup gotchas

These cost hours during initial setup (2026-05-13); read before debugging auth issues:

- **`[auth]` block is flat** (no `[auth.microsoft]` subsection) → call `st.login()` with **no argument**. Passing `"microsoft"` errors with "provider not found".
- **`httpx` must be in `pyproject.toml`** — `streamlit[auth]` extra alone doesn't pull it in, but Authlib's starlette_client requires it.
- **Do NOT add `client_kwargs = { prompt = "select_account" }`** — caused `MismatchingStateError` on Streamlit Cloud during testing.
- **Streamlit Cloud secrets are SEPARATE** from the local `secrets.toml`. Edit them independently in the Cloud dashboard (Settings → Secrets). Only `redirect_uri` should differ between them.
- Azure App Registration must have **both** redirect URIs registered under the **Web** platform: `http://localhost:8501/oauth2callback` and `https://<app>.streamlit.app/oauth2callback`. Use the plain `/oauth2callback` path — **not** `/~/+/oauth2callback` (that workaround is for Auth0, not Microsoft).

## Deployment

The app runs in **three places at once** from the same `main` branch; all three
must keep working. Both hosts auto-deploy on push.

| Target | URL | Secrets |
|--------|-----|---------|
| Local | `localhost:8501` | `streamlit/.streamlit/secrets.toml` |
| Streamlit Cloud | `https://<app>.streamlit.app` | Cloud dashboard |
| Railway | `https://civil-ihu.up.railway.app` | service env vars |

Railway (workspace "Georgios Panagopoulos's Projects", Hobby plan):

- Project `civil-ihu-pyappz` — `3e0a1fa7-993f-4f3b-8f8e-9b67d35f0858`
- Service `civil-ihu-pyappz` — `90a05a29-32ff-4d6a-9e8a-3ff28073fcdd`
- Builder Railpack; start command
  `streamlit run streamlit/home.py --server.port $PORT --server.address 0.0.0.0`
- Variables set: `gsheet_perigrammata_id`, `gsheet_exams_schedule_id` (flat
  env vars — resolved via `shared/settings.py`; `gsheet_mitroa_id` is still set but no
  longer read — page 2 was deleted on 2026-09-15 and nothing else opens that
  sheet, so it can be removed at any time),
  `DATABASE_URL` (reference to the Postgres service), and for OIDC
  `AUTH_CLIENT_ID`, `AUTH_CLIENT_SECRET`, `AUTH_COOKIE_SECRET`,
  `AUTH_SERVER_METADATA_URL` (+ optional `AUTH_REDIRECT_URI`, otherwise derived
  from `RAILWAY_PUBLIC_DOMAIN`)
- Start command runs `scripts/write_secrets_toml.py` first, then Streamlit
- The container filesystem is **ephemeral**: generated files do not survive a
  restart. Anything to keep must be downloaded and committed to the repo.
- **A failed build does not take the site down**: Railway keeps serving the
  previous deployment. «Railpack failed to prepare the build» a couple of
  seconds after scheduling, with no build log, is a builder-side hiccup
  (2026-09-26) — redeploy the same deployment before looking for a cause.
  Nothing the build reads (`pyproject.toml`, `uv.lock`, `runtime.txt`) had
  changed.
- **Rolling back** is one click on a previous deployment in Railway (or the
  MCP `list-deployments`, whose entries say `canRollback`). Code-only changes
  can always be rolled back this way; a change that migrated the schema cannot
  be undone by redeploying old code alone.
- **Risky refactors go on a branch** — only `main` deploys — with an
  annotated tag on the last good commit first (`pre-restructure` marks the
  flat-module layout before 2026-09-27).

If OIDC is re-enabled on a host, its `<host>/oauth2callback` must be added to
the Azure App Registration (Web platform) *and* that host's `redirect_uri` must
match — all three environments can be registered simultaneously.

## Database (Postgres on Railway)

Managed Postgres service `Postgres` — `c8e64954-3bab-4655-9d38-a32e4a12d44c`,
in the same project, with a volume. **Deliberately has no public TCP proxy**:
it is reachable only from inside Railway, over the private network. The app
service reads it through `DATABASE_URL = ${{Postgres.DATABASE_URL}}`.

It holds four unrelated groups of tables — `mitroa_*`, `perigrammata_*`,
`eudoxus_*` and `timetable_*` — each described in its package's CLAUDE.md
(see "Feature notes" below). All are installed and
seeded from [streamlit/shared/database.py](streamlit/shared/database.py), which is the only place a
connection is made.

Consequences to keep in mind:

- **You cannot connect from a developer machine or from Streamlit Cloud.**
  There is no `DATABASE_PUBLIC_URL`. Anything requiring the database must run
  inside Railway.
- Locally and on Streamlit Cloud there is no `DATABASE_URL`, so
  `db.get_engine()` returns `None`. Every DB-backed feature **must degrade, not
  crash** — the file-backed tabs of page 5 keep working in all three
  environments. Page 6 degrades by *stopping* with a message: its file source
  is a pre-handover archive, so falling back to it would serve stale
  περιγράμματα rather than none.
- Because nothing outside Railway can seed it, the schema and the historical
  data are installed **by the app itself**, in two places on purpose:
  - **Schema and migrations** run inside `db.get_engine()`. **Not** in
    `bootstrap()`: Streamlit executes only the page you actually open, so a
    visitor landing straight on page 5 never runs `home.py`. A migration that
    depends on the landing page is a migration that silently does not happen —
    this cost a `CheckViolation` in production on 2026-09-05.
  - **Seeding** the historical years stays in `db.bootstrap()`, called from
    both `home.py` and page 5. Slow only the first time; afterwards it is one
    cheap query per year file.

  Both are `@st.cache_resource`, so each runs once per process, and both are
  idempotent. `[db.bootstrap] …` on stdout reports the seed status and the list
  of tables — with no SSH and no public proxy, the deployment log is the only
  way to confirm a schema change landed.
- `SCHEMA_SQL` uses `CREATE TABLE IF NOT EXISTS`, which will **not** alter a
  table that already exists. Anything changing an existing table goes in
  `MIGRATIONS_SQL`, written to be safe on every start (`DROP CONSTRAINT IF
  EXISTS` then `ADD CONSTRAINT`).

## Feature notes (one CLAUDE.md per package)

Each feature's tables, workflows and domain rules live beside its code and load
only when files there are read:

- μητρώα (page 5) — [streamlit/mitroa/CLAUDE.md](streamlit/mitroa/CLAUDE.md)
- περιγράμματα (page 6) — [streamlit/perigrammata/CLAUDE.md](streamlit/perigrammata/CLAUDE.md)
- Εύδοξος (page 7) — [streamlit/eudoxus/CLAUDE.md](streamlit/eudoxus/CLAUDE.md)
- εβδομαδιαίο πρόγραμμα (page 8) — [streamlit/timetable/CLAUDE.md](streamlit/timetable/CLAUDE.md)

Rules that cross features, repeated here so a session in another package sees them:

- `mdb.MIGRATIONS_SQL` runs **before** the timetable `SCHEMA_SQL`, so anything
  touching `timetable_*` tables goes at the end of the timetable schema, not
  there — it would fail on a fresh database.
- The timetable seed must run **after** the περιγράμματα seed: class names are
  read from `perigrammata_courses` by code.
- `coordinator_emails` is one list: it decides μητρώα proposals and locks years
  (page 5) **and** edits the timetable (page 8).

## Branding and theme

Two colours, both sampled from the logo files rather than picked by eye:

| | hex | where |
|---|---|---|
| Department indigo | `#393184` | `files/logos/civil_ihu_logo.png` |
| University navy | `#1C3C61` | `files/logos/ihu_logo.png` |

- **The theme lives in [streamlit/.streamlit/config.toml](streamlit/.streamlit/config.toml)**, which
  Streamlit reads as a *script-level* config because it sits in the entry
  script's directory — the same rule that puts `secrets.toml` there.
- **`.gitignore` exempts it on purpose.** The rest of `.streamlit/` is ignored,
  and git **cannot re-include a file whose parent directory is excluded**, so
  the rule had to become `**/.streamlit/*` plus `!**/.streamlit/config.toml`.
  Without that exemption the theme works on every developer machine and
  silently never reaches Railway — the app deploys unthemed and nothing says so.
  `secrets.toml` stays ignored by its own rule.
- **`primaryColor` differs between light and dark on purpose.** Streamlit paints
  primary-button text white, so the colour has to stay dark enough to carry it:
  `#393184` gives 10.8:1 in light mode but only 1.75:1 against the dark
  background, so `[theme.dark]` lifts it to `#6F66C9` (4.75:1 under white text,
  4.0:1 against the background). Links need to work as body text, hence the
  lighter `#9B93E8` in dark mode.

[streamlit/shared/branding.py](streamlit/shared/branding.py) holds `apply_branding()`, which
calls `st.logo`. **`st.logo` applies to the page it is called from, not to the
app**, so every page script calls it — that is why it is a helper and not one
line in `home.py`. The module deliberately imports nothing but Streamlit,
because pages 3, 4 and 8 are public and import no other shared module.

**The logos are committed under `files/logos/`, not hot-linked.** The
department's URLs are CMS-generated (`/wp-content/uploads/2026/02/…`) and will
move when the site is reorganised, a remote fetch costs a round trip on every
page load, and hot-linking would send every visitor's IP to the department's
server. Both files together are 20 KB.

### Streamlit conventions

The `developing-with-streamlit` skill (installed globally) ships **inside the
installed Streamlit package**, version-matched:
`.venv/Lib/site-packages/streamlit/.agents/skills/developing-with-streamlit/`
(`references/*.md`). For behaviour the docs only hint at, read Streamlit's own
source in the venv — that is how the `pages/` folder rule under "Navigation"
was established.

Current Streamlit guidance, worth following in new code:

- **`use_container_width` is deprecated** — use `width="stretch"` (or
  `"content"`). All 21 live occurrences were converted.
- Prefer Material Symbols (`:material/name:`) over emoji, `st.container(border=True)`
  for grouping, and sentence casing for headings and labels.
- Set the theme in `config.toml` rather than injecting CSS: native theming
  applies to every element and survives upgrades, while CSS selectors target
  internal class names that change.
- `pyproject.toml` selects `E402` for ruff. The pages must call
  `st.set_page_config()` before importing anything that touches Streamlit, so
  imports legitimately follow code; selecting the rule keeps the `# noqa: E402`
  comments that document this meaningful. Since ruff 0.16 they are otherwise
  reported as unused directives on every page — 33 of them.
### Navigation (st.navigation, since 2026-09-15)

[streamlit/home.py](streamlit/home.py) is **only** a router: it declares every
page with `st.Page` and calls `st.navigation(...).run()`. The sidebar therefore
carries real Greek titles instead of `5_📊_mitroa_v2`-style filenames.

- **The folder is `app_pages/`, and renaming it back would silently disable all
  of this.** Streamlit still runs its legacy multipage machinery whenever a
  `pages/` directory sits beside the entry script (`_mpa_v1` in
  `runtime/scriptrunner/script_runner.py`, keyed on
  `PagesManager.uses_pages_directory`). The switch is the folder's mere
  *existence* (`Path(...).exists()` in `runtime/pages_manager.py`), not its
  contents: with pages in it, a URL naming one of them runs that script and
  never reaches `home.py`'s `st.navigation`; even an empty `pages/` makes the
  first run of each process go through `_mpa_v1` with `home.py` as its only
  page, until `st.navigation` switches the mode off. A stray empty `pages/`
  holding only an old `__pycache__` was deleted on 2026-09-27 — delete one on
  sight.
- **The page filenames keep their number and emoji, and must.** `st.Page`
  derives a page's URL from the filename with the same function the old folder
  used (`source_util.page_icon_and_name`, which strips the leading number and
  the emoji), so `/mitroa_v2`, `/exams-schedule` and the rest still resolve and
  old links keep working. The numbers no longer order anything — the list in
  `home.py` does — but renaming a file changes its URL.
- **That rule was used deliberately on 2026-09-18**: page 8 was renamed
  `8_🗓_timetable_v2.py` → `8_📅_weekly_timetable.py`, which hands it
  `/weekly_timetable` — the URL the workbook-backed page had — so links people
  already hold now reach the database version. Page 4 became
  `4_📅_weekly_timetable (legacy).py` and moved to `/weekly_timetable_(legacy)`.
  A rename is the only way to move a URL; there is no redirect.
- **Superseded pages are listed in `LEGACY_PAGES`, not `PAGES`**, and reach
  `st.navigation` only when the sidebar's «Παλαιές σελίδες» checkbox is ticked
  — a page absent from the list never runs. The checkbox lives in the router
  because the list is built before any page runs, and it is cheap enough to be
  worth what every page pays for it. Unticking it while a legacy page is open
  is safe: Streamlit answers an unknown path with the `default=True` page, so
  the user lands on «Αρχική».
- Titles are passed explicitly and repeat what each page sets in its own
  `st.set_page_config(page_title=…)`, so the sidebar label and the browser tab
  agree. Icons are passed explicitly too and are the same emoji the filenames
  carry.
- **The router runs on every rerun, before the selected page.** Anything put
  there is paid for by every page, which is why the landing page is an ordinary
  page (`app_pages/0_home.py`) rather than the body of `home.py`, and why
  `db.bootstrap()` stayed with it. The pages that need the database still call
  `bootstrap()` themselves.
- `st.set_page_config` accepts **repeated, additive calls**: the router sets the
  app-wide title, favicon and sidebar state, and a page's own call overrides
  what it names. The router deliberately sets no `layout`, so the pages that
  want a wide one keep deciding for themselves.

## Active data files

Update these paths inside the page files when switching academic year:

| Page | Active file |
|------|-------------|
| Exam schedule | `files/exams/exams-2026-06.xlsm` |
| Timetable | `files/timetables/2025-2026.xlsm` |

Pages 3 and 4 hardcode their file (page 8 reads the database); page 5 discovers files by glob, so a new
yearly export appears in its dropdowns with no code change.

## Known improvement backlog

These are planned refactors (no functionality changes):

1. **Remove legacy `civil_ihu_pyappz/perigrammata.py`** — dead code, duplicates page 1.
2. **Delete commented-out dead code** in exams-schedule.py.
3. **Replace magic strings with constants** — column names, time slots, file paths.
4. **Add smoke tests** for document generation.
5. **Move active file paths to a config section** so year updates are a single-line change.
6. **Local Postgres for development (agreed 2026-09-09, half done).** The
   tests have one: `pixeltable-pgserver` (dev extra, imported as
   `pixeltable_pgserver`) bundles Postgres binaries and
   [tests/test_db_rename.py](tests/test_db_rename.py) starts a throwaway
   instance in a temp directory, so schema changes are no longer a
   push-and-read-the-log affair — copy its `database` fixture for the next
   one. Running the *app* against a local Postgres is still not set up:
   `settings.get_secret` falls back to environment variables, so any local
   Postgres plus `DATABASE_URL` is enough — `get_engine` installs all three
   schemas and `bootstrap` seeds from the committed files, a full local copy
   with no access to production data.

## Notes

- Python 3.13 required (since 2026-09-26), pinned in three places that must
  move together: `requires-python` in pyproject.toml, `runtime.txt` (which
  Railway's builder reads) and `uv.lock`. Streamlit Cloud ignores all three —
  its version is chosen in the dashboard.
- **The tests use `pixeltable-pgserver`, not `pgserver`**: the original stopped
  at 0.1.4 with no wheels past Python 3.12. The fork imports under another name,
  and the tests load it with `pytest.importorskip` — so a wrong module name does
  not fail, it **silently skips every database test**. After touching it, check
  the run reports no skips.
- **The repository stores LF** (`text=auto`; the "LF will be replaced by
  CRLF" warnings are git's normal noise on Windows). Python's
  `Path.write_text` writes **CRLF** on Windows, so a script that rewrites
  source files should write bytes (`write_bytes(s.encode())`) or pass
  `newline="\n"`, or the diff turns into whole-file churn.
