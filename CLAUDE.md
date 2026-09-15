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
│   ├── auth.py                       # OIDC gate helpers (require_ihu_login, render_login_block)
│   ├── settings.py                   # get_secret / require_secret — secrets.toml OR env vars
│   ├── db.py                         # Postgres engine + schema + bootstrap (Railway only)
│   ├── seed_external.py              # Loads external_<year>.xlsx into mitroa_external_electors
│   ├── external_table.py             # THE submitted table's layout + registry join
│   ├── external_report.py            # Consolidated Word report (landscape A4)
│   ├── proposals_ui.py               # "Προετοιμασία <έτους>" tab — the only writing UI
│   ├── perigrammata_db.py            # Περιγράμματα: column spec, schema, load/save/history
│   ├── seed_perigrammata.py          # Loads files/perigrammata/*.csv into the DB (once)
│   ├── perigrammata_report.py        # Περιγράμματα Word output (one / all / changes)
│   ├── branding.py                   # st.logo + department/university names, used by every page
│   ├── eudoxus_client.py             # Unofficial client for service.eudoxus.gr
│   ├── eudoxus_db.py                 # Εύδοξος: schema, year copy/lock, catalogue
│   ├── seed_eudoxus.py               # Loads files/eudoxus/* into the DB (once)
│   ├── timetable_db.py               # Πρόγραμμα: schema, terms, staff, rooms, classes, conflicts
│   ├── seed_timetable.py             # Loads files/timetables/{staff,rooms}.csv + 2025-2026.xlsm (once)
│   ├── app_pages/                    # NOT "pages/" — see "Navigation" below
│   │   ├── 0_home.py                 # Landing page + Microsoft login/logout UI
│   │   ├── 1_📇_perigrammata (legacy).py  # Syllabi v1 (Google Sheets) — superseded by page 6
│   │   ├── 3_⛱_exams-schedule.py    # Exam schedule (public) — reads files/exams/*.xlsm
│   │   ├── 4_📅_weekly_timetable.py  # Weekly timetable (public) — reads files/timetables/*.xlsm
│   │   ├── 5_📊_mitroa_v2.py         # Registries v2 (5 tabs) — login gate ACTIVE
│   │   ├── 6_📇_perigrammata_v2.py   # Syllabi v2 (Postgres, editable) — login gate ACTIVE
│   │   ├── 7_📚_eudoxus.py           # Εύδοξος book lists (Postgres) — login gate ACTIVE
│   │   └── 8_🗓_timetable_v2.py      # Timetable v2 (Postgres) — public view, coordinator edits
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
│   │   ├── eudoxus_catalogue_20260909.csv    # what Eudoxus said about those 222 books
│   │   └── eudoxus.py                        # the original standalone script
│   └── mitroa/                       # Registries — see "Registry data files" below
│       ├── db_backups/           # CSV zips downloaded from the app (see "Table names")
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

## Dependencies

Uses `uv` as the package manager.

```bash
uv sync           # install all dependencies
uv sync --extra dev   # include dev tools (pytest, ruff, black)
```

Key libraries: `streamlit[auth]` (>=1.42 for OIDC), `httpx` (transitive auth dep), `pandas`, `openpyxl`, `python-docx`, `docxtpl`, `docxcompose`, `streamlit-calendar`, `pydantic`, `sqlalchemy` + `psycopg[binary]` (Postgres), `requests` (the Eudoxus client — it was always pulled in by Streamlit, now declared).

## Secrets / credentials

Settings are read through [streamlit/settings.py](streamlit/settings.py), **never
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
# Restricted to @ihu.gr accounts by streamlit/auth.py.
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

[streamlit/settings.py](streamlit/settings.py) papers over this. It tries
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
through `settings.py` — so on a host without a secrets file, OIDC needs a real
file. [scripts/write_secrets_toml.py](scripts/write_secrets_toml.py) writes one
from flat `AUTH_*` variables and is chained into the Railway start command. It
is a no-op without those variables and never overwrites an existing file.

Streamlit searches three locations, last wins: `~/.streamlit/secrets.toml`,
`<cwd>/.streamlit/secrets.toml`, `<entry script dir>/.streamlit/secrets.toml`.
The third is why the file lives in `streamlit/.streamlit/` here, and is where
the script writes.

## Authentication

> **Active since 2026-09-06.** Microsoft login works in all three environments.
> Gated: pages 1, 5, 6 and 7 (`require_ihu_login()`). Public: pages 3 and 4 —
> timetables and exam schedules carry no personal data.
>
> `st.login()` raises `StreamlitAuthError` where `[auth]` is missing, so never
> call it unguarded — `auth.is_configured()` exists for that, and
> `render_login_block()` shows a warning instead of a traceback.

The gated pages sit behind Microsoft Entra ID OIDC via Streamlit's native `st.login()`. The gate lives in [streamlit/auth.py](streamlit/auth.py):

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
  env vars — resolved via `settings.py`; `gsheet_mitroa_id` is still set but no
  longer read — page 2 was deleted on 2026-09-15 and nothing else opens that
  sheet, so it can be removed at any time),
  `DATABASE_URL` (reference to the Postgres service), and for OIDC
  `AUTH_CLIENT_ID`, `AUTH_CLIENT_SECRET`, `AUTH_COOKIE_SECRET`,
  `AUTH_SERVER_METADATA_URL` (+ optional `AUTH_REDIRECT_URI`, otherwise derived
  from `RAILWAY_PUBLIC_DOMAIN`)
- Start command runs `scripts/write_secrets_toml.py` first, then Streamlit
- The container filesystem is **ephemeral**: generated files do not survive a
  restart. Anything to keep must be downloaded and committed to the repo.

If OIDC is re-enabled on a host, its `<host>/oauth2callback` must be added to
the Azure App Registration (Web platform) *and* that host's `redirect_uri` must
match — all three environments can be registered simultaneously.

## Database (Postgres on Railway)

Managed Postgres service `Postgres` — `c8e64954-3bab-4655-9d38-a32e4a12d44c`,
in the same project, with a volume. **Deliberately has no public TCP proxy**:
it is reachable only from inside Railway, over the private network. The app
service reads it through `DATABASE_URL = ${{Postgres.DATABASE_URL}}`.

It holds four unrelated groups of tables: the `mitroa_*` ones described below, the
`perigrammata_*` ones — see "Περιγράμματα (page 6, Postgres)" — the
`eudoxus_*` ones — see "Εύδοξος (page 7, Postgres)" — and the `timetable_*`
ones — see "Εβδομαδιαίο πρόγραμμα (page 8, Postgres)". All are installed and
seeded from [streamlit/db.py](streamlit/db.py), which is the only place a
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

### Table names (renamed 2026-09-15)

The three μητρώα tables were created as `external_electors`, `year_status`
and `proposals` and renamed to `mitroa_external_electors`,
`mitroa_year_status` and `mitroa_proposals` once the 2026 year was locked, so
that a future admin session's `\dt` shows the app's tables as three groups
(`mitroa_*`, `perigrammata_*`, `eudoxus_*`). `db._rename_legacy_tables` does it
on start, **before** `SCHEMA_SQL` and in the same transaction — run after it,
`CREATE TABLE IF NOT EXISTS` would create three empty tables under the new
names and the app would silently start against them. It also renames the
constraints, indexes and the `proposals_id_seq` sequence, which `ALTER TABLE …
RENAME` leaves alone, and prints `[db.get_engine] Μετονομάστηκαν πίνακες: …`
to the deployment log. Before renaming, each table is copied verbatim to
`mitroa_backup_20260915_<old name>`: nothing outside Railway can take a backup,
so the snapshot lives in the database until it has been downloaded.

**Backups leave the database through the app.** The coordinator section of
the «Προετοιμασία <έτους>» tab has «Αντίγραφο ασφαλείας της βάσης (CSV)»
(`db.backup_archive`): every `mitroa_*` table as a CSV in one zip. Commit the
download under `files/mitroa/db_backups/`. **Committing it is what removes
the copies**: on start, `db._drop_committed_backup_copies` drops every
`mitroa_backup_<date>_*` table for which a `mitroa_db_<date or later>-*.zip`
exists in that folder, and logs it. The archive includes the copies, so a zip
from that day or later holds them; a copy nobody has downloaded yet stays.
The 2026-09-15 snapshot is committed, so the first start after this code
landed drops its copies.

### `mitroa_external_electors`

One row per (year, γνωστικό αντικείμενο, elector) — see [streamlit/db.py](streamlit/db.py):

| Column | Notes |
| ------ | ----- |
| `year`, `field_code`, `elector_id` | composite primary key |
| `characterization` | `TEXT` + `CHECK IN ('ΙΔΙΟΥ','ΣΥΝΑΦΟΥΣ')` — not a boolean, so a third category costs no migration |
| `reasoning` | «Αιτιολόγηση συνάφειας» |
| `created_at` | |

Only the *decisions* live here. Name, φορέας, βαθμίδα, ΦΕΚ etc. are joined in
from an ΑΠΕΛΛΑ export at display time by
[streamlit/external_table.py](streamlit/external_table.py), which owns the join
and the column order so the browse tab, the preview of a year in preparation
and the Word report cannot drift apart. A finalised year joins **its own**
snapshot (the electors as they stood when submitted); a year being prepared
joins the current one. Nothing is duplicated, and a historical table still
renders as it was submitted. The year → export mapping is [files/mitroa/registry_snapshots.csv](files/mitroa/registry_snapshots.csv)
(`db.registry_file_for_year`) — a file, not a table, because it must also
resolve where there is no database, and because the export filenames are date
stamps rather than years. **Never delete an old parquet export**: without it
that year's table cannot be reconstructed.

The `α/α` column is deliberately not stored — it is derived on render. (Only
9 of the 52 sheets in `external_2025.xlsx` reproduce exactly by sorting ΙΔΙΟΥ
first then alphabetically; the rest carry hand-placed rows. That is noise, not
information.)

Seeding lives in [streamlit/seed_external.py](streamlit/seed_external.py),
which parses `external_<year>.xlsx` and skips any year that already has rows.
2025 loads as 1.476 rows / 52 αντικείμενα / 500 distinct electors.

**Direction of travel:** from 2026 on there is to be **no `external_<year>.xlsx`
at all** — the year is built and kept in the database, and the workbooks that
exist stay only as the historical record. Keep the file source working for those
years; do not build new features that require one.

`load_external_from_db` (page 5) rebuilds a stored year in exactly the shape
`load_external_workbook` returns — same keys, same column order — so the tab
renders either source with one code path. Electors missing from that year's
export are kept with blank columns and counted in a warning, never dropped.

### Preparing a new year (proposals workflow)

A year under preparation is **never stored as rows** while it is open. Its table
is computed on read as *baseline year + accepted proposals*
(`db.working_electors`); only `db.finalize_year` writes it into
`mitroa_external_electors`. So `mitroa_external_electors` always means "officially approved",
and the page-5 tab and the Word report need no notion of drafts.

- `mitroa_year_status(year, status, baseline_year, …)` — `ΑΝΟΙΧΤΟ` → `ΚΛΕΙΔΩΜΕΝΟ`.
  `db.open_year(2026, baseline_year=2025, …)` copies nothing; it just records
  the baseline.
- `mitroa_proposals` — one row per proposed change: `ΠΡΟΣΘΗΚΗ` / `ΑΦΑΙΡΕΣΗ` /
  `ΜΕΤΑΒΟΛΗ` (carries the new characterisation *and* the new justification, so
  changing both stays one decision a coordinator cannot half-accept; the older
  split `ΧΑΡΑΚΤΗΡΙΣΜΟΣ` / `ΑΙΤΙΟΛΟΓΗΣΗ` still replay), with a mandatory `note`,
  the `author` (from
  `st.user.email`, so it is asserted by Microsoft rather than typed) and
  `ΕΚΚΡΕΜΕΙ` / `ΕΓΚΡΙΘΗΚΕ` / `ΑΠΟΡΡΙΦΘΗΚΕ` / `ΑΠΟΣΥΡΘΗΚΕ`.

**Why proposals rather than editing the table directly:** several members work
on the same 52 subjects and Streamlit locks nothing. Direct edits would mean
last-write-wins, silently. Two proposals on the same elector simply coexist and
the coordinator resolves them.

`working_electors(year, include_pending=True)` replays the undecided proposals
on top of the accepted ones, projecting what the table would become if
everything proposed were approved. The tab renders it under "Ο πίνακας του
<έτους> μετά τις προτεινόμενες αλλαγές", between the proposal forms and the
coordinator section: removals are gone, additions and changes are listed, and
the person columns come from the current ΑΠΕΛΛΑ export. It is shaped by
`external_table`, so it carries **the submitted layout** — this is the table
that goes to the department, not a review view; a checkbox adds a marker column
and can be cleared for the exact submitted form, and it downloads as Excel.
Accepted proposals are replayed before pending ones — they are already reality,
so a pending change to the same elector should win.

Replay rules (`working_electors`): proposals are applied in decision order, so a
later accepted one wins; `ΧΑΡΑΚΤΗΡΙΣΜΟΣ`/`ΑΙΤΙΟΛΟΓΗΣΗ` aimed at an elector who
has since been removed are **no-ops, not errors**. `finalize_year` refuses while
any proposal is still pending.

The tab lives in [streamlit/proposals_ui.py](streamlit/proposals_ui.py) — the
only part of the app that writes anything. Per-subject it shows the computed
table (🔴 on anyone with a κώλυμα in the current registry, and ➕➖🔄✏️ for
pending proposals), then sub-tabs for Μεταβολή / Προσθήκη / Οι προτάσεις μου.
Below the per-subject part come the year-wide sections — the projected table,
the new electors across all subjects, the Word report — and a coordinator-only
section to decide proposals and lock the year. The
coordinator can decide the current subject's pending proposals **in bulk**
(`db.decide_field_proposals` — one UPDATE, so a subject is decided whole or not
at all; a half-applied batch is hard to reason about when one elector has
several proposals), behind a confirmation checkbox, with
`db.pending_by_field` showing where the remaining work is. Actions
sit in a form under the table rather than as buttons on each row: ~26 electors
× 3 buttons would rebuild ~80 widgets per rerun for a worse layout.

🟡 marks anyone whose βαθμίδα, γνωστικό αντικείμενο or φορέας moved since the
baseline year, with the before → after in its own column
(`registry_changes`). It compares through `fold_greek_series`, **not**
`casefold`: the exports re-typed subjects in title case, and `ΔΥΝΑΜΙΚΗ` vs
`Δυναμική` differs by an accent casefold keeps — that alone was two false
findings out of 42.

The mandatory «Αιτιολόγηση της μεταβολής» offers ready-made reasons
(`REMOVAL_SUGGESTIONS`, `ADDITION_SUGGESTIONS`, and one built from the actual
characterisation change) through `_suggestion_pick`, which only *fills* the box
— the text stays editable and "Άλλο" clears it. The picker and the
characterisation radio both sit outside the form for the reason below, and the
note's key includes the chosen suggestion so a new choice replaces the default.

**The elector selectbox is outside `st.form` on purpose.** A widget inside a
form does not rerun until submit, so the fields below kept showing the previous
elector's justification. Widget keys also include the elector id, because
Streamlit keeps the stored value of a key that has not changed and would
override the new defaults.

**Losing eligibility in the ΑΠΕΛΛΑ registry removes an elector automatically.**
It is not a judgement anyone makes, so it is not a proposal: `working_electors`
takes `blocked_ids` and drops them at the source, so no view and no write can
forget. `finalize_year` passes the same set, so a locked year cannot contain
one. The tab keeps **two views** of the year: the filtered one drives the preview,
the report and finalisation, while the overview table renders the *unfiltered*
one so the removed electors stay visible in place with 🔴 (a caption says they
are not in the totals and do not reach the submitted table). Seeing who dropped
out, in position, is what makes the table readable at a glance — hiding them
was a regression on 2026-09-06.
`db.auto_removals` reconstructs who left, since there is no proposal row to
look at, and the report prints them with `AUTO_REMOVAL_NOTE` ("Διαγραφή λόγω μη
επιλεξιμότητας στο μητρώο του ΑΠΕΛΛΑ") and status `ΑΥΤΟΜΑΤΗ`. For 2026 that is
58 rows / 34 people across 32 subjects. Members still propose *other* removals
themselves, with their own justification.

The report's change table has **three sources**, because each leaves a different
trace: proposals (a row in `mitroa_proposals`), automatic removals (no row at all), and
**registry updates** — βαθμίδα, φορέας, γνωστικό αντικείμενο moving between the
baseline year's export and the current one (`registry_changes`, action
`ΕΝΗΜΕΡΩΣΗ`). The last are not changes to the list, but they *are* changes to
what the list prints, so a reader must see the values were updated rather than
wonder why they differ from last year. Entries are ordered
ΑΦΑΙΡΕΣΗ → ΠΡΟΣΘΗΚΗ → ΜΕΤΑΒΟΛΗ → ΕΝΗΜΕΡΩΣΗ within a subject.

A `ΜΕΤΑΒΟΛΗ` that leaves the characterisation where it was is **left out of the
report**: rewording a justification changes the text printed next to an elector,
not their standing in the μητρώο, and listing it buries the changes that matter.
It is compared against what held *before* — the baseline year plus any addition
already replayed — never against the proposal itself, which carries the new
value. The legacy `ΑΙΤΙΟΛΟΓΗΣΗ` action is reasoning-only by definition and is
dropped the same way. The preview and the tab still show these; only the
submitted document omits them. For 2026: 58
removals and 132 registry updates, so every one of the 52 subjects has a table.

Roles are two, and there is deliberately **no users table**: a coordinator is an
email listed in the `coordinator_emails` setting (the same list also governs
who edits the timetable on page 8) (same mechanism as
`allowed_emails`), everyone else who passes the login gate is a member. Members
may propose on any subject; only a coordinator decides proposals and locks a
year. A table would need an admin screen to manage and still need some way to
appoint the first admin.

### Νέοι εκλέκτορες, όλα τα αντικείμενα μαζί (added 2026-09-09)

Between the per-subject preview and the Word report, `_additions_block` lists
**every elector new since the baseline year, across all 52 subjects**, with a
filter for the still-pending ones and an Excel download. The per-subject
preview answers "what does this subject look like now"; this answers "who did
we take in this year", which is the question asked when the whole cycle is
reviewed — and 52 subjects is too many to answer by clicking through them.

- **"New" is a set difference, not a read of the `ΠΡΟΣΘΗΚΗ` proposals.** It
  compares `(field_code, elector_id)` between the projected year and the
  baseline table, so it reports what the year *is*: an addition that was later
  withdrawn or removed again does not appear, nor does one for somebody who has
  since lost eligibility (`working_electors` already filtered them). Reading
  the proposals instead would list decisions, several of which no longer hold.
- The key is the **pair**, so an elector moved to a different γνωστικό
  αντικείμενο is correctly new *there* — the same rule the rest of the μητρώα
  code follows.
- The proposal that caused each addition is looked up separately
  (`_addition_meta`) only for the κατάσταση / author / note columns. A row with
  no proposal behind it still renders, with those columns blank, rather than
  being dropped.
- The download follows the pending filter, so the file matches what is on
  screen.
- `render` computes `working_electors(..., include_pending=True)` and the
  baseline **once** and passes them to both this block and the report; they
  were being recomputed.

### Consolidated report

[streamlit/external_report.py](streamlit/external_report.py) builds one Word
document covering all 52 subjects — a summary table, then a landscape A4 page
each. It takes the same parsed structure either source produces, so the button
works identically for file and database. ~9 s for 1.476 rows, so it sits behind
a button and a spinner rather than being built on load.

`build_report(..., changes=..., draft=...)` puts, **before** each subject's
table, a short table of the additions, removals and changes with the reason for
each — what moved reads first, the list it produced second — and only
who / what / why, since the electors' full details follow underneath. The
`Κατάσταση` column disappears when nothing is `ΕΚΚΡΕΜΕΙ` (it would repeat the
same value on every row), and its width goes to the reason. Rejected and
withdrawn proposals are left out. `draft=True` stamps
"ΠΡΟΧΕΙΡΟ — περιλαμβάνει προτάσεις που δεν έχουν εγκριθεί ακόμη" under the
title, so a report of a year still in preparation cannot be mistaken for the
final one; the "Προετοιμασία <έτους>" tab generates exactly this from the
projected table, and drops the stamp once the year is locked.

**Word, not PDF, on purpose.** `docx2pdf` (already in `pyproject.toml`) shells
out to a real Microsoft Word install and is Windows-only — it cannot run in the
Railway container at all. PDF there would mean WeasyPrint/wkhtmltopdf and system
packages, for a document that gets edited before submission anyway.

Column widths in `COLUMN_WIDTHS` total 27.6 cm against the 27.7 cm usable on
landscape A4 at 1 cm margins. **Word silently ignores every width if the total
overflows the page**, so adding a column means taking the room from another.

**The database view is not byte-identical to the workbook, by design.** Of
13.284 compared cells for 2025, 475 differ: 40 are only capitalisation (the
workbook is hand-typed in caps), 334 are the known Κατηγορία Χρήστη relabelling
(`Ημεδαπής` → `Καθηγητής Ημεδαπής` — the workbook predates the export it is
joined to), and ~100 are real drift in Βαθμίδα, ΦΕΚ and Σχολή. The database view
shows the **official registry values**; the file view shows what was typed.

## Περιγράμματα (page 6, Postgres)

Page 1 reads the Google Sheet and **keeps working** — leave it alone until
explicitly told to remove it. Page 6 is its replacement and never touches the
sheet.

The sheet was exported once, on 2026-09-09, into `files/perigrammata/` and the
database is the master from that point. That export is the *only* thing that
reads those CSVs — after the first successful seed they are the historical
record of what the sheet held at handover, exactly as `external_<year>.xlsx`
is for the μητρώα. **Page 6 does not fall back to them**: with no
`DATABASE_URL` it stops with a message, because a page quietly serving
pre-handover data is worse than a page that says it cannot run. That does mean
page 6 works **only on Railway** until a local Postgres exists (see the
backlog).

All three worksheets (`gr`, `gr_2025`, `eng`) carry identical column names, so
it is one table with `(curriculum, locale, code)` as the key. `eng` matched the
2018 curriculum exactly (102 of its 103 codes), so English is a *locale*, not a
separate programme.

### `perigrammata_courses`

One row per course. The 39 content columns are generated into the DDL from
`CONTENT_COLUMNS` in [streamlit/perigrammata_db.py](streamlit/perigrammata_db.py),
so the table, the edit form, the Word context and the seeder cannot drift apart
— add a column there and everything follows. `FIELD_GROUPS` in the same file
carries the Greek label and the widget kind for each, so a new column cannot be
added without deciding how it is edited.

- **`locale`, not `lang`.** `lang` is already a *content* column: the language
  the course is taught in («Ελληνική»). Two different things, and calling both
  `lang` would have been a bug waiting to happen.
- **Numerics are stored as numbers** (`examino` INTEGER, hours/ects NUMERIC) so
  εξάμηνο can be charted and the workload summed. Every value in both years is
  a whole number, and `perigrammata_report.format_value` prints them as such —
  page 1 rendered `4.0` into Word where the sheet said `4`.
- **`sort_order`** is the sheet's old `id`. It is row order within a worksheet,
  not an identity; it survives only so the full report prints in the order
  people are used to.
- **2018 is read-only** (`EDITABLE_CURRICULUM = 2025`), enforced in
  `save_course`, not just hidden in the UI.
- `refs_more_title` / `refs_more` exist in every worksheet but in **neither Word
  template**. They are stored and editable and simply do not print; the Word tab
  says so when a course has them. Dropping them on import would have lost data.

### `perigrammata_revisions`

Append-only, one row per save: `data` is the full new state, `changes` is only
the fields that moved. Both, deliberately — `data` makes a course recoverable,
`changes` makes the coordinator's report a read rather than a diff of
consecutive snapshots.

**No proposals workflow here, unlike μητρώα.** That machinery exists because 52
subjects are worked on by many members with no owner and Streamlit locks
nothing. A περίγραμμα has one natural owner per course, so anyone who passes
the login gate edits directly and the history carries the accountability. The
one concurrency guard is optimistic: the form remembers the `updated_at` it
rendered from and the UPDATE matches on it, so a second tab that loaded earlier
fails loudly instead of silently overwriting.

The seed writes an origin revision per course with an **empty `changes`
object**, so the initial import is recoverable but never shows up as 198
courses "changed" on the day the database was filled.

### Word output

[streamlit/perigrammata_report.py](streamlit/perigrammata_report.py), three
documents:

- `render_course` — one περίγραμμα, from the same `docxtpl` template as page 1
  but read from `files/` instead of fetched from GitHub on every click.
- `build_full_report` — every course, one per page, merged with `docxcompose`.
  ~10 s for the 96 courses of 2025, so it sits behind a button, a spinner, and a
  cache keyed on the newest `updated_at`. Each course is rendered separately and
  the results merged because the template is a whole-page form and `docxtpl`
  cannot repeat one.
- `build_changes_report` — what moved in a period, grouped by course, read
  straight out of `perigrammata_revisions`. Not derivable from the courses
  table, which only ever holds the present. Values are truncated to 400
  characters: `subject1` alone runs to 4.583, and printed in full one edit would
  bury every other change.

Column widths total 27.6 cm against the 27.7 cm usable on landscape A4 — the
same trap as the μητρώα report, where **Word ignores every width if the total
overflows**.

### English, deferred

`SEED_LOCALES = ("gr",)` in
[streamlit/seed_perigrammata.py](streamlit/seed_perigrammata.py). The English
sheet is exported and committed but not loaded: it writes εξάμηνο as `1st` /
`2nd`, which needs a mapping before it can enter an INTEGER column. Adding
`"eng"` there is the easy half.

## Εύδοξος (page 7, Postgres)

The books offered per academic year, and the list for the year being prepared.
A year is named by the one it starts in: `2025` is «2025-26» (`year_label`).
Database-only, for the same reason as page 6.

Tables: `eudoxus_years` · `eudoxus_courses` · `eudoxus_selections` ·
`eudoxus_books` · `eudoxus_changes`, all in
[streamlit/eudoxus_db.py](streamlit/eudoxus_db.py).

### A new year is a copy, not a replay

The μητρώα tables compute an open year as *baseline + accepted proposals* and
store nothing until it is finalised. That machinery exists because 52 subjects
are worked on by many members with no owner. A book list has one teacher per
course, so `open_year` **copies** the baseline outright and people edit their
own courses directly, exactly as they edit περιγράμματα. What the coordinator
reviews is `changes_vs_baseline` — a FULL OUTER JOIN of the two years, so it
cannot go stale and shows the *net* change; `eudoxus_changes` is the separate
audit log of who did each edit, including the ones that cancelled out. One year
is open at a time, and `open_year` refuses a second.

### The εξάμηνο is part of the course key

**ΔΟΜ022 «Οικοδομική ΙΙ» is listed twice** in the 2025-26 export, once as 7th
εξάμηνο and once as 9th — the old programme running beside the new one — with
the same three books in a *different* priority order (0/1/2 vs 0/2/1). So
`eudoxus_courses` is keyed on `(year, course_code, examino)` and
`eudoxus_selections` on `(year, course_code, examino, book_id)`. Keying on the
course code alone silently loses one of the two orderings.

The list also spans **both** curricula: of its 101 courses, 83 are in the 2025
περιγράμματα and the other 18 only in the 2018 ones. Do not assume a book list
maps onto one πρόγραμμα σπουδών.

### `eudoxus_books` — the catalogue cache

Book metadata is **not** in `eudoxus_selections`: a book appears in several
courses (222 distinct books over 291 rows) and its availability changes on
Eudoxus' side, not ours. `found` distinguishes "withdrawn from the registry"
from "this code is not in the registry at all" — different problems for whoever
has to fix the list.

- **A book that has never been checked is not reported as a problem.** Absence
  of an answer is not a negative one; the page counts those separately.
- **`active` and `selectable` are different flags and both matter.** Of the 9
  problem books found on 2026-09-09, 4 are `active=True, selectable=False` —
  checking only `active` would have missed them.
- Type coercion in `_clean_book` is not decoration: ISBNs arrive as **integers**
  from both the JSON and pandas, and Postgres has no assignment cast from
  bigint to text, so an uncoerced one fails the INSERT. numpy scalars out of a
  DataFrame are the same class of problem — psycopg cannot adapt them at all.

### The Eudoxus client

[streamlit/eudoxus_client.py](streamlit/eudoxus_client.py) is an **unofficial,
reverse-engineered** client for the JSON endpoint behind the public
[Σύνθετη Αναζήτηση](https://service.eudoxus.gr/search/#/advanced) page. There is
no published contract:

- `PUT /search/rest/app/advanced-search`, and **unknown keys are rejected with
  HTTP 500** — send only fields in `_TEMPLATE`.
- `id` is an exact match on a single code and **there is no bulk endpoint**, so
  checking a year is inherently one request per book: ~1 s each, ~3.5 minutes
  for 222. That is why results are stored rather than fetched per page view,
  and why the check runs behind a button with a progress bar.
- `fetch_books` never raises — a dead code comes back `found=False` with the
  reason, so one bad book cannot abort a check of two hundred.

It grew out of [files/eudoxus/eudoxus.py](files/eudoxus/eudoxus.py), which stays
where it was written as a standalone script.

### Seed data

`files/eudoxus/eudoxus_books_2025-26.xlsx` is the department's own export
(291 rows → 102 course offerings, 291 selections) and
`eudoxus_catalogue_<YYYYMMDD>.csv` is a one-off dump of what Eudoxus said about
each of those 222 books, fetched once so the browse tab shows titles rather
than bare numeric codes from the first run.

This is the **only** import: from the next year on a list is opened as a copy
and edited in the app. The catalogue is loaded **only while the table is
empty** — after that the availability check owns it, and re-applying the dump
on every restart would replace a fresh answer with a stale one. Seeded years are
inserted as `ΚΛΕΙΔΩΜΕΝΟ`.

## Εβδομαδιαίο πρόγραμμα (page 8, Postgres)

Page 4 reads `files/timetables/2025-2026.xlsm` and **stays as it is** until
told otherwise. Page 8 shows the same five views from the database and adds
the preparation of the next semester. Viewing is public like page 4; editing
is for `coordinator_emails` only (decided 2026-09-15), so the page checks the
role where it matters instead of gating itself with `require_ihu_login`.

Tables in [streamlit/timetable_db.py](streamlit/timetable_db.py):
`timetable_staff` · `timetable_staff_terms` · `timetable_rooms` ·
`timetable_terms` · `timetable_classes` · `timetable_class_instructors` ·
`timetable_class_rooms` · `timetable_changes`.

### A term is a copy of last year's same period

A *term* is `(year, period)`: `(2026, 'Χειμερινό')` is the winter of 2026-27.
`open_term` **copies** the same period one year earlier — classes, links, and
who was active — and the coordinator rearranges; `lock_term` freezes it. One
term is open at a time. Same model as Εύδοξος, for the same reason: one owner,
so no proposals machinery.

### Staff, not instructor

The site lists ΕΤΕΠ who never teach, and "instructor" is the role a person has
on one class. `timetable_staff` holds everyone (category ΔΕΠ / ΕΔΙΠ / ΕΤΕΠ /
ΕΚΤΑΚΤΟΣ / ΑΛΛΟ, rank, site URL) and prints `short_name` — «Σαφούρη Χρ.» vs
«Σαφούρη Γ.». Λιαλιαμπής and Παπαϊωάννου are former ΔΕΠ who still teach and
are filed as ΔΕΠ. «ΔΕΠ» is a `placeholder` row for the courses all faculty
teach (ΓΕΝ009, ΓΕΝ010); the conflict check ignores placeholders.

**Activity is per term** (`timetable_staff_terms`), not a column on the
person: a column per semester would need a migration every term, and a from/to
range does not fit έκτακτοι who come and go. No row means *unknown*, and the
seed marks as active only whoever taught in that term. The Προσωπικό tab edits
the flag for the selected term with a `data_editor`.

### Classes

One row per meeting, as a workbook row was: εξάμηνο, course, `section`
(Θ · Ε · Ε1…Ε9 · Φ), day 1–5, whole `start_hour` (08–20, **no half hours**),
`duration`. `day IS NULL` is a course listed for the term and not yet placed —
the 16 workbook rows with only a semester survive that way, and the
Προετοιμασία tab lists them. Instructors and rooms are link tables: both are
already many-to-one in the data (ΓΕΝ002 has two instructors, ΔΟΜ011 two rooms).

**Names come from the περιγράμματα.** The row stores `curriculum` +
`course_code`, and `load_term` joins `perigrammata_courses` — so the seed must
run after the περιγράμματα one. The curriculum is stored rather than guessed
because the department is mid-transition: `NEW_CURRICULUM_EXAMINA` says which
εξάμηνα follow the 2025 programme in each year: 2025-26 the first year only
(εξάμηνα 1–2), 2026-27 the first two (1–4), everything else 2018 — one year of
study moves over per year (corrected 2026-09-15). Extend it when the next
year moves over. `open_term` re-resolves the curriculum of every copied row
by the rule of the *new* year — but only where that programme has the code
**in the same εξάμηνο** (ΔΟΜ007 is 3rd in 2018 and 4th in 2025, so it must not
flip) — so last winter's 3rd-εξάμηνο rows become 2025 rows in 2026-27 where
they can. The rest keep their old curriculum and `off_programme` lists them in
the Προετοιμασία tab for the coordinator to replace; nothing is dropped
silently. `candidate_courses` uses the same rule to list
what the programme offers for a period. ΔΟΜ004 sits in the 2nd εξάμηνο but is
not in the 2025 programme; the seed falls back to 2018 for such codes and says
so in the log.

`name_suffix` keeps the «ΔΥ, ΥΕ» marker the workbook wrote after elective
titles, and `display_name` prints it after the περίγραμμα name. It also drives
the conflict rule below.

`load_term` returns the workbook's column names (`course_id`, `class_name`,
`full_class_name`, `semester`, `instructors`, `day`, `start_time`, `duration`,
`room`…) so `utils/timetable_export.py` works unchanged on either source.

### Conflicts are warnings, not constraints

`conflicts(frame)` is pure and runs on what `load_term` returns. Two placed
classes on the same day with overlapping hours conflict when they share a
room, share a non-placeholder instructor, or belong to the same εξάμηνο —
**except** parallel lab groups of one course (Ε1 / Ε2 run together on
purpose) and electives whose stream markers share no letter (ΓΥ beside
«ΥΥ, ΔΕ» is fine; ΓΥ beside ΓΕ is not; a course without a marker is common to
all and conflicts with anything). Θ and Ε of the same course overlapping *is* a
conflict. The tab shows them as a table; saving is never blocked, because a
term being rearranged is allowed to be inconsistent. With the stream rule the
seeded 2025-26 still reports a handful, all genuinely in the workbook (ΔΟΜ001
Θ and ΔΟΜ002 Ε1 both in 301 on Tuesday 10:00; ΓΕΩ009 «ΓΥ» over ΥΔΡ006
«ΥΥ, ΓΕ»).

### Seed data

[streamlit/seed_timetable.py](streamlit/seed_timetable.py) loads
`files/timetables/staff.csv` and `rooms.csv` while those tables are empty, and
`2025-2026.xlsm` as two locked terms — once. `WORKBOOKS` is an explicit dict,
not a glob: **`2026-2027.xlsm` is a byte-identical copy of the 2025-26 file**
and must not be seeded as 2026-27 (it is kept until the database approach is
confirmed, then deleted). `ROOM_ALIASES` maps every room spelling the workbook
used to a code («Αίθ. Τεχν. Σχεδίου 1» and «Εργαστήριο Τεχνικού Σχεδίου Ι»
are the same room, ΤΣ1); an unmapped spelling **raises**, as does an unknown
instructor. The bare «Σαφούρη» on ΔΟΜ007 is Γεωργία (the drawing lab).

Rooms and staff are edited in the page from then on; the CSVs are the
historical record, like the μητρώα workbooks.

[tests/test_timetable.py](tests/test_timetable.py) covers the parser without
a database, and the seed, copy-open, edits, lock and the page script itself
(via `streamlit.testing.v1.AppTest`) against the embedded Postgres.

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

[streamlit/branding.py](streamlit/branding.py) holds `apply_branding()`, which
calls `st.logo`. **`st.logo` applies to the page it is called from, not to the
app**, so every page script calls it — that is why it is a helper and not one
line in `home.py`. The module deliberately imports nothing but Streamlit,
because pages 3 and 4 are public and import no other shared module.

**The logos are committed under `files/logos/`, not hot-linked.** The
department's URLs are CMS-generated (`/wp-content/uploads/2026/02/…`) and will
move when the site is reorganised, a remote fetch costs a round trip on every
page load, and hot-linking would send every visitor's IP to the department's
server. Both files together are 20 KB.

### Streamlit conventions

Current Streamlit guidance, worth following in new code:

- **`use_container_width` is deprecated** — use `width="stretch"` (or
  `"content"`). All 21 live occurrences were converted; the one left in
  `_ooo_exams-schedule_old.py` is in the archived file Streamlit never loads.
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
  `PagesManager.uses_pages_directory`): it builds navigation from the folder
  itself and never reaches the `st.navigation` call.
- **The page filenames keep their number and emoji, and must.** `st.Page`
  derives a page's URL from the filename with the same function the old folder
  used (`source_util.page_icon_and_name`, which strips the leading number and
  the emoji), so `/mitroa_v2`, `/exams-schedule` and the rest still resolve and
  old links keep working. The numbers no longer order anything — the list in
  `home.py` does — but renaming a file changes its URL.
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

Pages 3 and 4 hardcode their file; page 5 discovers files by glob, so a new
yearly export appears in its dropdowns with no code change.

## Page 5 — μητρώα v2

Tabs: **Εκλέκτορες ΑΠΕΛΛΑ** (browse an annual export) · **Γνωστικά αντικείμενα
ΔΙΠΑΕ** (the 52 subjects) · **Εξωτερικοί εκλέκτορες ΔΙΠΑΕ** (one subject at a
time, from **either** the database **or** the submitted workbook — a radio picks
the source, defaulting to the database, which is where a year lives from 2026
on; the file option is the historical archive, and the database option appears
only where there are stored years) · **Έλεγχος εγκυρότητας** (cross-check a
submitted year against a registry export) · **Αναζήτηση με λέξεις-κλειδιά**
(find candidates by γνωστικό αντικείμενο, OR/AND, flag those new since a chosen
year) · **Προετοιμασία <έτους>** (propose, decide and finalise the year being
prepared — see below).

`WORKING_YEAR` / `BASELINE_YEAR` at the top of the page name the year being
prepared and the finalised one it starts from. Bump both when the next cycle
begins, and add the new registry export to `registry_snapshots.csv`.

Reads only local files — no secrets, no network.

### Registry data files

- `files/mitroa/professors_tables/professors_export_<YYYYMMDD>.parquet` — annual
  ΑΠΕΛΛΑ export. **Schemas differ between years**: 2024/2025 have `Σε αναστολή`
  and `Παρ. 1, Άρ. 145, Ν. 4957/2022`; 2026 replaced these with `Ενεργή
  άδεια/κώλυμα αποκλεισμού από μητρώα` and `... από εκλεκτορικά/επιτροπές`. Code
  must tolerate this — never assume a fixed column set.
- `files/mitroa/mitroa_by_year/external_<year>.xlsx` — the submitted external
  electors. One worksheet per γνωστικό αντικείμενο, **named by code** (555–606);
  the readable title sits at row 6 (0-based `META_ROW`), the table header at row
  9 (`HEADER_ROW`). The `α/α` column holds spreadsheet formulas — renumber it.
- `files/mitroa/antikeimena.csv` — the 52 αντικείμενα (`Code, field, domain`).

Domain rules encoded in page 5:

- **Only `κώλυμα αποκλεισμού από μητρώα` disqualifies** (and legacy `Σε
  αναστολή`). Exclusion from εκλεκτορικά/επιτροπές is reported but does not
  block — see `BLOCKING_FLAG_KEYWORDS`. Flag columns are detected by their
  ΝΑΙ/ΟΧΙ *values*, so renamed columns keep working.
- **Κατηγορία Χρήστη is never compared across years** — the exports relabelled it
  (`Ημεδαπής` → `Καθηγητής Ημεδαπής`), which would produce 179 false findings.
- Χαρακτηρισμός is spelled inconsistently in the workbooks, in **both** the
  values (`ΙΔΙΟ`/`ΙΔΙΟΥ`, `ΣΥΝΑΦΕΣ`/`ΣΥΝΑΦΟΥΣ` — normalised via
  `CHARAKTIRISMOS_ALIASES`) and the column header itself
  (`Χαρακτηρισμός` vs the soft-hyphenated `Χαρακτη-ρισμός`). **Never match a
  workbook header literally** — `load_external_workbook` renames it to the
  canonical `CHARAKTIRISMOS_COL` using `fold_header` (fold_greek with
  non-letters dropped). A literal constant silently matched nothing and left
  the ΙΔΙΟΥ/ΣΥΝΑΦΟΥΣ metrics, filter and comparison column dead until
  2026-09-05. `seed_external.py` matches its headers the same way.

### Greek text matching

Use `fold_greek` / `fold_greek_series`, not `casefold()`. Plain casefolding fails
twice over: `σκυροδέμ` does not match `Σκυρόδεμα` (the accent sits on a different
vowel) and `ς` does not fold to `σ`. The helpers strip combining accents, casefold
and unify final sigma.

The parquet columns are Arrow-backed, so pandas runs their regexes through
**RE2, which rejects `\u` escapes** — build character classes from `chr()` (see
`COMBINING_MARKS_RE`) rather than writing `"[̀-ͯ]"`.

## Known improvement backlog

These are planned refactors (no functionality changes):

1. **Remove legacy `civil_ihu_pyappz/perigrammata.py`** — dead code, duplicates page 1.
2. **Delete commented-out dead code** in exams-schedule.py.
3. **Replace magic strings with constants** — column names, time slots, file paths.
4. **Add smoke tests** for document generation.
5. **Move active file paths to a config section** so year updates are a single-line change.
6. **Local Postgres for development (agreed 2026-09-09, half done).** The
   tests have one: `pgserver` (dev extra) bundles Postgres binaries and
   [tests/test_db_rename.py](tests/test_db_rename.py) starts a throwaway
   instance in a temp directory, so schema changes are no longer a
   push-and-read-the-log affair — copy its `database` fixture for the next
   one. Running the *app* against a local Postgres is still not set up:
   `settings.get_secret` falls back to environment variables, so any local
   Postgres plus `DATABASE_URL` is enough — `get_engine` installs all three
   schemas and `bootstrap` seeds from the committed files, a full local copy
   with no access to production data.

## Notes

- `streamlit/_ooo_exams-schedule_old.py` is an archived previous version of page 3 — kept for reference, not loaded by Streamlit.
- Python 3.12 required (pinned in pyproject.toml and runtime.txt).
