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
│   ├── home.py                       # Landing page + Microsoft login/logout UI
│   ├── auth.py                       # OIDC gate helpers (require_ihu_login, render_login_block)
│   ├── settings.py                   # get_secret / require_secret — secrets.toml OR env vars
│   ├── db.py                         # Postgres engine + schema + bootstrap (Railway only)
│   ├── seed_external.py              # Loads external_<year>.xlsx into external_electors
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
│   ├── pages/
│   │   ├── 1_📇_perigrammata (legacy).py  # Syllabi v1 (Google Sheets) — superseded by page 6
│   │   ├── 2_📊_mitroa (legacy).py        # Registries v1 — superseded by page 5
│   │   ├── 3_⛱_exams-schedule.py    # Exam schedule (public) — reads files/exams/*.xlsm
│   │   ├── 4_📅_weekly_timetable.py  # Weekly timetable (public) — reads files/timetables/*.xlsm
│   │   ├── 5_📊_mitroa_v2.py         # Registries v2 (5 tabs) — login gate ACTIVE
│   │   ├── 6_📇_perigrammata_v2.py   # Syllabi v2 (Postgres, editable) — login gate ACTIVE
│   │   └── 7_📚_eudoxus.py           # Εύδοξος book lists (Postgres) — login gate ACTIVE
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
│   ├── perigrammata/                 # Frozen Google Sheets export — seed input, then archive
│   │   ├── perigrammata_gr_2018.csv  # 102 courses (+1 debris row without a code)
│   │   ├── perigrammata_gr_2025.csv  # 96 courses
│   │   └── perigrammata_eng_2018.csv # captured, not seeded yet
│   ├── eudoxus/                      # Εύδοξος — seed input, then archive
│   │   ├── eudoxus_books_2025-26.xlsx        # the department's export: 291 rows
│   │   ├── eudoxus_catalogue_20260909.csv    # what Eudoxus said about those 222 books
│   │   └── eudoxus.py                        # the original standalone script
│   └── mitroa/                       # Registries — see "Registry data files" below
│       ├── professors_tables/        # Annual ΑΠΕΛΛΑ exports (.parquet/.feather/.xlsx)
│       ├── mitroa_by_year/           # Submitted external-elector workbooks (external_<year>.xlsx)
│       ├── antikeimena.csv           # The 52 γνωστικά αντικείμενα (Code, field, domain)
│       ├── registry_snapshots.csv    # year -> which professors_export_*.parquet to join
│       └── json2024/, json2025/      # Older JSON exports
├── jupyter/                          # Exploration notebooks (not part of app)
├── plans/                            # Implementation plans (markdown)
├── tests/                            # Minimal tests (pytest)
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
gsheet_mitroa_id = "..."
gsheet_exams_schedule_id = "..."

# Optional: restrict access to specific @ihu.gr emails. If omitted or empty,
# ANY @ihu.gr account is allowed. See "Authentication" section below.
# allowed_emails = ["someone@ihu.gr", "another@ihu.gr"]

# Microsoft Entra ID OIDC — required for pages 1 (perigrammata) and 2 (mitroa).
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
> Gated: pages 1, 2, 5, 6 and 7 (`require_ihu_login()`). Public: pages 3 and 4 —
> timetables and exam schedules carry no personal data.
>
> `st.login()` raises `StreamlitAuthError` where `[auth]` is missing, so never
> call it unguarded — `auth.is_configured()` exists for that, and
> `render_login_block()` shows a warning instead of a traceback.

Pages 1 (perigrammata) and 2 (mitroa) are gated behind Microsoft Entra ID OIDC via Streamlit's native `st.login()`. The gate lives in [streamlit/auth.py](streamlit/auth.py):

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
- Variables set: `gsheet_perigrammata_id`, `gsheet_mitroa_id`,
  `gsheet_exams_schedule_id` (flat env vars — resolved via `settings.py`),
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

It holds three unrelated groups of tables: the μητρώα ones described below, the
`perigrammata_*` ones — see "Περιγράμματα (page 6, Postgres)" — and the
`eudoxus_*` ones — see "Εύδοξος (page 7, Postgres)". All are installed and
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

### Pending: rename the three tables (agreed 2026-09-06, deliberately deferred)

Agreed but **not yet applied** — members are working in the app on the 2026
tables, and this is not a change to make under them. Do it once 2026 is
finalised:

| now | to |
| --- | -- |
| `external_electors` | `mitroo_electors` |
| `year_status` | `mitroo_years` |
| `proposals` | `mitroo_proposals` |

The common prefix is the point: with no SSH and no public proxy, a future
admin session's `\dt` should show the app's tables as a group.

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

### `external_electors`

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
`external_electors`. So `external_electors` always means "officially approved",
and the page-5 tab and the Word report need no notion of drafts.

- `year_status(year, status, baseline_year, …)` — `ΑΝΟΙΧΤΟ` → `ΚΛΕΙΔΩΜΕΝΟ`.
  `db.open_year(2026, baseline_year=2025, …)` copies nothing; it just records
  the baseline.
- `proposals` — one row per proposed change: `ΠΡΟΣΘΗΚΗ` / `ΑΦΑΙΡΕΣΗ` /
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
pending proposals), then sub-tabs for Μεταβολή / Προσθήκη / Οι προτάσεις μου,
and a coordinator-only section to decide proposals and lock the year. The
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
trace: proposals (a row in `proposals`), automatic removals (no row at all), and
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
email listed in the `coordinator_emails` setting (same mechanism as
`allowed_emails`), everyone else who passes the login gate is a member. Members
may propose on any subject; only a coordinator decides proposals and locks a
year. A table would need an admin screen to manage and still need some way to
appoint the first admin.

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
- Pages still use the legacy `pages/` folder. Streamlit now recommends
  `st.navigation` / `st.Page`, which would give proper Greek titles and icons in
  the sidebar instead of `5_📊_mitroa_v2`-style filenames. Not done: it touches
  every page.

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
6. **Local Postgres for development (agreed 2026-09-09, deferred).** Every
   DB-backed feature — the μητρώα tables *and* the new περιγράμματα ones — can
   only be exercised on Railway, which makes the edit form on page 6 and any
   schema change a push-and-read-the-log affair. `settings.get_secret` already
   falls back to environment variables, so a throwaway Postgres (Docker, or any
   local install) plus `DATABASE_URL` is enough: `get_engine` installs both
   schemas and `bootstrap` seeds from the committed files, giving a full local
   copy with no access to production data. Nothing in the code needs to change
   — this is a dev-setup task, and worth doing before the next schema change.

## Notes

- `streamlit/_ooo_exams-schedule_old.py` is an archived previous version of page 3 — kept for reference, not loaded by Streamlit.
- Python 3.12 required (pinned in pyproject.toml and runtime.txt).
