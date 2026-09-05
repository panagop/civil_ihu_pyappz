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
│   ├── pages/
│   │   ├── 1_📇_perigrammata.py      # Course syllabi — login gate ACTIVE
│   │   ├── 2_📊_mitroa.py            # Course registries — login gate ACTIVE
│   │   ├── 3_⛱_exams-schedule.py    # Exam schedule (public) — reads files/exams/*.xlsm
│   │   ├── 4_📅_weekly_timetable.py  # Weekly timetable (public) — reads files/timetables/*.xlsm
│   │   └── 5_📊_mitroa_v2.py         # Registries v2 (5 tabs) — login gate ACTIVE
│   └── .streamlit/
│       └── secrets.toml              # Google Sheets IDs + auth credentials (NOT in git — create locally)
├── scripts/
│   └── write_secrets_toml.py         # Writes [auth] to secrets.toml at container start
├── civil_ihu_pyappz/                 # Python package (legacy; perigrammata.py not used by the app)
├── files/
│   ├── exams/                        # Exam Excel files (.xlsm); active: exams-2026-06.xlsm
│   ├── timetables/                   # Timetable Excel files (.xlsm); active: 2025-2026.xlsm
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

Key libraries: `streamlit[auth]` (>=1.42 for OIDC), `httpx` (transitive auth dep), `pandas`, `openpyxl`, `python-docx`, `docxtpl`, `streamlit-calendar`, `pydantic`, `sqlalchemy` + `psycopg[binary]` (Postgres).

## Secrets / credentials

Settings are read through [streamlit/settings.py](streamlit/settings.py), **never
`st.secrets` directly** — see "Settings lookup" below for why.

`streamlit/.streamlit/secrets.toml` is gitignored. On a new machine, create it manually with:

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
> Gated: pages 1, 2 and 5 (`require_ihu_login()`). Public: pages 3 and 4 —
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

Consequences to keep in mind:

- **You cannot connect from a developer machine or from Streamlit Cloud.**
  There is no `DATABASE_PUBLIC_URL`. Anything requiring the database must run
  inside Railway.
- Locally and on Streamlit Cloud there is no `DATABASE_URL`, so
  `db.get_engine()` returns `None`. Every DB-backed feature **must degrade, not
  crash** — the file-backed tabs of page 5 keep working in all three
  environments.
- Because nothing outside Railway can seed it, the schema and the historical
  data are installed **by the app itself on first start**:
  `home.py` → `db.bootstrap()` (a `@st.cache_resource`, so once per process).
  Both steps are idempotent. Its status line is printed to stdout as
  `[db.bootstrap] ...` — the deployment logs are the only way to check it.

### `external_electors`

One row per (year, γνωστικό αντικείμενο, elector) — see [streamlit/db.py](streamlit/db.py):

| Column | Notes |
| ------ | ----- |
| `year`, `field_code`, `elector_id` | composite primary key |
| `characterization` | `TEXT` + `CHECK IN ('ΙΔΙΟΥ','ΣΥΝΑΦΟΥΣ')` — not a boolean, so a third category costs no migration |
| `reasoning` | «Αιτιολόγηση συνάφειας» |
| `created_at` | |

Only the *decisions* live here. Name, φορέας, βαθμίδα, ΦΕΚ etc. are joined in
from **that year's** ΑΠΕΛΛΑ export at display time, so they are never
duplicated and a historical table still renders as it was submitted. The
year → export mapping is [files/mitroa/registry_snapshots.csv](files/mitroa/registry_snapshots.csv)
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

## Active data files

Update these paths inside the page files when switching academic year:

| Page | Active file |
|------|-------------|
| Exam schedule | `files/exams/exams-2026-06.xlsm` |
| Timetable | `files/timetables/2025-2026.xlsm` |

Pages 3 and 4 hardcode their file; page 5 discovers files by glob, so a new
yearly export appears in its dropdowns with no code change.

## Page 5 — μητρώα v2

Tabs: **Σύνολο εκλεκτόρων** (browse an annual export) · **Γνωστικά αντικείμενα**
(the 52 subjects) · **Εξωτερικοί εκλέκτορες ανά αντικείμενο** (a submitted
workbook, one subject at a time) · **Έλεγχος εγκυρότητας** (cross-check a
submitted year against a registry export) · **Αναζήτηση με λέξεις-κλειδιά**
(find candidates by γνωστικό αντικείμενο, OR/AND, flag those new since a chosen
year).

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

## Notes

- `streamlit/_ooo_exams-schedule_old.py` is an archived previous version of page 3 — kept for reference, not loaded by Streamlit.
- Python 3.12 required (pinned in pyproject.toml and runtime.txt).
