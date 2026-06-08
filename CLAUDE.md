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
│   ├── pages/
│   │   ├── 1_📇_perigrammata.py      # Course syllabi — PROTECTED (requires @ihu.gr login)
│   │   ├── 2_📊_mitroa.py            # Course registries — PROTECTED (requires @ihu.gr login)
│   │   ├── 3_⛱_exams-schedule.py    # Exam schedule (public) — reads files/exams/*.xlsm
│   │   └── 4_📅_weekly_timetable.py  # Weekly timetable (public) — reads files/timetables/*.xlsm
│   └── .streamlit/
│       └── secrets.toml              # Google Sheets IDs + auth credentials (NOT in git — create locally)
├── civil_ihu_pyappz/                 # Python package (legacy; perigrammata.py not used by the app)
├── files/
│   ├── exams/                        # Exam Excel files (.xlsm); active: exams-2026-06.xlsm
│   └── timetables/                   # Timetable Excel files (.xlsm); active: 2025-2026.xlsm
│   └── mitroa/                       # Registry JSON exports (json2024/, json2025/)
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

Key libraries: `streamlit[auth]` (>=1.42 for OIDC), `httpx` (transitive auth dep), `pandas`, `openpyxl`, `python-docx`, `docxtpl`, `streamlit-calendar`, `pydantic`.

## Secrets / credentials

`streamlit/.streamlit/secrets.toml` is gitignored. On a new machine, create it manually with:

```toml
gsheets_id_perigrammata = "..."
gsheets_id_mitroa_eklektores = "..."
gsheets_id_mitroa_antikeimena = "..."
# add any other Sheet IDs used by the pages

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

## Authentication

Pages 1 (perigrammata) and 2 (mitroa) are gated behind Microsoft Entra ID OIDC via Streamlit's native `st.login()`. The gate lives in [streamlit/auth.py](streamlit/auth.py):

- `render_login_block()` — called from `home.py`. Shows the login button when logged out; user info + logout when logged in.
- `require_ihu_login()` — called at the top of each protected page (after `st.set_page_config(...)`). Checks state and `st.stop()`s if unauthorized.

By default, **any** `@ihu.gr` account is accepted. To restrict to a specific set of people, add a top-level `allowed_emails = [...]` list in `secrets.toml` (see commented example above). Behavior:

- `allowed_emails` set → only those exact emails (case-insensitive) work, even if `@ihu.gr`.
- `allowed_emails` missing/empty → any `@ihu.gr` account works (current default).

Non-`@ihu.gr` emails are always rejected.

### Azure / Streamlit Cloud setup gotchas

These cost hours during initial setup (2026-05-13); read before debugging auth issues:

- **`[auth]` block is flat** (no `[auth.microsoft]` subsection) → call `st.login()` with **no argument**. Passing `"microsoft"` errors with "provider not found".
- **`httpx` must be in `pyproject.toml`** — `streamlit[auth]` extra alone doesn't pull it in, but Authlib's starlette_client requires it.
- **Do NOT add `client_kwargs = { prompt = "select_account" }`** — caused `MismatchingStateError` on Streamlit Cloud during testing.
- **Streamlit Cloud secrets are SEPARATE** from the local `secrets.toml`. Edit them independently in the Cloud dashboard (Settings → Secrets). Only `redirect_uri` should differ between them.
- Azure App Registration must have **both** redirect URIs registered under the **Web** platform: `http://localhost:8501/oauth2callback` and `https://<app>.streamlit.app/oauth2callback`. Use the plain `/oauth2callback` path — **not** `/~/+/oauth2callback` (that workaround is for Auth0, not Microsoft).

## Active data files

Update these paths inside the page files when switching academic year:

| Page | Active file |
|------|-------------|
| Exam schedule | `files/exams/exams-2026-06.xlsm` |
| Timetable | `files/timetables/2025-2026.xlsm` |

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
