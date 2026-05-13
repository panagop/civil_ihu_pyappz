# Plan: Add Microsoft authentication to Streamlit app, gating selected pages by @ihu.gr domain

## Context

The civil-ihu-pyappz Streamlit app currently exposes all four pages publicly. We want to restrict two of them — **perigrammata** (course syllabi) and **mitroa** (course registries) — to authenticated IHU staff only. The remaining pages (exam schedule, weekly timetable) stay public.

Since IHU staff already use Microsoft-managed `@ihu.gr` accounts, the simplest, lowest-maintenance approach is **Streamlit's native OIDC auth (`st.login()` / `st.user`)** with **Microsoft Entra ID** as the identity provider, plus an email-suffix check that rejects anyone whose email doesn't end in `@ihu.gr`. No database, no extra Python dependency, no user list to maintain.

Streamlit added native OIDC auth in v1.42 (Feb 2025). The current `pyproject.toml` pins `streamlit` without a version, so a fresh `uv sync` already gets a compatible version — but we will pin it explicitly to be safe.

---

## Approach

### 1. Microsoft Entra ID app registration (one-time, manual, outside code)

Done by you (or IHU IT) in the Azure portal — Claude cannot do this:

1. Azure portal → **Entra ID** → **App registrations** → **New registration**
   - Name: e.g. `civil-ihu-pyappz`
   - Supported account types: **Single tenant** (IHU only) — recommended
   - Redirect URI (Web): `http://localhost:8501/oauth2callback` (dev) — add the production URL later
2. Note the **Application (client) ID** and **Directory (tenant) ID**.
3. **Certificates & secrets** → **New client secret** → copy the secret value immediately.
4. **API permissions** → ensure `openid`, `profile`, `email` (Microsoft Graph, delegated) are granted.

### 2. Configure secrets

Edit [streamlit/.streamlit/secrets.toml](../streamlit/.streamlit/secrets.toml) (gitignored — local only) to add an `[auth]` section alongside the existing Google Sheet IDs:

```toml
# existing keys stay as-is
gsheet_perigrammata_id = "..."
gsheet_mitroa_id = "..."

[auth]
redirect_uri = "http://localhost:8501/oauth2callback"
cookie_secret = "<generate with: python -c \"import secrets; print(secrets.token_hex(32))\">"
client_id = "<Application (client) ID from Azure>"
client_secret = "<client secret value from Azure>"
server_metadata_url = "https://login.microsoftonline.com/<TENANT_ID>/v2.0/.well-known/openid-configuration"
client_kwargs = { prompt = "select_account" }
```

Also update [CLAUDE.md](../CLAUDE.md) — the "Secrets / credentials" section — to document the new `[auth]` block so the next person setting up a dev machine knows it's needed.

### 3. Create a shared auth helper

New file: **[streamlit/auth.py](../streamlit/auth.py)** — placed directly in `streamlit/` (not in `lib/`, which is gitignored).

It exports a single function `require_ihu_login()` that:

- If `st.user.is_logged_in` is False → render an "Σύνδεση με λογαριασμό IHU" button that calls `st.login("microsoft")` (or whatever provider name we use in `secrets.toml`), then `st.stop()`.
- If logged in but `st.user.email` does not end with `@ihu.gr` (case-insensitive) → render an error message, a logout button, and `st.stop()`.
- Otherwise: render a small sidebar block showing the user's name/email and a logout button, then return.

Keeping it as one helper means each protected page adds exactly one line at the top.

### 4. Gate the two protected pages

In **[streamlit/pages/1_📇_perigrammata.py](../streamlit/pages/1_📇_perigrammata.py)** and **[streamlit/pages/2_📊_mitroa.py](../streamlit/pages/2_📊_mitroa.py)**, add — immediately after `st.set_page_config(...)` and before any data loading:

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from auth import require_ihu_login
require_ihu_login()
```

`st.set_page_config` must remain the first Streamlit call on the page; the gate goes right after it so unauthorized users never trigger the Google Sheets fetches or cache fills.

Do **not** modify pages 3 and 4 — they stay public.

### 5. Pin Streamlit ≥ 1.42 in [pyproject.toml](../pyproject.toml)

Change `"streamlit"` → `"streamlit>=1.42"` in the `dependencies` list so the OIDC API is guaranteed.

### 6. Optional polish on home.py

In [streamlit/home.py](../streamlit/home.py), optionally show the login status in the sidebar (without forcing login on the landing page) so users discover the auth flow naturally. This is a small nice-to-have — can be skipped.

---

## Why Microsoft OIDC over the alternatives

| Option | Verdict |
|---|---|
| **Native `st.login()` + Microsoft Entra ID** ✅ | Zero extra deps, SSO with existing `@ihu.gr` accounts, no passwords to manage, MFA inherited from IHU's tenant. |
| `streamlit-authenticator` (YAML) | Adds a dep, requires manual user/password management, no MFA, weaker than SSO IHU already runs. |
| Supabase Auth | Overkill — adds Supabase as a dependency only for auth when IHU's IdP is already available. Worth it only if we also need a Postgres DB, which we don't. |

---

## Files to create / modify

| File | Action |
|---|---|
| [streamlit/auth.py](../streamlit/auth.py) | **Create** — `require_ihu_login()` helper (top-level, not in gitignored `lib/`) |
| [streamlit/pages/1_📇_perigrammata.py](../streamlit/pages/1_📇_perigrammata.py) | Add gate after `st.set_page_config` |
| [streamlit/pages/2_📊_mitroa.py](../streamlit/pages/2_📊_mitroa.py) | Add gate after `st.set_page_config` |
| [streamlit/.streamlit/secrets.toml](../streamlit/.streamlit/secrets.toml) | Add `[auth]` block (manual, local only) |
| [pyproject.toml](../pyproject.toml) | Pin `streamlit>=1.42` |
| [CLAUDE.md](../CLAUDE.md) | Document new `[auth]` secrets keys |

Pages 3 and 4 are untouched.

---

## Verification

1. `uv sync` — confirm Streamlit ≥ 1.42 is installed.
2. `uv run streamlit run streamlit/home.py`
3. **Public pages still work without login:** open the app, navigate to *Exam schedule* and *Weekly timetable* — they render normally with no login prompt.
4. **Protected pages require login:** open *Περιγράμματα* → see the "Σύνδεση με λογαριασμό IHU" button, no syllabus data is fetched. Click it → redirected to Microsoft login → consent → returned to the page.
5. **Domain restriction works:** sign in with a non-`@ihu.gr` Microsoft account (e.g. a personal `@outlook.com`) → page shows the access-denied error and logout button; no data loads.
6. **Happy path:** sign in with your `@ihu.gr` account → page loads syllabi from Google Sheets as today. Sidebar shows your name + logout button. Repeat for *Μητρώα*.
7. **Logout:** click logout → returned to login prompt; reloading the protected page does not bypass auth.
8. **Cache isolation sanity check:** the `@st.cache_data` blocks on both pages key on their inputs (lang/programma_spoudon/sheet name), not user identity — that's fine because the data is the same for all authorized users; the `st.stop()` gate prevents any data fetch before auth succeeds.

---

## Out of scope (explicitly)

- Per-user authorization (everyone with `@ihu.gr` gets equal access to both protected pages).
- Audit logging of who accessed what.
- Protecting pages 3 and 4.
- Production deployment configuration (the redirect URI for production must be added to the Azure app registration when you deploy, but that's a deployment-time concern).
