"""Write the OIDC ``[auth]`` block to secrets.toml before Streamlit starts.

``st.secrets`` reads TOML files only, and Streamlit's own ``st.login()`` reads
``[auth]`` straight out of them — ``settings.get_secret`` cannot paper over that
the way it does for the flat keys, because a host like Railway offers no way to
express a nested table as an environment variable.

So on such a host we write the file ourselves, from flat variables, before the
server boots. Run from the start command:

    python scripts/write_secrets_toml.py && streamlit run streamlit/home.py ...

It is a no-op when the auth variables are absent (local, Streamlit Cloud), and
it never overwrites an existing secrets.toml.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

# Written next to the entry script: of the three locations Streamlit searches
# (~/.streamlit, <cwd>/.streamlit, <script dir>/.streamlit) this one wins, and
# it is where the file already lives during local development.
TARGET = Path(__file__).resolve().parents[1] / "streamlit" / ".streamlit" / "secrets.toml"

REQUIRED = {
    "client_id": "AUTH_CLIENT_ID",
    "client_secret": "AUTH_CLIENT_SECRET",
    "cookie_secret": "AUTH_COOKIE_SECRET",
    "server_metadata_url": "AUTH_SERVER_METADATA_URL",
}


def redirect_uri() -> str | None:
    """Explicit if given, otherwise derived from the host's public domain.

    Deriving it keeps the value correct if the Railway domain is renamed — but
    note that whatever it resolves to must also be registered in the Azure App
    Registration, so a rename still needs a visit to the portal.
    """
    explicit = os.environ.get("AUTH_REDIRECT_URI")
    if explicit:
        return explicit
    domain = os.environ.get("RAILWAY_PUBLIC_DOMAIN")
    return f"https://{domain}/oauth2callback" if domain else None


def main() -> int:
    present = {key: os.environ.get(var) for key, var in REQUIRED.items()}
    if not any(present.values()):
        print("[write_secrets_toml] Καμία μεταβλητή AUTH_* — παραλείπεται.")
        return 0

    missing = [REQUIRED[key] for key, value in present.items() if not value]
    if missing:
        # Half-configured is worse than not configured: fail loudly rather than
        # let Streamlit start and break at the login button.
        print(f"[write_secrets_toml] ΣΦΑΛΜΑ: λείπουν {', '.join(missing)}", file=sys.stderr)
        return 1

    uri = redirect_uri()
    if not uri:
        print(
            "[write_secrets_toml] ΣΦΑΛΜΑ: ορίστε AUTH_REDIRECT_URI "
            "(δεν υπάρχει RAILWAY_PUBLIC_DOMAIN για να παραχθεί)",
            file=sys.stderr,
        )
        return 1

    if TARGET.exists():
        print(f"[write_secrets_toml] Υπάρχει ήδη {TARGET} — δεν πειράζεται.")
        return 0

    values = {"redirect_uri": uri, **present}
    # json.dumps for the values: TOML basic strings share JSON's escaping, so a
    # secret containing a quote or a backslash cannot break the file.
    body = "\n".join(f"{key} = {json.dumps(value)}" for key, value in values.items())

    TARGET.parent.mkdir(parents=True, exist_ok=True)
    TARGET.write_text(f"[auth]\n{body}\n", encoding="utf-8")
    try:
        TARGET.chmod(0o600)
    except OSError:
        pass  # not all filesystems support it; the container is single-tenant

    print(f"[write_secrets_toml] Γράφτηκε {TARGET} (redirect_uri={uri})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
