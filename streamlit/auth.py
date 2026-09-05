"""Authentication helpers for IHU staff pages.

Uses Streamlit's native OIDC auth (st.login / st.user, available since v1.42)
with Microsoft Entra ID as the identity provider. Configured in
.streamlit/secrets.toml under [auth].

The login flow lives on the home page (see streamlit/home.py). Protected
pages only *check* the user's state; they do not trigger st.login themselves,
because Streamlit's OIDC callback always returns the user to the app root.
"""

from __future__ import annotations

import streamlit as st

from settings import get_secret_list

ALLOWED_EMAIL_SUFFIX = "@ihu.gr"


def _allowed_emails() -> set[str]:
    """Lower-cased set of explicitly allowlisted emails.

    Read from secrets.toml locally / on Streamlit Cloud, or from an
    `allowed_emails` environment variable (comma-separated) on Railway.
    If missing/empty, only the domain-suffix check applies (any @ihu.gr
    account works).
    """
    return {e.lower() for e in get_secret_list("allowed_emails")}


def is_configured() -> bool:
    """True iff an OIDC provider is set up in this environment.

    ``st.login()`` raises ``StreamlitAuthError`` when ``[auth]`` is missing, and
    ``[auth]`` is nested TOML with no environment-variable equivalent — so a
    host that only offers env vars has none unless
    ``scripts/write_secrets_toml.py`` ran first. Check before offering a login
    button, or the button is a traceback waiting to be clicked.
    """
    try:
        return "auth" in st.secrets
    except Exception:
        return False  # no secrets file at all on this host


def _email_allowed(email: str | None) -> bool:
    if not email:
        return False
    email = email.lower()
    if not email.endswith(ALLOWED_EMAIL_SUFFIX):
        return False
    allowlist = _allowed_emails()
    return not allowlist or email in allowlist


def is_authorized() -> bool:
    """True iff the current user is logged in with an allowed email."""
    user = st.user
    return bool(getattr(user, "is_logged_in", False)) and _email_allowed(
        getattr(user, "email", None)
    )


def require_ihu_login() -> None:
    """Gate the current page. Call once, right after st.set_page_config().

    If unauthorized, render a message directing the user to the home page
    and call st.stop(). Does NOT trigger st.login() — that lives on home.py.
    """
    user = st.user

    if not getattr(user, "is_logged_in", False):
        st.markdown("## 🔒 Απαιτείται σύνδεση")
        st.info(
            "Παρακαλώ συνδεθείτε από την **αρχική σελίδα** "
            "(👈 Home στην πλαϊνή στήλη) με τον λογαριασμό σας "
            f"`{ALLOWED_EMAIL_SUFFIX}`."
        )
        st.stop()

    email = getattr(user, "email", None)
    if not _email_allowed(email):
        st.error(
            f"Ο λογαριασμός **{email or 'άγνωστος'}** δεν έχει πρόσβαση. "
            "Επικοινωνήστε με τον διαχειριστή αν χρειάζεστε πρόσβαση."
        )
        if st.button("Αποσύνδεση"):
            st.logout()
        st.stop()

    _render_sidebar_user(user, email)


def render_login_block() -> None:
    """Render the login/logout UI on the home page."""
    user = st.user

    if not getattr(user, "is_logged_in", False):
        if not is_configured():
            st.warning(
                "Η σύνδεση δεν είναι ρυθμισμένη σε αυτό το περιβάλλον "
                "(λείπει το `[auth]` από τα secrets)."
            )
            return
        st.info(
            "Ορισμένες σελίδες (Περιγράμματα, Μητρώα) απαιτούν σύνδεση "
            f"με λογαριασμό `{ALLOWED_EMAIL_SUFFIX}`."
        )
        if st.button("Σύνδεση με λογαριασμό IHU", type="primary"):
            st.login()
        return

    email = getattr(user, "email", None)
    if not _email_allowed(email):
        st.error(
            f"Ο λογαριασμός **{email or 'άγνωστος'}** δεν έχει πρόσβαση. "
            "Επικοινωνήστε με τον διαχειριστή αν χρειάζεστε πρόσβαση."
        )
        if st.button("Αποσύνδεση"):
            st.logout()
        return

    name = getattr(user, "name", None) or email
    st.success(f"Συνδεδεμένος ως **{name}** ({email}).")
    if st.button("Αποσύνδεση"):
        st.logout()

    _render_sidebar_user(user, email)


def _render_sidebar_user(user, email: str) -> None:
    with st.sidebar:
        name = getattr(user, "name", None) or email
        st.caption(f"👤 {name}")
        st.caption(email)
        if st.button("Αποσύνδεση", key="_auth_sidebar_logout"):
            st.logout()
