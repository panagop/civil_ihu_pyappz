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

ALLOWED_EMAIL_SUFFIX = "@ihu.gr"


def _email_allowed(email: str | None) -> bool:
    return bool(email) and email.lower().endswith(ALLOWED_EMAIL_SUFFIX)


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
            f"Απαιτείται email που λήγει σε {ALLOWED_EMAIL_SUFFIX}."
        )
        if st.button("Αποσύνδεση"):
            st.logout()
        st.stop()

    _render_sidebar_user(user, email)


def render_login_block() -> None:
    """Render the login/logout UI on the home page."""
    user = st.user

    if not getattr(user, "is_logged_in", False):
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
            f"Απαιτείται email που λήγει σε {ALLOWED_EMAIL_SUFFIX}."
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
