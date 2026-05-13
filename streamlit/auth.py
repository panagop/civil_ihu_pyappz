"""Authentication helpers for IHU staff pages.

Uses Streamlit's native OIDC auth (st.login / st.user, available since v1.42)
with Microsoft Entra ID as the identity provider. Configured in
.streamlit/secrets.toml under [auth].
"""

from __future__ import annotations

import streamlit as st

ALLOWED_EMAIL_SUFFIX = "@ihu.gr"


def _email_allowed(email: str | None) -> bool:
    return bool(email) and email.lower().endswith(ALLOWED_EMAIL_SUFFIX)


def require_ihu_login() -> None:
    """Gate the current page behind a Microsoft login restricted to @ihu.gr.

    Call this once, immediately after st.set_page_config(...), before any
    data loading. If the user is not authorized, the function renders a
    login (or access-denied) UI and calls st.stop().
    """
    user = st.user

    if not getattr(user, "is_logged_in", False):
        st.markdown("## 🔒 Απαιτείται σύνδεση")
        st.write(
            "Η σελίδα αυτή είναι διαθέσιμη μόνο σε μέλη του ΔΠΘ/ΔΙΠΑΕ "
            f"με λογαριασμό **{ALLOWED_EMAIL_SUFFIX}**."
        )
        if st.button("Σύνδεση με λογαριασμό IHU", type="primary"):
            st.login()
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

    with st.sidebar:
        name = getattr(user, "name", None) or email
        st.caption(f"👤 {name}")
        st.caption(email)
        if st.button("Αποσύνδεση", key="_auth_logout"):
            st.logout()
