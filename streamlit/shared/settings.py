"""Settings lookup that works in all three deployment targets.

Locally and on Streamlit Cloud the values live in ``.streamlit/secrets.toml``
and are read through ``st.secrets``. On Railway — and any other host that only
offers environment variables — the same keys are read from ``os.environ``.

``st.secrets`` is always tried first, so wherever a secrets file exists the
behaviour is exactly what it was before this module existed.

Note that ``st.secrets`` raises ``StreamlitSecretNotFoundError`` when there is
no secrets file at all; it does not simply report the key as missing. Even
``"key" in st.secrets`` and ``st.secrets.get(key, default)`` raise. That is why
every access here is wrapped rather than using ``.get()``.
"""

from __future__ import annotations

import os
from typing import Any

import streamlit as st


def get_secret(key: str, default: Any = None) -> Any:
    """Return a setting from st.secrets, else the environment, else default."""
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        # No secrets file on this host (e.g. Railway) — fall through to env.
        pass
    return os.environ.get(key, default)


def get_secret_list(key: str) -> list[str]:
    """Same as :func:`get_secret`, for values that are lists.

    A TOML list comes back as a list; an environment variable can only be a
    string, so a comma-separated one is split.
    """
    value = get_secret(key)
    if value is None:
        return []
    if isinstance(value, str):
        return [item.strip() for item in value.split(",") if item.strip()]
    return [str(item).strip() for item in value if str(item).strip()]


def require_secret(key: str) -> str:
    """Return a required setting, or stop the page with a clear message."""
    value = get_secret(key)
    if not value:
        st.error(f"Λείπει η ρύθμιση `{key}`.")
        st.error(
            "Ορίστε την σε ένα αρχείο `.streamlit/secrets.toml` (τοπικά ή "
            "Streamlit Cloud) ή ως μεταβλητή περιβάλλοντος (Railway)."
        )
        st.stop()
    return value
