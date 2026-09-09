"""Shared branding for every page.

``st.logo`` applies to the page it is called from, not to the app, so each page
script has to call it — hence a helper rather than one line in ``home.py``.
Pages 3 and 4 are public and import no other shared module, which is why this
one deliberately depends on nothing but Streamlit.

The logo files are committed under ``files/logos/`` rather than hot-linked from
civil.ihu.gr: the site's URLs are CMS-generated (``/wp-content/uploads/2026/02/``)
and will move, a remote fetch costs a round trip on every page load, and
hot-linking would send every visitor's IP address to the department's server.
"""

from __future__ import annotations

from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
LOGO_DIR = ROOT / "files" / "logos"
DEPARTMENT_LOGO = LOGO_DIR / "civil_ihu_logo.png"
UNIVERSITY_LOGO = LOGO_DIR / "ihu_logo.png"

DEPARTMENT_URL = "https://www.civil.ihu.gr/"
UNIVERSITY_URL = "https://www.ihu.gr/"

DEPARTMENT_NAME = "Τμήμα Πολιτικών Μηχανικών"
UNIVERSITY_NAME = "Διεθνές Πανεπιστήμιο της Ελλάδος"


def apply_branding() -> None:
    """Put the department mark in the header and the sidebar.

    Never raises on a missing file: a logo that failed to deploy should cost the
    branding, not the page.
    """
    if DEPARTMENT_LOGO.exists():
        st.logo(str(DEPARTMENT_LOGO), size="large", link=DEPARTMENT_URL)
