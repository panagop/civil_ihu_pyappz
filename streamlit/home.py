"""Entry point. Declares the app's pages and runs the selected one.

Every page is named here with ``st.Page``, so the sidebar carries real Greek
titles instead of the ``5_📊_mitroa_v2``-style filenames the legacy ``pages/``
folder produced. The page files keep those names: ``st.Page`` derives a page's
URL from its filename exactly as the old folder did — it drops the leading
number and the emoji — so ``/mitroa_v2``, ``/exams-schedule`` and the rest
still resolve and existing links keep working.

**The folder had to be renamed ``pages`` → ``app_pages``.** Streamlit still
runs the legacy multipage machinery whenever a ``pages/`` directory sits beside
the entry script: it builds its own navigation out of the folder and runs the
chosen page itself, so ``st.navigation`` below would never be reached.

This script runs on every rerun, before the page. Keep it to navigation — work
placed here is work every page pays for. That is why the landing page is a page
like any other (``app_pages/0_home.py``) rather than the body of this file, and
why the database bootstrap stayed there with it.
"""

import warnings

import streamlit as st

# openpyxl warns on every read of .xlsm files that contain data-validation rules
# (dropdown lists etc.). We only read these files, so the warning is noise.
warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported and will be removed",
    category=UserWarning,
    module="openpyxl",
)

st.set_page_config(
    page_title="Πολιτικοί Μηχανικοί — ΔΙΠΑΕ",
    page_icon=":material/foundation:",
    initial_sidebar_state="expanded",
)

# No `layout` here on purpose: the pages that want a wide one set it themselves,
# and naming it here would fight them.
#
# Titles are the ones each page already declares in its own `set_page_config`,
# so the sidebar label and the browser tab agree. The icons are the emoji the
# filenames carry — the legacy folder read them out of the names; `st.Page` is
# told them instead. The order is the one the leading numbers used to impose.
PAGES = [
    st.Page(
        "app_pages/0_home.py",
        title="Αρχική",
        icon=":material/foundation:",
        default=True,
    ),
    st.Page(
        "app_pages/1_📇_perigrammata (legacy).py",
        title="Περιγράμματα (παλαιό)",
        icon="📇",
    ),
    st.Page(
        "app_pages/3_⛱_exams-schedule.py",
        title="Πρόγραμμα Εξετάσεων",
        icon="⛱",
    ),
    st.Page(
        "app_pages/4_📅_weekly_timetable.py",
        title="Εβδομαδιαίο Πρόγραμμα",
        icon="📅",
    ),
    st.Page(
        "app_pages/5_📊_mitroa_v2.py",
        title="Μητρώα γνωστικών αντικειμένων",
        icon="📊",
    ),
    st.Page(
        "app_pages/6_📇_perigrammata_v2.py",
        title="Περιγράμματα μάθημάτων",
        icon="📇",
    ),
    st.Page(
        "app_pages/7_📚_eudoxus.py",
        title="Εύδοξος - Συγγράμματα",
        icon="📚",
    ),
    st.Page(
        "app_pages/8_🗓_timetable_v2.py",
        title="Εβδομαδιαίο πρόγραμμα v2",
        icon="🗓",
    ),
]

st.navigation(PAGES).run()
