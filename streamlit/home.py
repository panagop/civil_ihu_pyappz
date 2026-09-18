"""Entry point. Declares the app's pages and runs the selected one.

Every page is named here with ``st.Page``, so the sidebar carries real Greek
titles instead of the ``5_📊_mitroa_v2``-style filenames the legacy ``pages/``
folder produced. The page files keep those names: ``st.Page`` derives a page's
URL from its filename exactly as the old folder did — it drops the leading
number and the emoji — so ``/mitroa_v2``, ``/exams-schedule`` and the rest
still resolve and existing links keep working.

That rule is why ``8_🗓_timetable_v2.py`` was renamed to
``8_📅_weekly_timetable.py`` when it replaced page 4 (2026-09-18): the name
hands it ``/weekly_timetable``, the URL the workbook-backed page used to
answer, so links people already have keep working and now reach the database
version. The old page, renamed in turn, moved to ``/weekly_timetable_(legacy)``.

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
    page_title="Ψηφιακές υπηρεσίες Τμήματος Πολιτικών Μηχανικών — ΔΙΠΑΕ",
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
        "app_pages/3_⛱_exams-schedule.py",
        title="Πρόγραμμα Εξετάσεων",
        icon="⛱",
    ),
    st.Page(
        "app_pages/8_📅_weekly_timetable.py",
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
]

# Superseded pages: the Google Sheets περιγράμματα and the workbook timetable.
# They are kept because their replacements are young, not because anyone should
# open them, so they stay out of the navigation unless the box below is ticked.
# A page that is not in the list Streamlit is given never runs at all.
LEGACY_PAGES = [
    st.Page(
        "app_pages/1_📇_perigrammata (legacy).py",
        title="Περιγράμματα (παλαιό)",
        icon="📇",
    ),
    st.Page(
        "app_pages/4_📅_weekly_timetable (legacy).py",
        title="Εβδομαδιαίο Πρόγραμμα (παλαιό)",
        icon="📅",
    ),
]

# In the sidebar, so it is reachable from every page — and here in the router
# rather than on a page, because the list below is built before any page runs.
# Streamlit renders its navigation at the top of the sidebar whatever the code
# order, so this lands underneath it.
show_legacy = st.sidebar.checkbox(
    "Παλαιές σελίδες",
    key="show_legacy_pages",
    help=(
        "Εμφανίζει τις σελίδες που έχουν αντικατασταθεί: τα Περιγράμματα από "
        "το Google Sheet και το Εβδομαδιαίο Πρόγραμμα από το αρχείο Excel. "
        "Διατηρούνται μόνο για σύγκριση."
    ),
)

# Unticking it while a legacy page is open is safe: Streamlit answers a path it
# was not given with the default page, so the user lands on «Αρχική».
st.navigation([*PAGES, *LEGACY_PAGES] if show_legacy else PAGES).run()
