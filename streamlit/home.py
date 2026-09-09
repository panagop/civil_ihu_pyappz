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

from auth import render_login_block  # noqa: E402
from branding import (  # noqa: E402
    DEPARTMENT_NAME,
    UNIVERSITY_LOGO,
    UNIVERSITY_NAME,
    UNIVERSITY_URL,
    apply_branding,
)

apply_branding()

# One row, department mark on the left (via st.logo, top-left) and the
# university mark on the right.
with st.container(horizontal=True, horizontal_alignment="right"):
    if UNIVERSITY_LOGO.exists():
        st.image(str(UNIVERSITY_LOGO), width=200, link=UNIVERSITY_URL)

st.title(DEPARTMENT_NAME)
st.caption(UNIVERSITY_NAME)

st.markdown(
    "Εσωτερική εφαρμογή του Τμήματος για τα περιγράμματα μαθημάτων, τα μητρώα "
    "εκλεκτόρων, τα συγγράμματα του Ευδόξου, καθώς και τα προγράμματα "
    "εξετάσεων και διδασκαλίας."
)

render_login_block()

st.subheader("Ενότητες", divider="gray")

# Sentence-case descriptions, one line each — the sidebar does the navigating.
SECTIONS = [
    (
        ":material/description:",
        "Περιγράμματα μαθημάτων",
        (
            "Τα περιγράμματα των προγραμμάτων σπουδών 2018 και 2025, "
            "με επεξεργασία και εξαγωγή σε Word."
        ),
    ),
    (
        ":material/groups:",
        "Μητρώα εκλεκτόρων",
        (
            "Εσωτερικοί και εξωτερικοί εκλέκτορες ανά γνωστικό αντικείμενο, "
            "με προτάσεις μεταβολών και συγκεντρωτική αναφορά."
        ),
    ),
    (
        ":material/menu_book:",
        "Εύδοξος",
        (
            "Τα συγγράμματα ανά ακαδημαϊκό έτος και έλεγχος διαθεσιμότητάς "
            "τους για το επόμενο έτος."
        ),
    ),
    (
        ":material/event:",
        "Πρόγραμμα εξετάσεων",
        "Το πρόγραμμα της τρέχουσας εξεταστικής περιόδου.",
    ),
    (
        ":material/calendar_month:",
        "Ωρολόγιο πρόγραμμα",
        "Το εβδομαδιαίο πρόγραμμα διδασκαλίας του τρέχοντος ακαδημαϊκού έτους.",
    ),
]

for icon, title, description in SECTIONS:
    with st.container(border=True):
        st.markdown(f"##### {icon} {title}")
        st.caption(description)

st.caption(
    "Η πλοήγηση γίνεται από το πλαϊνό μενού. Ορισμένες ενότητες απαιτούν "
    "σύνδεση με λογαριασμό @ihu.gr."
)

# Last, deliberately: the schema and the historical data are installed on first
# start and the database is reachable only from inside Railway. Streamlit
# streams the page top to bottom, so everything above is already on screen
# while this runs.
import db  # noqa: E402

if db.is_available():
    st.sidebar.caption(db.bootstrap())
