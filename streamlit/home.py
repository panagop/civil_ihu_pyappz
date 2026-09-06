import warnings

import streamlit as st
# import pandas as pd
# import json

# openpyxl warns on every read of .xlsm files that contain data-validation rules
# (dropdown lists etc.). We only read these files, so the warning is noise.
warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported and will be removed",
    category=UserWarning,
    module="openpyxl",
)


st.set_page_config(page_title="Περιγράμματα μαθημάτων", page_icon=":house:", initial_sidebar_state="expanded")

from auth import render_login_block  # noqa: E402

st.title("Civil Engineering — IHU")
render_login_block()

# The Postgres database is reachable only from inside Railway, so its schema
# and the historical data are installed here, on first start. Runs once per
# process and is a no-op everywhere else (no DATABASE_URL).
import db  # noqa: E402

if db.is_available():
    st.sidebar.caption(db.bootstrap())