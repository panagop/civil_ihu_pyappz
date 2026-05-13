import warnings

import streamlit as st
import pandas as pd
import json

# openpyxl warns on every read of .xlsm files that contain data-validation rules
# (dropdown lists etc.). We only read these files, so the warning is noise.
warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported and will be removed",
    category=UserWarning,
    module="openpyxl",
)


st.set_page_config(page_title="Περιγράμματα μαθημάτων", page_icon=":house:", initial_sidebar_state="expanded")