import streamlit as st
import pandas as pd
import json

import io
import requests


try:
    with open('../files/keys.json') as f:
    # path is relative to app.py, not this file
        data = json.load(f)
        gsheet_mitroa_id = data['gsheet_mitroa']
except:
    gsheet_mitroa_id = st.secrets['gsheet_mitroa']


st.markdown('## Μητρώα γνωστικών αντικειμένων')

@st.cache_data
def load_gsheet(sheet_name) -> pd.DataFrame:
    sheet_id = gsheet_mitroa_id
    url = fr"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    df = pd.read_csv(url, dtype_backend='pyarrow', index_col=0)
    return df


df_eklektores = load_gsheet('eklektores')
df_antikeimena = load_gsheet('antikeimena')


tab_table_eklektores, tab_table_antikeimena, tab_statistics = st.tabs(
    ["Πίνακας εκλεκτόρων", "Πίνακας αντικειμένων", "Στατιστικά"])

with tab_table_eklektores:
    st.markdown('### Εκλεκτορες')
    st.dataframe(df_eklektores)

with tab_table_antikeimena: 
    st.markdown('### Αντικείμενα')
    st.dataframe(df_antikeimena)

with tab_statistics:
    st.markdown('### Κατηγορία Χρήστη')
    st.bar_chart(df_eklektores['Κατηγορία Χρήστη'].value_counts())

    st.markdown('### Φορέας')
    st.bar_chart(df_eklektores['Φορέας Χρήστη'].value_counts())