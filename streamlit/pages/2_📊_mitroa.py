import streamlit as st
import pandas as pd
import json
import pyarrow as pa
import numpy as np

import io
import requests


# Load Google Sheets ID from secrets
try:
    gsheet_mitroa_id = st.secrets['gsheet_mitroa_id']
except Exception as e:
    st.error(f"Error loading Google Sheets ID from secrets: {e}")
    st.error("Make sure you have a .streamlit/secrets.toml file with "
             "gsheet_mitroa_id configured")
    st.stop()


st.markdown('## Μητρώα γνωστικών αντικειμένων')

@st.cache_data
def load_gsheet(sheet_name) -> pd.DataFrame:
    sheet_id = gsheet_mitroa_id
    url = fr"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    df = pd.read_csv(url, dtype_backend='pyarrow', index_col=0)
    return df

def reload():
    st.cache_data.clear()
    df_eklektores = load_gsheet('eklektores')
    df_antikeimena = load_gsheet('antikeimena')

df_eklektores = load_gsheet('eklektores')
df_antikeimena = load_gsheet('antikeimena')

st.sidebar.button('Ενημέρωση από Google Sheets', on_click=reload)

df_antikeimena['Εξωτερικοί Ιδίου'] = df_antikeimena['Εξωτερικοί Ιδίου'].fillna('')


tab_table_eklektores, tab_table_antikeimena, tab_statistics, tab_reports = st.tabs(
    ["Πίνακας εκλεκτόρων", "Πίνακας αντικειμένων", "Στατιστικά", "Εξαγωγή αναφορών"])

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

    st.markdown('### Βαθμίδα')
    st.bar_chart(df_eklektores['Βαθμίδα'].value_counts())

    st.markdown('### Επιστημονικό πεδίο')
    st.bar_chart(df_antikeimena['Επιστημονικό πεδίο'].value_counts())


def get_codes_for_eklektores(df: pd.DataFrame, charaktirismos: str, selected_antikeimeno:str) -> list[int]:
    codes = df[df['Γνωστικό αντικείμενο'] == selected_antikeimeno][charaktirismos].values[0].split('-')
    if '' in codes: 
        codes.remove('')
    codes = [int(i) for i in codes]
    return codes


with tab_reports:
    antikeimena_list = sorted(df_antikeimena['Γνωστικό αντικείμενο'].unique())
    selected_antikeimeno = st.selectbox('Επιλογή αντικειμένου', antikeimena_list)

    codes_external_idiou = get_codes_for_eklektores(df_antikeimena, 'Εξωτερικοί Ιδίου', selected_antikeimeno)
    codes_external_synafous = get_codes_for_eklektores(df_antikeimena, 'Εξωτερικοί Συναφούς', selected_antikeimeno)
    codes = codes_external_idiou + codes_external_synafous
    # st.write(codes_external_idiou)
    # st.write(codes_external_synafous)
    # st.write(codes)

    df_antikeimeno_selected = df_eklektores[df_eklektores.index.isin(codes)]

    df_antikeimeno_selected['Χαρακτηρισμός'] = np.where(df_antikeimeno_selected.index.isin(codes_external_idiou), 'Ιδίου', 'Συναφούς')
    df_antikeimeno_selected = df_antikeimeno_selected.fillna('')
    df_antikeimeno_selected = df_antikeimeno_selected.sort_values(by=['Χαρακτηρισμός', 'Επώνυμο', 'Όνομα'])
    col_to_move = df_antikeimeno_selected.pop('Χαρακτηρισμός')
    df_antikeimeno_selected.insert(0, col_to_move.name, col_to_move)

    st.dataframe(df_antikeimeno_selected)

    # st.write(codes_external_synafous_str)
    
    buffer = io.BytesIO()
    df_antikeimeno_selected.to_excel(buffer)
    # doc.save('gen/perigramma.docx')

    btn = st.download_button(
        label="Download file",
        data=buffer.getvalue(),
        file_name=f"{selected_antikeimeno}.xlsx"
    )

