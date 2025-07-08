import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import requests

# Session state
if 'lang' not in st.session_state:
    st.session_state['lang'] = 'Ελληνικά'


# Load data from google sheets
@st.cache_data
def load_gsheet(lang: str) -> pd.DataFrame:
    sheet_id = gsheet_perigrammata_id
    if lang == "Ελληνικά":
        sheet_name = "gr"
    else:
        sheet_name = "eng"
    url = fr"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    df = pd.read_csv(url, dtype_backend='pyarrow', index_col=0)
    return df


def reload_data():
    """Clear cache and reload data from Google Sheets"""
    st.cache_data.clear()
    # Force reload by clearing session state data flags
    if 'perigrammata_data_loaded' in st.session_state:
        del st.session_state['perigrammata_data_loaded']
    # Set a flag to indicate reload was requested
    st.session_state['perigrammata_reload_requested'] = True


def load_data():
    """Load data with session state management"""
    # Check if reload was requested
    if st.session_state.get('perigrammata_reload_requested', False):
        st.session_state['perigrammata_reload_requested'] = False
        if 'perigrammata_data_loaded' in st.session_state:
            del st.session_state['perigrammata_data_loaded']
    
    if 'perigrammata_data_loaded' not in st.session_state:
        st.session_state['df_perigrammata'] = load_gsheet(
            st.session_state['lang'])
        st.session_state['perigrammata_data_loaded'] = True
    
    # Check if language changed and reload if needed
    current_lang = st.session_state['lang']
    if (st.session_state.get('last_lang') != current_lang or
            'df_perigrammata' not in st.session_state):
        st.session_state['df_perigrammata'] = load_gsheet(current_lang)
        st.session_state['last_lang'] = current_lang
    
    return st.session_state['df_perigrammata']


st.sidebar.button('Ενημέρωση από Google Sheets', on_click=reload_data)


def replace_none_with_empty_str(some_dict: dict) -> dict:
    return {k: ('' if v is None else v) for k, v in some_dict.items()}


# Load Google Sheets ID from secrets
try:
    gsheet_perigrammata_id = st.secrets['gsheet_perigrammata_id']
except Exception as e:
    st.error(f"Error loading Google Sheets ID from secrets: {e}")
    st.error("Make sure you have a .streamlit/secrets.toml file with "
             "gsheet_perigrammata_id configured")
    st.stop()

st.markdown('## Περιγράμματα μαθημάτων')

st.radio("Γλώσσα", ("Ελληνικά", "Αγγλικά"),
          key='lang', on_change=reload_data)


# Load data using the improved function
df = load_data()

# st.write(doc.undeclared_template_variables)


tab_table, tab_statistics, tab_word_download = st.tabs(
    ["Πίνακας", "Στατιστικά", "Αρχείο word"])


def make_word_file(row_dict: dict):
    if st.session_state['lang'] == "Ελληνικά":
        url = r"https://github.com/panagop/civil_ihu_pyappz/raw/main/files/perigrammata-template-gr.docx"
    else:
        url = r"https://github.com/panagop/civil_ihu_pyappz/raw/main/files/perigrammata-template-eng.docx"

    response = requests.get(url, timeout=5)
    bytes_io = io.BytesIO(response.content)

    doc = DocxTemplate(bytes_io)

    doc.render(row_dict)
    buffer = io.BytesIO()
    doc.save(buffer)

    return buffer.getvalue()
    


with tab_table:
    st.write(df)


with tab_statistics:
    st.markdown('### Αριθμός μαθημάτων ανά εξάμηνο')
    st.bar_chart(df['examino'].value_counts())

    st.markdown('### Τύπος μαθημάτων')
    st.bar_chart(df['type'].value_counts())

with tab_word_download:

    course_examino = st.selectbox("Επιλέξτε εξάμηνο", df['examino'].unique())
    course_code = st.selectbox("Επιλέξτε κωδικό μαθήματος",
                             df[df['examino'] == course_examino]['code'].unique())

    # Find the matching row directly using boolean indexing
    matching_rows = df[(df['examino'] == course_examino) &
                       (df['code'] == course_code)]
    
    if len(matching_rows) > 0:
        row = matching_rows.iloc[0]
    else:
        st.error("No matching course found!")
        st.stop()
    row_dict = row.to_dict()
    row_dict = replace_none_with_empty_str(row_dict)

    with st.expander("Στοιχεία μαθήματος (πλήρη)"):
        st.write(row_dict)

    btn = st.download_button(
        label="Download file",
        data=make_word_file(row_dict),
        file_name=f"Περίγραμμα-{course_code}-{st.session_state['lang']}.docx",
        mime="document/docx"
    )
