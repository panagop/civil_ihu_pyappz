import streamlit as st
import pandas as pd
import numpy as np
import io

# Load Google Sheets ID from secrets
try:
    gsheet_exams_schedule_id = st.secrets['gsheet_exams_schedule_id']
except Exception as e:
    st.error(f"Error loading Google Sheets ID from secrets: {e}")
    st.error("Make sure you have a .streamlit/secrets.toml file with "
             "gsheet_mitroa_id configured")
    st.stop()

st.markdown('## Πρόγραμμα εξετάσεων')


@st.cache_data
def load_gsheet(sheet_name) -> pd.DataFrame:
    sheet_id = gsheet_exams_schedule_id
    url = fr"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    df = pd.read_csv(url, dtype_backend='pyarrow', index_col=0)
    return df


def reload():
    """Clear cache to force reload from Google Sheets"""
    st.cache_data.clear()


def get_data():
    """Get current data from sheets"""
    df_september = load_gsheet('september_data')
    df_september = df_september.dropna(subset=['exams_date'])
    df_september = df_september.reset_index(drop=False)
    return df_september


def create_pivot_table(df):
    """Create pivot table showing course details by semester and course name"""
    
    # Method 1: Simple grouping (most readable)
    pivot_simple = df.groupby(['semester', 'course_name']).agg({
        'teacher': 'first',
        'number_of_students': 'first', 
        'exams_date': 'first'
    }).reset_index()
    
    # Method 2: Multi-index pivot (more traditional pivot table)
    pivot_multi = df.pivot_table(
        index=['semester', 'course_name'],
        values=['teacher', 'number_of_students', 'exams_date'],
        aggfunc='first'
    )
    
    return pivot_simple, pivot_multi


# Load data
df_september = get_data()

st.sidebar.button('Ενημέρωση από Google Sheets', on_click=reload)

# Create tabs for different views
tab_raw, tab_pivot_simple, tab_pivot_multi = st.tabs([
    "Ακατέργαστα δεδομένα", 
    "Απλός πίνακας", 
    "Πίνακας συγκέντρωσης"
])

with tab_raw:
    st.markdown("### Ακατέργαστα δεδομένα")
    st.dataframe(df_september)

# Check if we have the expected columns
expected_columns = ['course_id', 'course_name', 'semester', 'teacher',
                    'number_of_students', 'exams_date']
if all(col in df_september.columns for col in expected_columns):
    
    pivot_simple, pivot_multi = create_pivot_table(
        df_september)
    
    with tab_pivot_simple:
        st.markdown("### Απλός πίνακας (Πιο ευανάγνωστος)")
        st.markdown("Ομαδοποίηση ανά εξάμηνο και μάθημα")
        st.dataframe(pivot_simple)
    
    with tab_pivot_multi:
        st.markdown("### Πίνακας συγκέντρωσης (Multi-index)")
        st.markdown("Παραδοσιακός pivot table με πολλαπλούς δείκτες")
        st.dataframe(pivot_multi)
    

else:
    missing_cols = [col for col in expected_columns
                    if col not in df_september.columns]
    st.error(f"Λείπουν οι εξής στήλες: {missing_cols}")
    st.info(f"Διαθέσιμες στήλες: {list(df_september.columns)}")

