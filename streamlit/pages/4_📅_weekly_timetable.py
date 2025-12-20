import pandas as pd
from pathlib import Path
import streamlit as st
from streamlit_calendar import calendar
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

st.set_page_config(
    layout="wide",
    page_title="Εβδομαδιαίο Πρόγραμμα Μαθημάτων",
    page_icon="📅",
)

st.title("📅 Εβδομαδιαίο Πρόγραμμα Μαθημάτων")

# Επιλογή εξαμήνου
semester_selection = st.radio(
    "Επιλέξτε εξάμηνο:",
    options=["Χειμερινό", "Εαρινό"],
    index=0,
    key="semester_selection"
)

st.markdown(f"Έχετε επιλέξει: **{semester_selection} Εξάμηνο**")

# Ρυθμίσεις αρχείου
INPUT_EXCEL = Path(__file__).parent.parent.parent / "files" / "timetables" / "2025-2026.xlsm"


def load_data(semester: str) -> pd.DataFrame:
    """Διαβάζει τα δεδομένα από το Excel."""
    
    # Έλεγχος ύπαρξης αρχείου
    if not INPUT_EXCEL.exists():
        st.error(f"❌ Το αρχείο {INPUT_EXCEL} δεν βρέθηκε!")
        st.info(f"Αναζητούμενη διαδρομή: {INPUT_EXCEL.absolute()}")
        st.stop()
    
    try:
        # Έλεγχος διαθέσιμων sheets
        excel_file = pd.ExcelFile(INPUT_EXCEL)
        available_sheets = excel_file.sheet_names
        
        # Προσδιορισμός sheet name με βάση το εξάμηνο
        sheet_name = 'timetable'  # Προσαρμόστε ανάλογα με τα πραγματικά ονόματα των sheets
        
        if sheet_name not in available_sheets:
            st.error(f"❌ Το sheet '{sheet_name}' δεν βρέθηκε στο αρχείο!")
            st.info(f"Διαθέσιμα sheets: {', '.join(available_sheets)}")
            st.stop()
        
        df = pd.read_excel(INPUT_EXCEL, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"❌ Σφάλμα κατά το άνοιγμα του αρχείου: {e}")
        st.stop()
    
    return df


# Φόρτωση δεδομένων
try:
    df = load_data(semester_selection)
    
    st.subheader(f"Πρόγραμμα {semester_selection} Εξαμήνου 2025-2026")
    
    # Εμφάνιση δεδομένων
    st.dataframe(df, use_container_width=True)
    
        
except Exception as e:
    st.error(f"Σφάλμα: {e}")
    st.info("Παρακαλώ ελέγξτε τη δομή του αρχείου Excel και τα ονόματα των sheets.")

