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

            # Βεβαιώσου ότι τα ονόματα στηλών ταιριάζουν με αυτά
        required_cols = [
            "course_id",
            "course_name",
            "class_name",
            "semester",
            "teaching_period",
            "instructors",
            "day",
            "start_time",
            "duration",
            "room",
            "notes",
        ]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            raise ValueError(f"Λείπουν οι στήλες: {missing}")
        
        # Φιλτράρισμα με βάση το teaching_period
        df = df[df['teaching_period'] == semester_selection]
            
    except Exception as e:
        st.error(f"❌ Σφάλμα κατά το άνοιγμα του αρχείου: {e}")
        st.stop()
    
    return df


# Φόρτωση δεδομένων
try:
    df = load_data(semester_selection)
    
    st.subheader(f"Πρόγραμμα {semester_selection} Εξαμήνου 2025-2026")
    
    # Tabs
    tab_table, tab_calendar = st.tabs(["Πίνακας", "Εβδομαδιαία Προβολή"])
    
    with tab_table:
        # Εμφάνιση δεδομένων
        st.dataframe(df, use_container_width=True)
    
    with tab_calendar:
        st.markdown("### Εβδομαδιαίο Πρόγραμμα")
        
        # Φίλτρο εξαμήνων σπουδών
        semesters_all = sorted(df["semester"].unique().tolist())
        semester_options = [f"Εξάμηνο {int(s)}" for s in semesters_all]
        
        selected_semesters = st.multiselect(
            "Φιλτράρισμα κατά εξάμηνο σπουδών:",
            options=semester_options,
            default=semester_options,
            key="semester_filter"
        )
        
        # Φιλτράρισμα δεδομένων
        if selected_semesters and len(selected_semesters) < len(semester_options):
            semester_nums = [int(s.split()[-1]) for s in selected_semesters]
            df_filtered = df[df["semester"].isin(semester_nums)]
        else:
            df_filtered = df
        
        # Χρώματα ανά εξάμηνο
        semester_colors = {
            1: '#E74C3C',  2: '#3498DB',  3: '#2ECC71',  4: '#F39C12',  5: '#9B59B6',
            6: '#1ABC9C',  7: '#E67E22',  8: '#34495E',  9: '#16A085',  10: '#D35400',
        }
        
        # Δημιουργία πίνακα εβδομαδιαίου προγράμματος
        days_greek = ['Δευτέρα', 'Τρίτη', 'Τετάρτη', 'Πέμπτη', 'Παρασκευή']
        
        # Συλλογή μοναδικών ωρών έναρξης
        unique_times = sorted(df_filtered['start_time'].dropna().unique())
        
        # Δημιουργία πίνακα
        st.markdown("---")
        
        # Header row
        cols = st.columns([1] + [3]*5)
        cols[0].markdown("**Ώρα**")
        for i, day in enumerate(days_greek):
            cols[i+1].markdown(f"**{day}**")
        
        # Data rows
        for time_slot in unique_times:
            cols = st.columns([1] + [3]*5)
            cols[0].markdown(f"**{time_slot}**")
            
            for day_idx, day in enumerate(days_greek):
                # Εύρεση μαθημάτων για αυτή την ημέρα και ώρα
                day_classes = df_filtered[
                    (df_filtered['day'] == day) & 
                    (df_filtered['start_time'] == time_slot)
                ]
                
                if not day_classes.empty:
                    with cols[day_idx+1]:
                        for _, class_row in day_classes.iterrows():
                            semester = int(class_row['semester']) if pd.notna(class_row['semester']) else 1
                            color = semester_colors.get(semester, '#95A5A6')
                            
                            class_info = f"""
                            <div style="background-color: {color}; padding: 8px; margin: 4px 0; border-radius: 4px; color: white; font-size: 12px;">
                                <strong>Εξ.{semester} - {class_row['course_name']}</strong><br/>
                                {class_row['instructors']}<br/>
                                <small>{class_row['room']} | {class_row['duration']} ώρες</small>
                            </div>
                            """
                            st.markdown(class_info, unsafe_allow_html=True)
                else:
                    cols[day_idx+1].markdown("")
        
        st.markdown("---")
        st.write(f"📚 Σύνολο μαθημάτων: {len(df_filtered)}")
        
except Exception as e:
    st.error(f"Σφάλμα: {e}")
    st.info("Παρακαλώ ελέγξτε τη δομή του αρχείου Excel και τα ονόματα των sheets.")


