import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import streamlit as st
from streamlit_calendar import calendar

# ------------ ΡΥΘΜΙΣΕΙΣ ΧΡΗΣΤΗ ------------
# INPUT_EXCEL = "lessons-calendars.xlsm"   # το αρχείο όπου έχεις το sheet Data
INPUT_SHEET = "ExamsJan26"
INPUT_EXCEL = Path(__file__).parent.parent.parent / "jupyter" / "programmata" / "lessons-calendars.xlsm"


@st.cache_data
def load_data() -> pd.DataFrame:
    """Διαβάζει τα δεδομένα από το Excel (sheet Data)."""
    df = pd.read_excel(INPUT_EXCEL, sheet_name=INPUT_SHEET)

    # Βεβαιώσου ότι τα ονόματα στηλών ταιριάζουν με αυτά
    required_cols = [
        "course_id",
        "course_name",
        "semester",
        "instructor",
        "exam_date",
        "start_time",
        "room",
        "notes",
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Λείπουν οι στήλες: {missing}")

    # Μετατροπές τύπων
    df["exam_date"] = pd.to_datetime(df["exam_date"]).dt.date  # μόνο ημερομηνία
    
    # Drop rows where exam_date is missing
    df = df.dropna(subset=["exam_date"])

    # Αν start_time είναι string τύπου "09:00"
    df["start_time"] = df["start_time"].astype(str)

    # Συνένωση σε datetime για αρχή
    df["start_dt"] = pd.to_datetime(
        df["exam_date"].astype(str) + " " + df["start_time"],
        errors="coerce",
    )

    # Υπολογισμός end_dt με default διάρκεια
    # Στήλη εβδομάδας (για weekly views)
    df["week_number"] = df["start_dt"].dt.isocalendar().week

    return df

def reload():
    """Clear cache to force reload from Google Sheets"""
    st.cache_data.clear()

tab_full_table, tab_instructor_filter, tab_semester_filter, tab_calendar = st.tabs(
    [
        "Πλήρης Πίνακας Εξετάσεων",
        "Φιλτράρισμα κατά Διδάσκοντα",
        "Φιλτράρισμα κατά Εξάμηνο",
        "Ημερολόγιο Εξετάσεων"
    ]
)    

df = load_data()


with tab_full_table:
    st.subheader("Πλήρης Πίνακας Εξετάσεων")
    st.dataframe(df)

instructors = sorted(df["instructor"].unique().tolist())


with tab_instructor_filter:
    selected_instructor = st.selectbox(
        "Επιλέξτε διδάσκοντα για φιλτράρισμα:",
        options=instructors)

    df_instr = df[df["instructor"] == selected_instructor].sort_values(
        by=["start_dt"]
    )  

    st.subheader(f"Πρόγραμμα Εξετάσεων Διδάσκοντα - {selected_instructor}")
    st.dataframe(df_instr)

with tab_semester_filter:
    semesters = sorted(df["semester"].unique().tolist())
    selected_semester = st.selectbox(
        "Επιλέξτε εξάμηνο για φιλτράρισμα:",
        options=semesters
    )

    df_sem = df[df["semester"] == selected_semester].sort_values(
        by=["start_dt"]
    )  

    st.subheader(f"Πρόγραμμα Εξετάσεων Εξαμήνου - {selected_semester}")
    st.dataframe(df_sem)    


# with tab_calendar:
st.subheader("Ημερολόγιο Εξετάσεων")
calendar_options = {
    "initialView": "dayGridMonth",
    "selectable": True,
    "weekends": False,
    "headerToolbar": {
        "left": "today prev,next",
        "center": "title",
        "right": "dayGridMonth,timeGridWeek,timeGridDay"
    }
}

# Convert exam data to calendar events
calendar_events = []
for _, row in df.iterrows():
    if pd.notna(row["start_dt"]):
        # Format as string YYYY-MM-DDTHH:MM:SS
        start_str = row["start_dt"].strftime("%Y-%m-%dT%H:%M:%S")
        
        # Calculate end time (2 hours after start)
        end_dt = row["start_dt"] + timedelta(hours=2)
        end_str = end_dt.strftime("%Y-%m-%dT%H:%M:%S")
        
        # Safely handle potential None values, convert to string, and remove problematic characters
        def clean_text(value):
            if pd.notna(value):
                # Convert to string and remove newlines, quotes, backslashes
                text = str(value).replace('\n', ' ').replace('\r', ' ')
                text = text.replace('"', '').replace("'", '').replace('\\', '')
                return text.strip()
            return ""
        
        course_name = clean_text(row['course_name'])
        instructor = clean_text(row['instructor'])
        room = clean_text(row['room'])
        semester = str(int(row['semester'])) if pd.notna(row['semester']) else ""
        
        event = {
            "title": f'Εξ.{semester} - {course_name}',
            "start": start_str,
            "end": end_str
        }
        calendar_events.append(event)

# Debug: show number of events
st.write(f"📅 Total events: {len(calendar_events)}")

calendar_data = calendar(
    events=calendar_events,
    options=calendar_options,
    key="my_calender"
)

# st.write("Calendar interaction information:", calendar_data)