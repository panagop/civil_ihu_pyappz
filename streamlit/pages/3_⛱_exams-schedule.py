import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import streamlit as st
from streamlit_calendar import calendar

st.set_page_config(
    layout="wide",
)

# ------------ ΡΥΘΜΙΣΕΙΣ ΧΡΗΣΤΗ ------------
# INPUT_EXCEL = "lessons-calendars.xlsm"   # το αρχείο όπου έχεις το sheet Data
INPUT_SHEET = "2026-01"
INPUT_EXCEL = Path(__file__).parent.parent.parent / "files" / "exams" / "exams-2026-01.xlsm"


# @st.cache_data
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

    # Add day of week column (in Greek)
    day_names_greek = {
        0: 'Δευτέρα',
        1: 'Τρίτη', 
        2: 'Τετάρτη',
        3: 'Πέμπτη',
        4: 'Παρασκευή',
        5: 'Σάββατο',
        6: 'Κυριακή'
    }
    df["day_of_week"] = pd.to_datetime(df["exam_date"]).dt.dayofweek.map(day_names_greek)

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
    st.dataframe(df_sem, height=600)    


with tab_calendar:
    st.subheader("Ημερολόγιο Εξετάσεων")

    # Add semester filter for calendar
    semesters_all = sorted(df["semester"].unique().tolist())
    semester_options = [f"Εξάμηνο {int(s)}" for s in semesters_all]

    selected_calendar_semesters = st.multiselect(
        "Φιλτράρισμα ημερολογίου κατά εξάμηνο:",
        options=semester_options,
        default=semester_options,  # All semesters selected by default
        key="calendar_semester_filter"
    )

    # Filter dataframe based on selection
    if not selected_calendar_semesters or len(selected_calendar_semesters) == len(semester_options):
        df_calendar = df
    else:
        # Extract semester numbers from "Εξάμηνο X"
        semester_nums = [int(s.split()[-1]) for s in selected_calendar_semesters]
        df_calendar = df[df["semester"].isin(semester_nums)]

    # Get the earliest exam date for initial calendar view
    initial_date = df_calendar["exam_date"].min() if not df_calendar.empty else datetime.now().date()

    calendar_options = {
        "initialView": "dayGridMonth",
        "initialDate": initial_date.strftime("%Y-%m-%d"),
        "selectable": True,
        "weekends": False,
        "slotMinTime": "08:00:00",
        "slotMaxTime": "22:00:00",
        "headerToolbar": {
            "left": "today prev,next",
            "center": "title",
            "right": "dayGridMonth,timeGridWeek,timeGridDay"
        }
    }

    # Convert exam data to calendar events
    calendar_events = []
    for _, row in df_calendar.iterrows():
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
                "title": f'Εξ.{semester} - {course_name} - {instructor}',
                "start": start_str,
                "end": end_str
            }
            calendar_events.append(event)

    # Debug: show number of events
    if not selected_calendar_semesters or len(selected_calendar_semesters) == len(semester_options):
        st.write(f"📅 Σύνολο εξετάσεων: {len(calendar_events)} (όλα τα εξάμηνα)")
    else:
        semesters_text = ", ".join(selected_calendar_semesters)
        st.write(f"📅 Σύνολο εξετάσεων: {len(calendar_events)} ({semesters_text})")

    # Create a unique key based on selected semesters to force calendar re-render
    calendar_key = f"calendar_{'_'.join(sorted(selected_calendar_semesters)) if selected_calendar_semesters else 'all'}"

    calendar_data = calendar(
        events=calendar_events,
        options=calendar_options,
        key=calendar_key
    )

# st.write("Calendar interaction information:", calendar_data)