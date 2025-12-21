import pandas as pd
from pathlib import Path
import streamlit as st
from datetime import datetime, timedelta
from streamlit_calendar import calendar

st.set_page_config(
    layout="wide",
    page_title="Εβδομαδιαίο Πρόγραμμα Μαθημάτων",
    page_icon="📅",
)

st.title("📅 Εβδομαδιαίο Πρόγραμμα Μαθημάτων")

# Επιλογή εξαμήνου
period_selection = st.radio(
    "Επιλέξτε εξάμηνο:",
    options=["Χειμερινό", "Εαρινό"],
    index=0,
    key="period_selection"
)

st.markdown(f"Έχετε επιλέξει: **{period_selection} Εξάμηνο**")

# Ρυθμίσεις αρχείου
INPUT_EXCEL = Path(__file__).parent.parent.parent / "files" / "timetables" / "2025-2026.xlsm"


def load_data() -> pd.DataFrame:
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
        df = df[df['teaching_period'] == period_selection]
        
        # Δημιουργία συνδυαστικής στήλης για καλύτερη αναγνώριση
        df['full_class_name'] = df.apply(
            lambda row: f"{row['course_name']} - {row['class_name']}" 
            if pd.notna(row['class_name']) else str(row['course_name']), 
            axis=1
        )
        
        # Μετατροπή start_time σε ώρα (integer) για υπολογισμούς
        df['start_hour'] = df['start_time'].apply(
            lambda x: x.hour if hasattr(x, 'hour') else int(x)
        )
        
        # Υπολογισμός ώρας λήξης και δημιουργία end_time
        df['end_hour'] = df['start_hour'] + df['duration']
        df['end_time'] = df.apply(
            lambda row: f"{int(row['end_hour'])}:00", 
            axis=1
        )
            
    except Exception as e:
        st.error(f"❌ Σφάλμα κατά το άνοιγμα του αρχείου: {e}")
        st.stop()
    
    return df


# Φόρτωση δεδομένων
try:
    df = load_data()
    
    # st.subheader(f"Πρόγραμμα {semester_selection} Εξαμήνου 2025-2026")
    
    # Tabs
    tab_table, tab_calendar = st.tabs(["Πίνακας", "Εβδομαδιαία Προβολή"])
    
    with tab_table:
        # Εμφάνιση δεδομένων (με end_time)
        display_cols = ['course_id', 'course_name', 'class_name', 'full_class_name', 'semester', 
                       'teaching_period', 'instructors', 'day', 'start_time', 'end_time', 
                       'duration', 'room', 'notes']
        available_cols = [col for col in display_cols if col in df.columns]
        st.dataframe(df[available_cols], use_container_width=True)
    
    with tab_calendar:
        st.markdown("### Εβδομαδιαίο Πρόγραμμα")
        
        # Φίλτρο εξαμήνων σπουδών
        semesters_all = sorted(df["semester"].unique().tolist())
        semester_options = [f"Εξάμηνο {int(s)}" for s in semesters_all]
        
        selected_semester = st.selectbox(
            "Επιλέξτε εξάμηνο σπουδών:",
            options=semester_options,
            index=0,
            key="semester_filter"
        )
        
        # Φιλτράρισμα δεδομένων
        semester_num = int(selected_semester.split()[-1])
        df_filtered = df[df["semester"] == semester_num]
        
        # Χρώματα ανά εξάμηνο
        semester_colors = {
            1: '#E74C3C',  2: '#E74C3C',  3: '#E74C3C',  4: '#E74C3C',  5: '#E74C3C',
            6: '#E74C3C',  7: '#E74C3C',  8: '#E74C3C',  9: '#E74C3C',  10: '#E74C3C',
        }
        # semester_colors = {
        #     1: '#E74C3C',  2: '#3498DB',  3: '#2ECC71',  4: '#F39C12',  5: '#9B59B6',
        #     6: '#1ABC9C',  7: '#E67E22',  8: '#34495E',  9: '#16A085',  10: '#D35400',
        # }
        
        # Safely handle potential None values, convert to string, and remove problematic characters
        def clean_text(value):
            if pd.notna(value):
                # Convert to string and remove newlines, quotes, backslashes
                text = str(value).replace('\n', ' ').replace('\r', ' ')
                text = text.replace('"', '').replace("'", '').replace('\\', '')
                return text.strip()
            return ""
        
        # Convert to calendar events
        calendar_events = []
        
        # Map Greek days to weekday numbers (0=Monday)
        day_map = {
            'Δευτέρα': 0, 'Τρίτη': 1, 'Τετάρτη': 2, 'Πέμπτη': 3, 'Παρασκευή': 4,
            'Σάββατο': 5, 'Κυριακή': 6
        }
        
        # Use a reference week (e.g., a week in January 2025)
        reference_date = datetime(2025, 1, 6)  # Monday, January 6, 2025
        
        for _, row in df_filtered.iterrows():
            try:
                if pd.notna(row['day']) and pd.notna(row['start_time']):
                    # Get the day of week
                    day_name = str(row['day']).strip()
                    weekday = day_map.get(day_name, 0)
                    
                    # Calculate the date for this event
                    event_date = reference_date + timedelta(days=weekday)
                    
                    # Extract hour from start_time
                    start_time_str = str(row['start_time'])
                    if ':' in start_time_str:
                        start_hour = int(start_time_str.split(':')[0])
                    else:
                        start_hour = int(float(start_time_str))
                    
                    # Create start datetime
                    start_dt = event_date.replace(hour=start_hour, minute=0, second=0)
                    start_str = start_dt.strftime("%Y-%m-%dT%H:%M:%S")
                    
                    # Calculate end time based on duration
                    duration = int(row['duration']) if pd.notna(row['duration']) else 1
                    end_dt = start_dt + timedelta(hours=duration)
                    end_str = end_dt.strftime("%Y-%m-%dT%H:%M:%S")
                    
                    # Clean text
                    full_class_name = clean_text(row['full_class_name'])
                    instructors = clean_text(row['instructors'])
                    room = clean_text(row['room'])
                    semester = int(row['semester']) if pd.notna(row['semester']) else 1
                    
                    # Get color
                    color = semester_colors.get(semester, '#95A5A6')
                    
                    # Create concise title
                    title_parts = [full_class_name]
                    if instructors:
                        title_parts.append(instructors)
                    if room:
                        title_parts.append(f'({room})')
                    
                    event = {
                        "title": ' - '.join(title_parts),
                        "start": start_str,
                        "end": end_str,
                        "color": color
                    }
                    calendar_events.append(event)
            except Exception as e:
                # Skip rows with errors
                continue
        
        st.write(f"📚 Σύνολο μαθημάτων: {len(calendar_events)}")
        
        
        # CSS to hide dates and make it look generic
        st.markdown("""
        <style>
        /* Hide the date numbers in column headers */
        .fc-col-header-cell-cushion {
            font-size: 14px !important;
        }
        .fc-daygrid-day-number {
            display: none !important;
        }
        /* Hide the full date range in title */
        .fc-toolbar-title {
            display: none !important;
        }
        /* Style for cleaner look */
        .fc-toolbar-chunk:first-child {
            display: none !important;
        }
        /* Preserve whitespace and line breaks in event titles */
        .fc-event-title, .fc-event-title-container, .fc-timegrid-event-harness, .fc-event-main {
            white-space: pre-line !important;
        }
        .fc-timegrid-event {
            white-space: pre-line !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Calendar options
        calendar_options = {
            "initialView": "timeGridWeek",
            "initialDate": "2025-01-06",  # Start on Monday
            "headerToolbar": {
                "left": "",
                "center": "", 
                "right": ""
            },
            "slotMinTime": "08:00:00",
            "slotMaxTime": "21:00:00",
            "allDaySlot": False,
            "height": 850,
            "locale": "el",
            "firstDay": 1,  # Monday
            "weekends": False,  # Hide weekends
            "navLinks": False,
            "editable": False,
            "selectable": False,
            "dayHeaderFormat": {"weekday": "long"},  # Show only day names
            "displayEventTime": False,  # Hide time in event boxes
        }
        
        # Create a unique key based on selected semester
        semester_num = int(selected_semester.split()[-1])
        calendar_key = f"timetable_sem_{semester_num}"
        
        # Initialize session state for calendar display
        if 'show_timetable_calendar' not in st.session_state:
            st.session_state.show_timetable_calendar = False

        # Button to show calendar
        if not st.session_state.show_timetable_calendar:
            if st.button("📅 Εμφάνιση Ημερολογίου", key="show_timetable_cal_btn"):
                st.session_state.show_timetable_calendar = True
                st.rerun()
        
        # Render calendar if button was clicked
        if st.session_state.show_timetable_calendar:
            if calendar_events:
                calendar_data = calendar(
                    events=calendar_events,
                    options=calendar_options,
                    key=calendar_key
                )
            else:
                st.info("Δεν υπάρχουν μαθήματα για εμφάνιση με τα επιλεγμένα φίλτρα.")
        
except Exception as e:
    st.error(f"Σφάλμα: {e}")
    st.info("Παρακαλώ ελέγξτε τη δομή του αρχείου Excel και τα ονόματα των sheets.")


