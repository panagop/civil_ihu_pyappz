import streamlit as st
import pandas as pd
import numpy as np
import io
from streamlit_calendar import calendar

st.set_page_config(
    layout="wide",
)

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
    
    # Convert the date column if it exists
    if 'exams_date' in df.columns:
        # Handle format like "8 Sep 15:00" - add current year
        current_year = pd.Timestamp.now().year
        
        # First, try to add year and parse
        try:
            # Add year to the date string (assuming current year)
            df['exams_date'] = df['exams_date'].astype(str).apply(
                lambda x: f"{x} {current_year}" 
                if pd.notna(x) and str(x) != 'nan' else x
            )
            # Parse with the format: "8 Sep 15:00 2024"
            df['exams_date'] = pd.to_datetime(
                df['exams_date'],
                format='%d %b %H:%M %Y',
                errors='coerce'
            )
        except Exception:
            # Fallback: try other common formats
            try:
                df['exams_date'] = pd.to_datetime(
                    df['exams_date'], errors='coerce')
            except Exception:
                st.warning("Could not parse exam dates during loading")
    
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


def create_calendar_view(df):
    """Create a calendar-like view of exam dates"""
    # Convert exams_date to datetime if it's not already
    df = df.copy()
    
    # Try to convert dates with error handling
    try:
        df['exams_date'] = pd.to_datetime(df['exams_date'], errors='coerce')
    except Exception:
        # If conversion fails, try different formats
        try:
            df['exams_date'] = pd.to_datetime(df['exams_date'], format='%d/%m/%Y', errors='coerce')
        except Exception:
            try:
                df['exams_date'] = pd.to_datetime(df['exams_date'], format='%Y-%m-%d', errors='coerce')
            except Exception:
                # If all else fails, keep as string and show warning
                st.warning("Could not parse exam dates. Showing as text.")
                df['exams_date_parsed'] = df['exams_date'].astype(str)
                return df.groupby('exams_date').apply(
                    lambda x: pd.Series({
                        'courses': ', '.join(x['course_name'].astype(str)),
                        'teachers': ', '.join(x['teacher'].astype(str)),
                        'total_students': x['number_of_students'].sum(),
                        'course_count': len(x)
                    })
                ).reset_index()
    
    # Remove rows where date conversion failed
    df = df.dropna(subset=['exams_date'])
    
    if df.empty:
        st.error("No valid exam dates found after date parsing.")
        return pd.DataFrame()
    
    # Group by date to show all exams on each day
    calendar_data = df.groupby('exams_date').apply(
        lambda x: pd.Series({
            'courses': ', '.join(x['course_name'].astype(str)),
            'teachers': ', '.join(x['teacher'].astype(str)),
            'total_students': x['number_of_students'].sum(),
            'course_count': len(x)
        })
    ).reset_index()
    
    # Sort by date
    calendar_data = calendar_data.sort_values('exams_date')
    
    return calendar_data


def create_calendar_events(df):
    """Convert exam data to calendar events format"""
    df = df.copy()
    
    # Ensure dates are datetime
    if not pd.api.types.is_datetime64_any_dtype(df['exams_date']):
        df['exams_date'] = pd.to_datetime(df['exams_date'], errors='coerce')
    
    # Remove rows with invalid dates
    df = df.dropna(subset=['exams_date'])
    
    if df.empty:
        return []
    
    events = []
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
              '#DDA0DD', '#98D8C8', '#F7DC6F', '#BB8FCE', '#85C1E9']
    
    for idx, row in df.iterrows():
        start_datetime = row['exams_date']
        # Add 2 hours for exam duration
        end_datetime = start_datetime + pd.Timedelta(hours=2)
        
        # Create event
        event = {
            'title': f"{row['course_name']} - {row['teacher']}",
            'start': start_datetime.isoformat(),
            'end': end_datetime.isoformat(),
            'color': colors[idx % len(colors)],
            'extendedProps': {
                'course_id': str(row.get('course_id', '')),
                'semester': str(row.get('semester', '')),
                'students': str(row.get('number_of_students', ''))
            }
        }
        events.append(event)
    
    return events


# Load data
df_september = get_data()

st.sidebar.button('Ενημέρωση από Google Sheets', on_click=reload)

# Create tabs for different views
tabs = st.tabs([
    "Ακατέργαστα δεδομένα",
    "Απλός πίνακας",
    "Πίνακας συγκέντρωσης",
    "Ημερολογιακή προβολή",
    "Διαδραστικό ημερολόγιο"
])
(tab_raw, tab_pivot_simple, tab_pivot_multi,
 tab_calendar, tab_interactive_calendar) = tabs

with tab_raw:
    st.markdown("### Ακατέργαστα δεδομένα")
    st.dataframe(df_september)

# Check if we have the expected columns
expected_columns = ['course_id', 'course_name', 'semester', 'teacher',
                    'number_of_students', 'exams_date']
if all(col in df_september.columns for col in expected_columns):
    
    pivot_simple, pivot_multi = create_pivot_table(df_september)
    calendar_data = create_calendar_view(df_september)
    
    with tab_pivot_simple:
        st.markdown("### Απλός πίνακας (Πιο ευανάγνωστος)")
        st.markdown("Ομαδοποίηση ανά εξάμηνο και μάθημα")
        st.dataframe(pivot_simple)
    
    with tab_pivot_multi:
        st.markdown("### Πίνακας συγκέντρωσης (Multi-index)")
        st.markdown("Παραδοσιακός pivot table με πολλαπλούς δείκτες")
        st.dataframe(pivot_multi)
    
    with tab_calendar:
        st.markdown("### Ημερολογιακή προβολή εξετάσεων")
        st.markdown("Προβολή εξετάσεων ανά ημερομηνία")
        
        # Show calendar data
        st.dataframe(
            calendar_data,
            column_config={
                "exams_date": st.column_config.DateColumn(
                    "Ημερομηνία εξέτασης",
                    format="DD/MM/YYYY"
                ),
                "courses": "Μαθήματα",
                "teachers": "Καθηγητές",
                "total_students": st.column_config.NumberColumn(
                    "Σύνολο φοιτητών",
                    format="%d"
                ),
                "course_count": st.column_config.NumberColumn(
                    "Αριθμός μαθημάτων",
                    format="%d"
                )
            },
            hide_index=True
        )
        
        # Timeline/Gantt-like visualization
        st.markdown("#### Χρονολογική προβολή")
        
        # Create a simple timeline chart
        chart_data = df_september.copy()
        
        # Try to convert dates for chart
        try:
            chart_data['exams_date'] = pd.to_datetime(chart_data['exams_date'], errors='coerce')
            chart_data = chart_data.dropna(subset=['exams_date'])
            chart_data = chart_data.sort_values('exams_date')
            
            if not chart_data.empty:
                # Create a bar chart showing number of students per exam date
                date_summary = chart_data.groupby('exams_date').agg({
                    'number_of_students': 'sum',
                    'course_name': 'count'
                }).reset_index()
                date_summary.columns = ['Ημερομηνία', 'Σύνολο φοιτητών',
                                        'Αριθμός εξετάσεων']
                
                st.bar_chart(
                    date_summary.set_index('Ημερομηνία')['Σύνολο φοιτητών'],
                    use_container_width=True
                )
            else:
                st.warning("Δεν υπάρχουν έγκυρες ημερομηνίες για το γράφημα")
        except Exception as e:
            st.error(f"Σφάλμα στη δημιουργία γραφήματος: {e}")
        
        st.markdown("#### Λεπτομέρειες ανά ημερομηνία")
        
        # Allow users to select a specific date to see details
        unique_dates = sorted(df_september['exams_date'].dropna().unique())
        if unique_dates:
            selected_date = st.selectbox(
                "Επιλέξτε ημερομηνία για λεπτομέρειες:",
                unique_dates
            )
            
            date_details = df_september[
                df_september['exams_date'] == selected_date
            ][['course_name', 'teacher', 'number_of_students', 'semester']]
            
            if not date_details.empty:
                st.markdown(f"**Εξετάσεις στις {selected_date}:**")
                st.dataframe(date_details, hide_index=True)
            else:
                st.info("Δεν υπάρχουν εξετάσεις για την επιλεγμένη ημερομηνία")
        else:
            st.info("Δεν υπάρχουν διαθέσιμες ημερομηνίες εξετάσεων")
    
    with tab_interactive_calendar:
        st.markdown("### 📅 Διαδραστικό ημερολόγιο εξετάσεων")
        st.markdown("Ημερολόγιο με πλήρη διαδραστικότητα (διάρκεια εξετάσεων: 2 ώρες)")
        
        # Create calendar events
        events = create_calendar_events(df_september)
        
        if events:
            # Calendar options
            calendar_options = {
                "editable": "false",
                "navLinks": "true",
                "resources": [],
                "selectable": "true",
                "initialView": "dayGridMonth",
                "height": 1000,  # Set calendar height in pixels
                # "aspectRatio": 1.5,  # Width to height ratio
                "headerToolbar": {
                    "left": "prev,next today",
                    "center": "title",
                    "right": "dayGridMonth,timeGridWeek,timeGridDay"
                }
            }
            
            # Custom CSS for better styling
            custom_css = """
                .fc-event-past {
                    opacity: 0.8;
                }
                .fc-event-time {
                    font-weight: bold;
                }
                .fc-event-title {
                    font-weight: normal;
                }
            """
            
            # Display the calendar
            calendar_component = calendar(
                events=events,
                options=calendar_options,
                custom_css=custom_css,
                key="exam_calendar"
            )
            
            # Debug calendar component state
            calendar_keys = (list(calendar_component.keys())
                             if calendar_component else "None")
            st.write("Calendar component keys:", calendar_keys)
            
            # Show event details when clicked
            if calendar_component.get('eventClick'):
                st.markdown("#### 🔍 Πληροφορίες εξέτασης")
                
                event = calendar_component['eventClick']['event']
                start_dt = pd.to_datetime(event['start'])
                end_dt = pd.to_datetime(event['end'])
                
                # Format event details
                time_range = (f"{start_dt.strftime('%H:%M')} - "
                              f"{end_dt.strftime('%H:%M')}")
                props = event.get('extendedProps', {})
                students = props.get('students', 'N/A')
                semester = props.get('semester', 'N/A')
                
                st.success(f"""
                **📚 Μάθημα:** {event['title']}
                **📅 Ημερομηνία:** {start_dt.strftime('%d/%m/%Y')}
                **🕐 Ώρα:** {time_range}
                **👥 Φοιτητές:** {students}
                **📖 Εξάμηνο:** {semester}
                """)
            else:
                st.info("👆 Κάντε κλικ σε μια εξέταση στο ημερολόγιο για "
                        "να δείτε λεπτομέρειες")
            
            # Summary statistics
            st.markdown("#### 📊 Στατιστικά εξετάσεων")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Συνολικές εξετάσεις", len(events))
            
            with col2:
                total_students = df_september['number_of_students'].sum()
                st.metric("Συνολικοί φοιτητές", total_students)
            
            with col3:
                unique_dates = df_september['exams_date'].dt.date.nunique()
                st.metric("Ημέρες εξετάσεων", unique_dates)
                
        else:
            st.warning("Δεν βρέθηκαν εξετάσεις για εμφάνιση στο ημερολόγιο")


else:
    missing_cols = [col for col in expected_columns
                    if col not in df_september.columns]
    st.error(f"Λείπουν οι εξής στήλες: {missing_cols}")
    st.info(f"Διαθέσιμες στήλες: {list(df_september.columns)}")

