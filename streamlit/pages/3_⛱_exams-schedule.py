from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from streamlit_calendar import calendar

from lib.exams_data import load_data
from lib.exams_export import create_weekly_calendar_document

st.set_page_config(
    layout="wide",
    page_title="Πρόγραμμα Εξετάσεων",
    page_icon="🗓️",
)

period_selection: str = st.radio(
    "Επιλέξτε εξάμηνο:",
    options=["Χειμερινό", "Εαρινό"],
    index=1,
    key="exams_period_selection",
)

exam_period = f"{period_selection} Εξάμηνο 2025-2026"

st.title(f"🗓️ Πρόγραμμα Εξετάσεων - {exam_period}")

program_selection: str = st.radio(
    "Επιλογή προγράμματος σπουδών:",
    options=["ΔΙΠΑΕ", "ΤΕΙ"],
    index=0,
    key="program_selection",
)

st.markdown(f"Έχετε επιλέξει το πρόγραμμα σπουδών: **{program_selection}**")

INPUT_SHEET = program_selection
INPUT_EXCEL = Path(__file__).parent.parent.parent / "files" / "exams" / "exams-2026-06.xlsm"


tab_full_table, tab_instructor_filter, tab_semester_filter, tab_epitiritis_filter, tab_calendar, tab_export_weekly = st.tabs(
    [
        "Πλήρης Πίνακας Εξετάσεων",
        "Φιλτράρισμα κατά Διδάσκοντα",
        "Φιλτράρισμα κατά Εξάμηνο",
        "Φιλτράρισμα κατά Επιτηρητή",
        "Ημερολόγιο Εξετάσεων",
        "Εξαγωγή Εβδομαδιαίου Προγράμματος",
    ]
)

extra_column = "students_total" if program_selection == "ΔΙΠΑΕ" else "φοιτΤΕΙ"
df = load_data(INPUT_EXCEL, INPUT_SHEET, extra_column=extra_column)


with tab_full_table:
    st.subheader("Πλήρης Πίνακας Εξετάσεων")
    display_cols = [col for col in df.columns if col not in ['start_dt', 'iso_week_number']]
    st.dataframe(df[display_cols])

instructors = sorted(df["instructor"].unique().tolist())


with tab_instructor_filter:
    selected_instructor = st.selectbox(
        "Επιλέξτε διδάσκοντα για φιλτράρισμα:",
        options=instructors)

    df_instr = df[df["instructor"] == selected_instructor].sort_values(
        by=["start_dt"]
    )

    st.subheader(f"Πρόγραμμα Εξετάσεων Διδάσκοντα - {selected_instructor}")
    display_cols = [col for col in df_instr.columns if col not in ['start_dt', 'iso_week_number']]
    st.dataframe(df_instr[display_cols])

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
    display_cols = [col for col in df_sem.columns if col not in ['start_dt', 'iso_week_number']]
    st.dataframe(df_sem[display_cols], height=600)

with tab_epitiritis_filter:
    epitirites_list = []
    for val in df["epitirites"].dropna().unique():
        if pd.notna(val):
            epitirites_list.extend([e.strip() for e in str(val).split(',')])

    epitirites_unique = sorted(list(set(epitirites_list)))

    if epitirites_unique:
        selected_epitiritis = st.selectbox(
            "Επιλέξτε επιτηρητή για φιλτράρισμα:",
            options=epitirites_unique
        )

        df_epit = df[df["epitirites"].apply(
            lambda x: selected_epitiritis in str(x) if pd.notna(x) else False
        )].sort_values(by=["start_dt"])

        st.subheader(f"Πρόγραμμα Επιτηρήσεων - {selected_epitiritis}")

        display_cols = [col for col in df_epit.columns if col not in ['start_dt', 'iso_week_number']]
        st.dataframe(df_epit[display_cols], height=600)

        st.markdown("### Στατιστικά")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Συνολικές Επιτηρήσεις", len(df_epit))
        with col2:
            unique_weeks = df_epit['week_number'].nunique()
            st.metric("Εβδομάδες με Επιτηρήσεις", unique_weeks)
    else:
        st.warning("⚠️ Δεν βρέθηκαν δεδομένα επιτηρητών στο αρχείο.")

with tab_calendar:
    st.subheader("Ημερολόγιο Εξετάσεων")

    semesters_all = sorted(df["semester"].unique().tolist())
    semester_options = [f"Εξάμηνο {int(s)}" for s in semesters_all]

    selected_calendar_semesters = st.multiselect(
        "Φιλτράρισμα ημερολογίου κατά εξάμηνο:",
        options=semester_options,
        default=semester_options,
        key="calendar_semester_filter"
    )

    if not selected_calendar_semesters or len(selected_calendar_semesters) == len(semester_options):
        df_calendar = df
    else:
        semester_nums = [int(s.split()[-1]) for s in selected_calendar_semesters]
        df_calendar = df[df["semester"].isin(semester_nums)]

    initial_date = df_calendar["exam_date"].min() if not df_calendar.empty else datetime.now().date()

    calendar_options = {
        "initialView": "timeGridWeek",
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

    semester_colors = {
        1: '#E74C3C',
        2: '#3498DB',
        3: '#2ECC71',
        4: '#F39C12',
        5: '#9B59B6',
        6: '#1ABC9C',
        7: '#E67E22',
        8: '#34495E',
        9: '#16A085',
        10: '#D35400',
    }

    calendar_events = []
    for _, row in df_calendar.iterrows():
        if pd.notna(row["start_dt"]):
            start_str = row["start_dt"].strftime("%Y-%m-%dT%H:%M:%S")

            end_dt = row["start_dt"] + timedelta(hours=2)
            end_str = end_dt.strftime("%Y-%m-%dT%H:%M:%S")

            def clean_text(value):
                if pd.notna(value):
                    text = str(value).replace('\n', ' ').replace('\r', ' ')
                    text = text.replace('"', '').replace("'", '').replace('\\', '')
                    return text.strip()
                return ""

            course_name = clean_text(row['course_name'])
            instructor = clean_text(row['instructor'])
            room = clean_text(row['room'])
            semester = int(row['semester']) if pd.notna(row['semester']) else 1
            semester_str = str(semester)

            color = semester_colors.get(semester, '#95A5A6')

            event = {
                "title": f'Εξ.{semester_str} - {course_name} - {instructor}',
                "start": start_str,
                "end": end_str,
                "color": color
            }
            calendar_events.append(event)

    if not selected_calendar_semesters or len(selected_calendar_semesters) == len(semester_options):
        st.write(f"📅 Σύνολο εξετάσεων: {len(calendar_events)} (όλα τα εξάμηνα)")
    else:
        semesters_text = ", ".join(selected_calendar_semesters)
        st.write(f"📅 Σύνολο εξετάσεων: {len(calendar_events)} ({semesters_text})")

    if 'show_exam_calendar' not in st.session_state:
        st.session_state.show_exam_calendar = False

    if not st.session_state.show_exam_calendar:
        if st.button("📅 Εμφάνιση Ημερολογίου", key="show_exam_cal_btn"):
            st.session_state.show_exam_calendar = True
            st.rerun()

    if st.session_state.show_exam_calendar:
        if calendar_events:
            calendar_data = calendar(
                events=calendar_events,
                options=calendar_options
            )
        else:
            st.info("Δεν υπάρχουν εξετάσεις για εμφάνιση με τα επιλεγμένα φίλτρα.")


with tab_export_weekly:
    st.subheader("Εξαγωγή Εβδομαδιαίου Προγράμματος Εξετάσεων")
    st.markdown("Δημιουργήστε αρχείο Word με το εβδομαδιαίο πρόγραμμα εξετάσεων για διανομή σε συναδέλφους.")

    st.markdown("### Επιλογές Φιλτραρίσματος")

    col1, col2 = st.columns(2)

    with col1:
        semesters_export = sorted(df["semester"].unique().tolist())
        semester_options_export = [f"Εξάμηνο {int(s)}" for s in semesters_export]

        selected_export_semesters = st.multiselect(
            "Επιλέξτε εξάμηνα:",
            options=semester_options_export,
            default=semester_options_export,
            key="export_semester_filter"
        )

    with col2:
        weeks_available = sorted(df['week_number'].unique().tolist())
        week_options = [f"Εβδομάδα {int(w)}" for w in weeks_available]

        selected_export_weeks = st.multiselect(
            "Επιλέξτε εβδομάδες:",
            options=week_options,
            default=week_options,
            key="export_week_filter"
        )

    include_epitirites = st.checkbox(
        "Συμπερίληψη επιτηρητών στο αρχείο Word",
        value=True,
        key="include_epitirites_checkbox",
        help="Επιλέξτε αν θέλετε να περιλαμβάνονται οι επιτηρητές στο εξαγόμενο αρχείο"
    )

    df_export = df.copy()

    if selected_export_semesters and len(selected_export_semesters) < len(semester_options_export):
        semester_nums = [int(s.split()[-1]) for s in selected_export_semesters]
        df_export = df_export[df_export["semester"].isin(semester_nums)]

    if selected_export_weeks and len(selected_export_weeks) < len(week_options):
        week_nums = [int(w.split()[-1]) for w in selected_export_weeks]
        df_export = df_export[df_export["week_number"].isin(week_nums)]

    st.markdown("### Προεπισκόπηση Δεδομένων")
    st.write(f"Σύνολο εξετάσεων προς εξαγωγή: {len(df_export)}")

    if not df_export.empty:
        st.dataframe(
            df_export[['exam_date', 'day_of_week', 'start_time', 'semester',
                       'course_name', 'instructor', 'room', 'epitirites']].sort_values(by=['exam_date', 'start_time']),
            height=400
        )

        st.markdown("### Λήψη Αρχείου")

        try:
            word_file = create_weekly_calendar_document(
                df_export,
                period=period_selection,
                include_epitirites=include_epitirites,
            )

            filename = f"Πρόγραμμα_Εξετάσεων_{program_selection}_{period_selection}_2025-2026.docx"

            st.download_button(
                label="📥 Λήψη Word Αρχείου",
                data=word_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                help="Κατεβάστε το εβδομαδιαίο πρόγραμμα εξετάσεων σε μορφή Word"
            )

            st.success("✅ Το αρχείο είναι έτοιμο για λήψη!")

        except Exception as e:
            st.error(f"Σφάλμα κατά τη δημιουργία του αρχείου: {e}")
    else:
        st.warning("⚠️ Δεν υπάρχουν δεδομένα με τα επιλεγμένα φίλτρα.")
