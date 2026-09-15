"""Εβδομαδιαίο πρόγραμμα v2 — the timetable from Postgres, and its preparation.

Page 4 reads ``files/timetables/<year>.xlsm`` and stays as it is. This page
shows the same five views from the database (``timetable_db``) and adds three
tabs: **Προετοιμασία**, where a coordinator opens the next term as a copy of
the same period one year earlier, places and moves classes and locks it;
**Προσωπικό**, the staff list with a per-term active flag; **Αίθουσες**.

Viewing is public, like page 4 — a timetable carries no personal data. Editing
is for ``coordinator_emails`` only (decided 2026-09-15), so the page does not
gate itself with ``require_ihu_login``; it checks the role where it matters.
"""

import io
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from streamlit_calendar import calendar

st.set_page_config(
    layout="wide",
    page_title="Εβδομαδιαίο πρόγραμμα v2",
    page_icon="🗓",
)

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import db  # noqa: E402
import timetable_db as tdb  # noqa: E402
from auth import is_authorized  # noqa: E402
from branding import apply_branding  # noqa: E402
from utils.colors import DEFAULT_SEMESTER_COLOR, SEMESTER_COLORS  # noqa: E402
from utils.timetable_export import create_weekly_timetable_document  # noqa: E402

apply_branding()
db.bootstrap()

st.title("🗓 Εβδομαδιαίο πρόγραμμα μαθημάτων")

if not db.is_available():
    st.error(
        "Η σελίδα λειτουργεί μόνο με σύνδεση στη βάση δεδομένων, "
        "η οποία είναι προσβάσιμη μόνο μέσα από το Railway. "
        "Το πρόγραμμα από το αρχείο Excel βρίσκεται στη σελίδα «Εβδομαδιαίο Πρόγραμμα»."
    )
    st.stop()

terms = tdb.stored_terms()
if not terms:
    st.warning("Δεν υπάρχουν αποθηκευμένα προγράμματα.")
    st.stop()

user_email = getattr(st.user, "email", "") or ""
coordinator = is_authorized() and db.is_coordinator(user_email)
open_list = tdb.open_terms()
open_term = open_list[0] if open_list else None


def term_option(term: tuple[int, str]) -> str:
    label = tdb.term_label(*term)
    return f"{label} (σε προετοιμασία)" if term == open_term else label


default_index = terms.index(open_term) if open_term in terms else 0
year, period = st.selectbox(
    "Εξάμηνο:", options=terms, index=default_index, format_func=term_option, key="term"
)
editable = (year, period) == open_term
df = tdb.load_term(year, period)
placed = df[df["placed"]]

if editable:
    st.info(
        "Το εξάμηνο αυτό είναι **σε προετοιμασία**: ό,τι φαίνεται εδώ μπορεί ακόμη να αλλάξει."
    )

# --------------------------------------------------------------------------
# Calendar helpers (the same FullCalendar setup page 4 used, once)
# --------------------------------------------------------------------------

CALENDAR_CSS = """
<style>
.fc-col-header-cell-cushion { font-size: 14px !important; }
.fc-daygrid-day-number, .fc-toolbar-title, .fc-toolbar-chunk:first-child { display: none !important; }
.fc-event-title, .fc-event-title-container, .fc-timegrid-event-harness, .fc-event-main,
.fc-timegrid-event { white-space: pre-line !important; }
</style>
"""
REFERENCE_MONDAY = datetime(2025, 1, 6)  # noqa: DTZ001 - a Monday, any Monday
CALENDAR_OPTIONS = {
    "initialView": "timeGridWeek",
    "initialDate": REFERENCE_MONDAY.strftime("%Y-%m-%d"),
    "headerToolbar": {"left": "", "center": "", "right": ""},
    "slotMinTime": f"{tdb.FIRST_HOUR:02d}:00:00",
    "slotMaxTime": f"{tdb.LAST_HOUR:02d}:00:00",
    "allDaySlot": False,
    "height": 850,
    "locale": "el",
    "firstDay": 1,
    "weekends": False,
    "navLinks": False,
    "editable": False,
    "selectable": False,
    "dayHeaderFormat": {"weekday": "long"},
    "displayEventTime": False,
}


def week_events(frame: pd.DataFrame, title) -> list[dict]:
    events = []
    for _, row in frame[frame["placed"]].iterrows():
        start = REFERENCE_MONDAY + timedelta(days=int(row["day_number"]) - 1, hours=int(row["start_hour"]))
        end = start + timedelta(hours=int(row["duration"]))
        events.append(
            {
                "title": title(row),
                "start": start.strftime("%Y-%m-%dT%H:%M:%S"),
                "end": end.strftime("%Y-%m-%dT%H:%M:%S"),
                "color": SEMESTER_COLORS.get(int(row["examino"]), DEFAULT_SEMESTER_COLOR),
            }
        )
    return events


def render_week(frame: pd.DataFrame, key: str, title) -> None:
    events = week_events(frame, title)
    if not events:
        st.info("Δεν υπάρχουν μαθήματα με ώρα για εμφάνιση.")
        return
    st.markdown(CALENDAR_CSS, unsafe_allow_html=True)
    calendar(events=events, options=CALENDAR_OPTIONS, key=key)


def title_with_room(row) -> str:
    parts = [row["full_class_name"]]
    if row["instructors"]:
        parts.append(row["instructors"])
    if row["room"]:
        parts.append(f"({row['room']})")
    return " - ".join(parts)


def title_with_semester(row) -> str:
    parts = [row["full_class_name"]]
    if row["instructors"]:
        parts.append(row["instructors"])
    parts.append(f"Εξ.{int(row['examino'])}")
    return " - ".join(parts)


DISPLAY_COLUMNS = {
    "id": "id",
    "examino": "Εξάμηνο",
    "course_code": "Κωδικός",
    "display_name": "Μάθημα",
    "section": "Τμήμα",
    "instructors": "Διδάσκοντες",
    "day": "Ημέρα",
    "start_time": "Έναρξη",
    "end_time": "Λήξη",
    "duration": "Ώρες",
    "room": "Αίθουσα",
    "notes": "Παρατηρήσεις",
}


def to_display(frame: pd.DataFrame) -> pd.DataFrame:
    return frame[list(DISPLAY_COLUMNS)].rename(columns=DISPLAY_COLUMNS)


def to_excel(frame: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        frame.to_excel(writer, index=False)
    return buffer.getvalue()


def class_label(row) -> str:
    when = f"{row['day']} {row['start_time']}" if row["placed"] else "χωρίς ώρα"
    return f"{row['course_code']} {row['section']} · {row['course_name']} · {when}"


# --------------------------------------------------------------------------
# Tabs
# --------------------------------------------------------------------------

(
    tab_calendar,
    tab_table,
    tab_rooms,
    tab_instructors,
    tab_export,
    tab_prepare,
    tab_staff,
    tab_room_list,
) = st.tabs(
    [
        "Εβδομαδιαία προβολή",
        "Πίνακας",
        "Αιθουσιολόγιο",
        "Ανά διδάσκοντα",
        "Εξαγωγή Word",
        "Προετοιμασία",
        "Προσωπικό",
        "Αίθουσες",
    ]
)

with tab_calendar:
    semesters = sorted(int(s) for s in df["examino"].unique())
    if not semesters:
        st.info("Το εξάμηνο δεν έχει μαθήματα.")
    else:
        semester = st.selectbox(
            "Εξάμηνο σπουδών:",
            options=semesters,
            format_func=lambda s: f"Εξάμηνο {s}",
            key="calendar_semester",
        )
        subset = df[df["examino"] == semester]
        unplaced = subset[~subset["placed"]]
        st.write(f"📚 Μαθήματα με ώρα: {int(subset['placed'].sum())}")
        if not unplaced.empty:
            st.caption(
                "Χωρίς ώρα: " + ", ".join(
                    f"{r['course_code']} {r['section']}" for _, r in unplaced.iterrows()
                )
            )
        render_week(subset, f"week_{year}_{period}_{semester}", title_with_room)

with tab_table:
    st.dataframe(to_display(df), width="stretch", hide_index=True)
    st.download_button(
        "Λήψη Excel",
        data=to_excel(to_display(df)),
        file_name=f"programma_{tdb.year_label(year)}_{period}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="table_xlsx",
    )

with tab_rooms:
    st.markdown("### Αιθουσιολόγιο")
    room_codes = sorted({code for codes in placed["room_codes"] for code in codes})
    if not room_codes:
        st.warning("Δεν υπάρχουν αίθουσες στο πρόγραμμα.")
    else:
        rooms_frame = tdb.load_rooms()
        room_names = dict(zip(rooms_frame["code"], rooms_frame["name"]))
        room = st.selectbox(
            "Αίθουσα:",
            options=room_codes,
            format_func=lambda c: f"{c} — {room_names.get(c, '')}",
            key="room_filter",
        )
        subset = placed[[room in codes for codes in placed["room_codes"]]]
        st.write(f"🏫 Μαθήματα στην αίθουσα {room}: {len(subset)}")
        render_week(subset, f"room_{year}_{period}_{room}", title_with_semester)

with tab_instructors:
    st.markdown("### Μαθήματα ανά διδάσκοντα")
    exploded = df[df["instructors"] != ""].copy()
    exploded["Διδάσκων"] = exploded["instructors"].str.split(", ")
    exploded = exploded.explode("Διδάσκων")
    if exploded.empty:
        st.warning("Δεν βρέθηκαν διδάσκοντες.")
    else:
        names = sorted(exploded["Διδάσκων"].unique())
        col1, col2 = st.columns([2, 1])
        with col1:
            who = st.selectbox("Διδάσκων:", options=["Όλοι"] + names, key="instructor_filter")
        with col2:
            if who == "Όλοι":
                st.metric("Διδάσκοντες", len(names))
                st.metric("Γραμμές προγράμματος", len(exploded))
            else:
                mine = exploded[exploded["Διδάσκων"] == who]
                st.metric("Γραμμές", len(mine))
                st.metric("Ώρες / εβδομάδα", int(mine["duration"].fillna(0).sum()))
        columns = ["examino", "course_code", "display_name", "section", "day", "start_time",
                   "duration", "room", "notes"]
        if who == "Όλοι":
            for name in names:
                mine = exploded[exploded["Διδάσκων"] == name]
                hours = int(mine["duration"].fillna(0).sum())
                with st.expander(f"📚 {name} ({hours} ώρες)"):
                    st.dataframe(
                        mine[columns].rename(columns=DISPLAY_COLUMNS), width="stretch", hide_index=True
                    )
        else:
            mine = exploded[exploded["Διδάσκων"] == who]
            st.dataframe(mine[columns].rename(columns=DISPLAY_COLUMNS), width="stretch", hide_index=True)
            render_week(mine, f"staff_{year}_{period}_{who}", title_with_room)

with tab_export:
    st.subheader("Εξαγωγή εβδομαδιαίου προγράμματος σε Word")
    semesters = sorted(int(s) for s in placed["examino"].unique())
    chosen = st.multiselect(
        "Εξάμηνα:",
        options=semesters,
        default=semesters,
        format_func=lambda s: f"Εξάμηνο {s}",
        key="export_semesters",
    )
    export = placed[placed["examino"].isin(chosen)]
    st.write(f"Σύνολο γραμμών προς εξαγωγή: {len(export)}")
    if export.empty:
        st.warning("Δεν υπάρχουν δεδομένα με τα επιλεγμένα φίλτρα.")
    else:
        # The export reads the workbook's column names; load_term provides them.
        word_file = create_weekly_timetable_document(export, period, tdb.year_label(year))
        st.download_button(
            "📥 Λήψη αρχείου Word",
            data=word_file,
            file_name=f"Προγραμμα_Μαθηματων_{period}_{tdb.year_label(year)}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="word_download",
        )

# --------------------------------------------------------------------------
# Προετοιμασία
# --------------------------------------------------------------------------

with tab_prepare:
    state = tdb.term_state(year, period) or {}
    st.markdown(f"### {tdb.term_label(year, period)} — {state.get('status', '')}")
    if state.get("baseline_year"):
        st.caption(
            f"Άνοιξε από {state['opened_by']} ως αντίγραφο του "
            f"{tdb.term_label(state['baseline_year'], state['baseline_period'])}."
        )

    problems = tdb.conflicts(df)
    st.markdown("#### Συγκρούσεις")
    if problems.empty:
        st.success("Καμία σύγκρουση αίθουσας, διδάσκοντα ή εξαμήνου.")
    else:
        st.warning(f"{len(problems)} συγκρούσεις.")
        st.dataframe(problems.drop(columns=["id_a", "id_b"]), width="stretch", hide_index=True)

    summary = (
        df.groupby("examino")
        .agg(Γραμμές=("id", "size"), Με_ώρα=("placed", "sum"))
        .rename(columns={"Με_ώρα": "Με ώρα"})
    )
    summary["Χωρίς ώρα"] = summary["Γραμμές"] - summary["Με ώρα"]
    st.dataframe(summary, width="stretch")

    stale = tdb.off_programme(df, year)
    if not stale.empty:
        st.warning(
            f"{len(stale)} γραμμές ανήκουν σε πρόγραμμα σπουδών διαφορετικό από αυτό που "
            "ακολουθεί το εξάμηνό τους φέτος (μεταφέρθηκαν από πέρυσι). "
            "Αντικαταστήστε τις με τα μαθήματα του σωστού προγράμματος."
        )
        st.dataframe(
            stale[["examino", "curriculum", "course_code", "display_name", "section", "day", "start_time"]]
            .rename(columns={**DISPLAY_COLUMNS, "curriculum": "Πρόγραμμα"}),
            width="stretch",
            hide_index=True,
        )

    if not coordinator:
        st.info("Οι αλλαγές γίνονται από τους συντονιστές (`coordinator_emails`).")
    elif not editable:
        if open_term:
            st.info(f"Ανοιχτό για επεξεργασία είναι το {tdb.term_label(*open_term)}.")
        else:
            st.markdown("#### Άνοιγμα νέου εξαμήνου")
            st.caption(
                "Το νέο εξάμηνο ξεκινά ως αντίγραφο του επιλεγμένου παραπάνω· "
                "συνήθως η ίδια περίοδος του προηγούμενου έτους."
            )
            with st.form("open_term"):
                new_year = st.number_input("Ακαδημαϊκό έτος (έναρξη):", value=year + 1, step=1)
                new_period = st.selectbox("Περίοδος:", options=list(tdb.PERIODS), index=tdb.PERIODS.index(period))
                if st.form_submit_button("Άνοιγμα", type="primary"):
                    error = tdb.open_term(int(new_year), new_period, year, period, user_email)
                    if error:
                        st.error(error)
                    else:
                        st.success(f"Άνοιξε το {tdb.term_label(int(new_year), new_period)}.")
                        st.rerun()
    else:
        staff = tdb.staff_for_term(year, period)
        available_staff = staff[staff["active"].fillna(True)]
        staff_options = list(available_staff["id"])
        staff_names = dict(zip(available_staff["id"], available_staff["short_name"]))
        rooms_frame = tdb.load_rooms(active_only=True)
        room_options = list(rooms_frame["code"])
        room_names = dict(zip(rooms_frame["code"], rooms_frame["name"]))
        candidates = tdb.candidate_courses(year, period)

        st.markdown("#### Επεξεργασία ανά εξάμηνο")
        semesters = sorted(
            {int(s) for s in df["examino"].unique()} | {int(s) for s in candidates["examino"].unique()}
        )
        semester = st.selectbox(
            "Εξάμηνο σπουδών:", options=semesters, format_func=lambda s: f"Εξάμηνο {s}", key="edit_semester"
        )
        subset = df[df["examino"] == semester]
        st.dataframe(to_display(subset), width="stretch", hide_index=True)
        render_week(subset, f"edit_week_{year}_{period}_{semester}", title_with_room)

        not_in_term = candidates[(candidates["examino"] == semester) & ~candidates["in_term"]]
        if not not_in_term.empty:
            with st.expander(f"Μαθήματα του προγράμματος σπουδών που λείπουν ({len(not_in_term)})"):
                st.dataframe(
                    not_in_term[["curriculum", "course_code", "course_name"]].rename(
                        columns={"curriculum": "Πρόγραμμα", "course_code": "Κωδικός", "course_name": "Μάθημα"}
                    ),
                    width="stretch",
                    hide_index=True,
                )

        # Course choices: what the term already has for this εξάμηνο, plus
        # what the curriculum offers and the term does not have yet.
        course_choices: dict[str, tuple[str, int, str]] = {}
        for _, row in subset.drop_duplicates(["course_code", "curriculum"]).iterrows():
            course_choices[f"{row['course_code']} — {row['course_name']}"] = (
                row["course_code"], int(row["curriculum"]), row["name_suffix"] or "")
        for _, row in not_in_term.iterrows():
            course_choices.setdefault(
                f"{row['course_code']} — {row['course_name']} (νέο)",
                (row["course_code"], int(row["curriculum"]), ""),
            )

        def placement_fields(prefix: str, row=None) -> tuple:
            is_placed = st.checkbox(
                "Με ημέρα και ώρα", value=bool(row is not None and row["placed"]), key=f"{prefix}_placed"
            )
            c1, c2, c3 = st.columns(3)
            day = c1.selectbox(
                "Ημέρα:", options=list(range(1, 6)), format_func=lambda d: tdb.DAYS[d - 1],
                index=int(row["day_number"]) - 1 if row is not None and row["placed"] else 0,
                key=f"{prefix}_day",
            )
            hours = list(range(tdb.FIRST_HOUR, tdb.LAST_HOUR))
            start = c2.selectbox(
                "Έναρξη:", options=hours, format_func=lambda h: f"{h:02d}:00",
                index=hours.index(int(row["start_hour"])) if row is not None and row["placed"] else 1,
                key=f"{prefix}_start",
            )
            duration = c3.number_input(
                "Διάρκεια (ώρες):", min_value=1, max_value=tdb.MAX_DURATION,
                value=int(row["duration"]) if row is not None and row["placed"] else 2,
                key=f"{prefix}_duration",
            )
            if not is_placed:
                return None, None, None
            return int(day), int(start), int(duration)

        sections = ["Θ", "Ε", "Ε1", "Ε2", "Ε3", "Ε4", "Φ"]

        st.markdown("#### Νέα γραμμή")
        with st.form("add_class", clear_on_submit=False):
            choice = st.selectbox("Μάθημα:", options=list(course_choices), key="add_course")
            c1, c2 = st.columns(2)
            section = c1.selectbox("Τμήμα:", options=sections, key="add_section")
            suffix = c2.text_input("Ένδειξη μετά τον τίτλο (π.χ. ΔΥ, ΥΕ):", key="add_suffix")
            instructors = st.multiselect(
                "Διδάσκοντες:", options=staff_options, format_func=lambda i: staff_names[i], key="add_staff"
            )
            rooms = st.multiselect(
                "Αίθουσες:", options=room_options,
                format_func=lambda c: f"{c} — {room_names[c]}", key="add_rooms",
            )
            day, start, duration = placement_fields("add")
            notes = st.text_input("Παρατηρήσεις:", key="add_notes")
            if st.form_submit_button("Προσθήκη", type="primary"):
                code, curriculum, default_suffix = course_choices[choice]
                error = tdb.add_class(
                    year, period, examino=semester, course_code=code, curriculum=curriculum,
                    section=section, instructor_ids=instructors, room_codes=rooms,
                    day=day, start_hour=start, duration=duration,
                    name_suffix=suffix or default_suffix, notes=notes, author=user_email,
                )
                if error:
                    st.error(error)
                else:
                    st.success("Προστέθηκε.")
                    st.rerun()

        st.markdown("#### Μεταβολή ή διαγραφή γραμμής")
        if subset.empty:
            st.info("Δεν υπάρχουν γραμμές σε αυτό το εξάμηνο.")
        else:
            # Outside the form on purpose: a widget inside one does not rerun
            # until submit, so the fields below would show the previous row.
            labels = {int(r["id"]): class_label(r) for _, r in subset.iterrows()}
            class_id = st.selectbox(
                "Γραμμή:", options=list(labels), format_func=lambda i: labels[i], key="edit_class"
            )
            row = subset[subset["id"] == class_id].iloc[0]
            prefix = f"edit_{class_id}"
            with st.form(f"edit_form_{class_id}"):
                c1, c2, c3 = st.columns(3)
                examino = c1.number_input("Εξάμηνο:", min_value=1, max_value=10, value=int(row["examino"]), key=f"{prefix}_examino")
                section = c2.selectbox(
                    "Τμήμα:", options=sections,
                    index=sections.index(row["section"]) if row["section"] in sections else 0,
                    key=f"{prefix}_section",
                )
                suffix = c3.text_input("Ένδειξη μετά τον τίτλο:", value=row["name_suffix"] or "", key=f"{prefix}_suffix")
                instructors = st.multiselect(
                    "Διδάσκοντες:", options=staff_options, format_func=lambda i: staff_names[i],
                    default=[i for i in row["instructor_ids"] if i in staff_names], key=f"{prefix}_staff",
                )
                rooms = st.multiselect(
                    "Αίθουσες:", options=room_options, format_func=lambda c: f"{c} — {room_names[c]}",
                    default=[c for c in row["room_codes"] if c in room_names], key=f"{prefix}_rooms",
                )
                day, start, duration = placement_fields(prefix, row)
                notes = st.text_input("Παρατηρήσεις:", value=row["notes"] or "", key=f"{prefix}_notes")
                b1, b2 = st.columns(2)
                if b1.form_submit_button("Αποθήκευση", type="primary"):
                    error = tdb.update_class(
                        class_id, examino=int(examino), section=section, instructor_ids=instructors,
                        room_codes=rooms, day=day, start_hour=start, duration=duration,
                        name_suffix=suffix, notes=notes, author=user_email,
                    )
                    if error:
                        st.error(error)
                    else:
                        st.success("Αποθηκεύτηκε.")
                        st.rerun()
                if b2.form_submit_button("Διαγραφή"):
                    error = tdb.delete_class(class_id, user_email)
                    if error:
                        st.error(error)
                    else:
                        st.success("Διαγράφηκε.")
                        st.rerun()

        st.markdown("#### Κλείδωμα")
        st.caption("Μετά το κλείδωμα το εξάμηνο δεν αλλάζει· το επόμενο ανοίγει ως αντίγραφό του.")
        confirm = st.checkbox(f"Κλείδωμα του {tdb.term_label(year, period)}", key="lock_confirm")
        if st.button("Κλείδωμα", type="primary", disabled=not confirm, key="lock_button"):
            error = tdb.lock_term(year, period, user_email)
            if error:
                st.error(error)
            else:
                st.success("Κλειδώθηκε.")
                st.rerun()

    with st.expander("Ιστορικό αλλαγών"):
        changes = tdb.term_changes(year, period)
        if changes.empty:
            st.caption("Καμία καταγεγραμμένη αλλαγή.")
        else:
            st.dataframe(
                changes.rename(columns={"created_at": "Πότε", "author": "Ποιος", "action": "Ενέργεια",
                                        "class_id": "Γραμμή", "detail": "Τι"}),
                width="stretch",
                hide_index=True,
            )

# --------------------------------------------------------------------------
# Προσωπικό
# --------------------------------------------------------------------------

with tab_staff:
    st.markdown(f"### Προσωπικό — ενεργοί το {tdb.term_label(year, period)}")
    staff = tdb.staff_for_term(year, period)
    view = staff[["id", "short_name", "last_name", "first_name", "category", "rank", "active", "subject", "notes"]]
    labels = {
        "short_name": "Σύντομο όνομα", "last_name": "Επώνυμο", "first_name": "Όνομα",
        "category": "Κατηγορία", "rank": "Βαθμίδα", "active": "Ενεργός/ή", "subject": "Αντικείμενο", "notes": "Σημειώσεις",
    }
    if coordinator:
        st.caption("Αλλάξτε τη στήλη «Ενεργός/ή» για το επιλεγμένο εξάμηνο και πατήστε Αποθήκευση.")
        edited = st.data_editor(
            view.rename(columns=labels),
            disabled=[c for c in labels.values() if c != "Ενεργός/ή"] + ["id"],
            hide_index=True,
            width="stretch",
            key=f"staff_editor_{year}_{period}",
        )
        if st.button("Αποθήκευση ενεργών", key="save_active"):
            changed = 0
            for (_, before), (_, after) in zip(view.iterrows(), edited.iterrows()):
                new_value = after["Ενεργός/ή"]
                if pd.isna(new_value):
                    continue
                if pd.isna(before["active"]) or bool(before["active"]) != bool(new_value):
                    tdb.set_staff_active(int(before["id"]), year, period, bool(new_value))
                    changed += 1
            st.success(f"Ενημερώθηκαν {changed} εγγραφές.")
            if changed:
                st.rerun()

        st.markdown("#### Προσθήκη ή διόρθωση προσώπου")
        existing = {int(r["id"]): r["short_name"] for _, r in staff.iterrows()}
        target = st.selectbox(
            "Πρόσωπο:", options=[0] + list(existing),
            format_func=lambda i: "— νέο πρόσωπο —" if i == 0 else existing[i], key="staff_target",
        )
        current = staff[staff["id"] == target].iloc[0] if target else None
        prefix = f"staff_{target}"

        def val(column: str) -> str:
            return "" if current is None or pd.isna(current[column]) else str(current[column])

        with st.form(f"staff_form_{target}"):
            c1, c2, c3 = st.columns(3)
            short_name = c1.text_input("Σύντομο όνομα (όπως τυπώνεται):", value=val("short_name"), key=f"{prefix}_short")
            last_name = c2.text_input("Επώνυμο:", value=val("last_name"), key=f"{prefix}_last")
            first_name = c3.text_input("Όνομα:", value=val("first_name"), key=f"{prefix}_first")
            c1, c2 = st.columns(2)
            category = c1.selectbox(
                "Κατηγορία:", options=list(tdb.CATEGORIES),
                index=tdb.CATEGORIES.index(current["category"]) if current is not None else 0,
                key=f"{prefix}_category",
            )
            rank = c2.text_input("Βαθμίδα:", value=val("rank"), key=f"{prefix}_rank")
            subject = st.text_input("Γνωστικό αντικείμενο:", value=val("subject"), key=f"{prefix}_subject")
            c1, c2 = st.columns(2)
            email = c1.text_input("Email:", value=val("email"), key=f"{prefix}_email")
            website = c2.text_input("Σελίδα στο site:", value=val("website_url"), key=f"{prefix}_url")
            notes = st.text_input("Σημειώσεις:", value=val("notes"), key=f"{prefix}_notes")
            if st.form_submit_button("Αποθήκευση", type="primary"):
                error = tdb.save_staff(
                    {
                        "short_name": short_name, "last_name": last_name, "first_name": first_name,
                        "category": category, "rank": rank, "subject": subject, "email": email,
                        "website_url": website, "notes": notes,
                        "placeholder": bool(current is not None and current["placeholder"]),
                    },
                    user_email,
                    staff_id=target or None,
                )
                if error:
                    st.error(error)
                else:
                    st.success("Αποθηκεύτηκε.")
                    st.rerun()
    else:
        st.dataframe(view.rename(columns=labels), hide_index=True, width="stretch")

# --------------------------------------------------------------------------
# Αίθουσες
# --------------------------------------------------------------------------

with tab_room_list:
    st.markdown("### Αίθουσες")
    rooms_frame = tdb.load_rooms()
    room_labels = {"code": "Κωδικός", "name": "Όνομα", "kind": "Είδος", "capacity": "Χωρητικότητα",
                   "active": "Σε χρήση", "notes": "Σημειώσεις"}
    st.dataframe(rooms_frame.rename(columns=room_labels), hide_index=True, width="stretch")
    if coordinator:
        st.markdown("#### Προσθήκη ή διόρθωση αίθουσας")
        codes = list(rooms_frame["code"])
        target = st.selectbox(
            "Αίθουσα:", options=[""] + codes,
            format_func=lambda c: "— νέα αίθουσα —" if c == "" else c, key="room_target",
        )
        current = rooms_frame[rooms_frame["code"] == target].iloc[0] if target else None
        prefix = f"room_{target or 'new'}"
        with st.form(f"room_form_{target or 'new'}"):
            c1, c2, c3 = st.columns(3)
            code = c1.text_input("Κωδικός:", value=target, disabled=bool(target), key=f"{prefix}_code")
            name = c2.text_input("Όνομα:", value="" if current is None else current["name"], key=f"{prefix}_name")
            kind = c3.selectbox(
                "Είδος:", options=list(tdb.ROOM_KINDS),
                index=tdb.ROOM_KINDS.index(current["kind"]) if current is not None else 0,
                key=f"{prefix}_kind",
            )
            c1, c2 = st.columns(2)
            capacity = c1.number_input(
                "Χωρητικότητα:", min_value=0, step=1,
                value=0 if current is None or pd.isna(current["capacity"]) else int(current["capacity"]),
                key=f"{prefix}_capacity",
            )
            active = c2.checkbox("Σε χρήση", value=True if current is None else bool(current["active"]), key=f"{prefix}_active")
            notes = st.text_input("Σημειώσεις:", value="" if current is None or pd.isna(current["notes"]) else current["notes"], key=f"{prefix}_notes")
            if st.form_submit_button("Αποθήκευση", type="primary"):
                error = tdb.save_room(
                    {"code": target or code, "name": name, "kind": kind,
                     "capacity": capacity or None, "active": active, "notes": notes}
                )
                if error:
                    st.error(error)
                else:
                    st.success("Αποθηκεύτηκε.")
                    st.rerun()
