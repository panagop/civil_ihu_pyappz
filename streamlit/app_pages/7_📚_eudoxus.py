"""Εύδοξος — the books offered per academic year, and next year's list.

Four tabs: browse a stored year, check every book of the open year against
Eudoxus, edit the open year course by course, and (coordinator only) open a new
year from the previous one, review the net change and lock it.

Like page 6 this is database-only: the workbook in ``files/eudoxus`` is the
pre-handover archive, and serving it as a live list would be showing last
year's answer to next year's question.
"""

import io
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Εύδοξος",
    layout="wide",
)

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import db  # noqa: E402
import eudoxus_db as edb  # noqa: E402
from auth import require_ihu_login  # noqa: E402
from branding import apply_branding  # noqa: E402
from eudoxus_client import Eudoxus  # noqa: E402

require_ihu_login()
apply_branding()

# Also called from home.py, but Streamlit runs only the page you open.
db.bootstrap()

st.markdown("## Εύδοξος — συγγράμματα ανά ακαδημαϊκό έτος")

if not db.is_available():
    st.error(
        "Η σελίδα λειτουργεί μόνο με σύνδεση στη βάση δεδομένων, "
        "η οποία είναι προσβάσιμη μόνο μέσα από το Railway."
    )
    st.stop()

years = edb.stored_years()
if not years:
    st.warning("Δεν υπάρχουν αποθηκευμένες λίστες συγγραμμάτων.")
    st.stop()

user_email = getattr(st.user, "email", "") or ""
coordinator = db.is_coordinator(user_email)
open_year_list = edb.open_years()
working_year = open_year_list[0] if open_year_list else None

# Columns worth showing in the browse and problem tables, in reading order.
DISPLAY_COLUMNS = {
    "examino": "Εξάμηνο",
    "course_code": "Κωδικός",
    "course_title": "Μάθημα",
    "teacher": "Διδάσκων",
    "priority": "Σειρά",
    "book_id": "Κωδ. Ευδόξου",
    "book_title": "Βιβλίο",
    "authors": "Συγγραφείς",
    "publisher": "Εκδότης",
    "publication_year": "Έτος",
    "isbn": "ISBN",
}


def to_display(frame: pd.DataFrame, extra: dict[str, str] | None = None) -> pd.DataFrame:
    columns = {**DISPLAY_COLUMNS, **(extra or {})}
    present = [name for name in columns if name in frame.columns]
    return frame[present].rename(columns=columns)


def to_excel(frame: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        frame.to_excel(writer, index=False, sheet_name="books")
    return buffer.getvalue()


tab_browse, tab_check, tab_edit, tab_admin = st.tabs(
    ["Βιβλία ανά έτος", "Έλεγχος διαθεσιμότητας", "Επεξεργασία", "Συντονιστής"]
)


# --------------------------------------------------------------------------
with tab_browse:
    browse_year = st.selectbox(
        "Ακαδημαϊκό έτος",
        years,
        format_func=edb.year_label,
        key="eudoxus_browse_year",
    )
    state = edb.year_state(browse_year)
    if state:
        st.caption(
            f"Κατάσταση: {state['status']}"
            + (
                f" · με βάση το {edb.year_label(state['baseline_year'])}"
                if state.get("baseline_year")
                else ""
            )
        )

    frame = edb.load_year(browse_year)
    if frame.empty:
        st.info("Καμία καταχώρηση για το έτος αυτό.")
    else:
        filters = st.columns(4)
        with filters[0]:
            periods = ["Όλες"] + sorted(frame["period"].dropna().unique().tolist())
            period = st.selectbox(
                "Περίοδος",
                periods,
                format_func=lambda value: edb.PERIOD_LABELS.get(value, value),
                key="eudoxus_period",
            )
        with filters[1]:
            semesters = ["Όλα"] + sorted(frame["examino"].dropna().unique().tolist())
            semester = st.selectbox("Εξάμηνο", semesters, key="eudoxus_examino")
        with filters[2]:
            teachers = ["Όλοι"] + sorted(frame["teacher"].dropna().unique().tolist())
            teacher = st.selectbox("Διδάσκων", teachers, key="eudoxus_teacher")
        with filters[3]:
            search = st.text_input("Αναζήτηση", key="eudoxus_search")

        shown = frame
        if period != "Όλες":
            shown = shown[shown["period"] == period]
        if semester != "Όλα":
            shown = shown[shown["examino"] == semester]
        if teacher != "Όλοι":
            shown = shown[shown["teacher"] == teacher]
        if search.strip():
            needle = search.strip().casefold()
            haystack = (
                shown[["course_code", "course_title", "book_title", "authors"]]
                .fillna("")
                .agg(" ".join, axis=1)
                .str.casefold()
            )
            shown = shown[haystack.str.contains(needle, regex=False)]

        counters = st.columns(4)
        counters[0].metric("Μαθήματα", shown.groupby(["course_code", "examino"]).ngroups)
        counters[1].metric("Επιλογές", len(shown))
        counters[2].metric("Διακριτά βιβλία", shown["book_id"].nunique())
        unknown = int(shown["checked_at"].isna().sum())
        counters[3].metric("Χωρίς έλεγχο", unknown)

        st.dataframe(to_display(shown), width="stretch", hide_index=True)
        st.download_button(
            "Λήψη σε Excel",
            data=to_excel(to_display(shown)),
            file_name=f"eudoxus_{edb.year_label(browse_year)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# --------------------------------------------------------------------------
with tab_check:
    check_year = st.selectbox(
        "Ακαδημαϊκό έτος",
        years,
        index=years.index(working_year) if working_year in years else 0,
        format_func=edb.year_label,
        key="eudoxus_check_year",
    )
    book_ids = edb.book_ids_for_year(check_year)
    catalogue = edb.load_books(book_ids)
    checked = (
        int(catalogue["checked_at"].notna().sum()) if not catalogue.empty else 0
    )
    st.caption(
        f"{len(book_ids)} διακριτά βιβλία · {checked} με καταγεγραμμένο έλεγχο. "
        "Ο έλεγχος ρωτά τον Εύδοξο ένα βιβλίο τη φορά — περίπου ένα δευτερόλεπτο "
        "για κάθε ένα."
    )

    if st.button("Έλεγχος διαθεσιμότητας τώρα", key="run_check", type="primary"):
        progress = st.progress(0.0, text="Έναρξη…")

        def report(done: int, total: int) -> None:
            progress.progress(done / total, text=f"{done} / {total} βιβλία")

        client = Eudoxus()
        rows = client.fetch_books(book_ids, progress=report)
        written = edb.upsert_books(rows)
        progress.empty()
        failures = [row for row in rows if not row["found"]]
        st.success(
            f"Ελέγχθηκαν {written} βιβλία."
            + (f" {len(failures)} δεν βρέθηκαν." if failures else "")
        )
        st.rerun()

    problems = edb.unavailable_for_year(check_year)
    frame = edb.load_year(check_year)
    never_checked = frame[frame["checked_at"].isna()] if not frame.empty else frame

    metrics = st.columns(3)
    metrics[0].metric("Προβληματικά", 0 if problems.empty else problems["book_id"].nunique())
    metrics[1].metric(
        "Μαθήματα που θίγονται",
        0 if problems.empty else problems.groupby(["course_code", "examino"]).ngroups,
    )
    metrics[2].metric(
        "Χωρίς έλεγχο", 0 if never_checked.empty else never_checked["book_id"].nunique()
    )

    if problems.empty:
        st.info(
            "Κανένα πρόβλημα στα βιβλία που έχουν ελεγχθεί."
            if checked
            else "Δεν έχει γίνει ακόμη έλεγχος."
        )
    else:
        st.markdown("### Βιβλία που δεν μπορούν να επιλεγούν ξανά")
        st.dataframe(
            to_display(problems, {"reason": "Αιτία"}),
            width="stretch",
            hide_index=True,
        )
        st.download_button(
            "Λήψη σε Excel",
            data=to_excel(to_display(problems, {"reason": "Αιτία"})),
            file_name=f"eudoxus_problems_{edb.year_label(check_year)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if not never_checked.empty:
        with st.expander(f"Χωρίς έλεγχο ({never_checked['book_id'].nunique()} βιβλία)"):
            st.caption(
                "Δεν αναφέρονται ως προβληματικά: η απουσία απάντησης δεν είναι "
                "αρνητική απάντηση."
            )
            st.dataframe(to_display(never_checked), width="stretch", hide_index=True)


# --------------------------------------------------------------------------
with tab_edit:
    if working_year is None:
        st.info(
            "Δεν υπάρχει ανοιχτό έτος. Ο Συντονιστής ανοίγει το επόμενο έτος "
            "από την καρτέλα «Συντονιστής»."
        )
    else:
        st.markdown(f"### Λίστα {edb.year_label(working_year)}")
        courses = edb.courses_for_year(working_year)
        if courses.empty:
            st.warning("Το ανοιχτό έτος δεν έχει μαθήματα.")
        else:
            # (code, εξάμηνο) rather than code alone: ΔΟΜ022 runs twice with a
            # different book order, and picking "ΔΟΜ022" would be ambiguous.
            keys = list(zip(courses["course_code"], courses["examino"]))
            labels = {
                key: f"{key[0]} — {row.title} ({key[1]}ο εξάμηνο, {row.books} βιβλία)"
                for key, (_, row) in zip(keys, courses.iterrows())
            }
            course_key = st.selectbox(
                "Μάθημα",
                keys,
                format_func=lambda key: labels[key],
                key="eudoxus_edit_course",
            )
            course_code, examino = course_key

            year_frame = edb.load_year(working_year)
            current = year_frame[
                (year_frame["course_code"] == course_code)
                & (year_frame["examino"] == examino)
            ].sort_values("priority")

            if current.empty:
                st.warning("Το μάθημα δεν έχει βιβλία.")
            else:
                st.dataframe(
                    to_display(current), width="stretch", hide_index=True
                )
                unusable = current[
                    current["checked_at"].notna()
                    & (
                        ~current["found"].fillna(False)
                        | ~current["active"].fillna(False)
                        | ~current["selectable"].fillna(False)
                    )
                ]
                if not unusable.empty:
                    st.warning(
                        "Δεν μπορούν να επιλεγούν ξανά: "
                        + ", ".join(str(book) for book in unusable["book_id"])
                    )

            sub_order, sub_remove, sub_add = st.tabs(
                ["Σειρά επιλογής", "Αφαίρεση", "Προσθήκη"]
            )

            with sub_order:
                if current.empty:
                    st.caption("Καμία επιλογή.")
                else:
                    with st.form(f"order_{course_code}_{examino}"):
                        wanted: dict[int, int] = {}
                        for _, row in current.iterrows():
                            title = row["book_title"] or f"({row['book_id']})"
                            wanted[int(row["book_id"])] = st.number_input(
                                f"{title} [{row['book_id']}]",
                                value=int(row["priority"]),
                                step=1,
                                min_value=0,
                                key=f"prio_{course_code}_{examino}_{row['book_id']}",
                            )
                        if st.form_submit_button("Αποθήκευση σειράς"):
                            ok, message = edb.set_priorities(
                                working_year, course_code, examino, wanted, user_email
                            )
                            (st.success if ok else st.warning)(message)
                            if ok:
                                st.rerun()

            with sub_remove:
                if current.empty:
                    st.caption("Καμία επιλογή.")
                else:
                    options = current["book_id"].astype(int).tolist()
                    titles = dict(zip(current["book_id"].astype(int), current["book_title"]))
                    victim = st.selectbox(
                        "Βιβλίο",
                        options,
                        format_func=lambda book: f"{book} — {titles.get(book) or '—'}",
                        key=f"remove_pick_{course_code}_{examino}",
                    )
                    reason = st.text_input(
                        "Αιτιολόγηση", key=f"remove_reason_{course_code}_{examino}"
                    )
                    if st.button("Αφαίρεση", key=f"remove_go_{course_code}_{examino}"):
                        ok, message = edb.remove_book(
                            working_year, course_code, examino, int(victim),
                            user_email, reason,
                        )
                        (st.success if ok else st.warning)(message)
                        if ok:
                            st.rerun()

            with sub_add:
                st.caption(
                    "Αναζήτηση στον Εύδοξο. Η αναζήτηση τίτλου/συγγραφέα είναι "
                    "υποσυμβολοσειρά, χωρίς διάκριση τόνων και πεζών-κεφαλαίων."
                )
                search_columns = st.columns(3)
                with search_columns[0]:
                    by_title = st.text_input(
                        "Τίτλος", key=f"add_title_{course_code}_{examino}"
                    )
                with search_columns[1]:
                    by_author = st.text_input(
                        "Συγγραφέας", key=f"add_author_{course_code}_{examino}"
                    )
                with search_columns[2]:
                    by_id = st.text_input(
                        "Κωδικός Ευδόξου", key=f"add_id_{course_code}_{examino}"
                    )

                if st.button("Αναζήτηση", key=f"add_search_{course_code}_{examino}"):
                    filters = {}
                    if by_title.strip():
                        filters["title"] = by_title.strip()
                    if by_author.strip():
                        filters["authors"] = by_author.strip()
                    if by_id.strip():
                        filters["id"] = by_id.strip()
                    if not filters:
                        st.warning("Συμπλήρωσε τουλάχιστον ένα κριτήριο.")
                    else:
                        client = Eudoxus()
                        with st.spinner("Αναζήτηση στον Εύδοξο…"):
                            st.session_state["eudoxus_hits"] = client.search_books(
                                page_size=25, **filters
                            )

                hits = st.session_state.get("eudoxus_hits") or []
                if hits:
                    hits_frame = pd.DataFrame(hits)
                    st.dataframe(
                        hits_frame[
                            ["book_id", "title", "authors", "publisher",
                             "publication_year", "active", "selectable"]
                        ],
                        width="stretch",
                        hide_index=True,
                    )
                    choices = hits_frame["book_id"].astype(int).tolist()
                    hit_titles = dict(zip(hits_frame["book_id"].astype(int), hits_frame["title"]))
                    picked = st.selectbox(
                        "Βιβλίο προς προσθήκη",
                        choices,
                        format_func=lambda book: f"{book} — {hit_titles.get(book)}",
                        key=f"add_pick_{course_code}_{examino}",
                    )
                    next_priority = 0 if current.empty else int(current["priority"].max()) + 1
                    priority = st.number_input(
                        "Σειρά επιλογής",
                        value=next_priority,
                        step=1,
                        min_value=0,
                        key=f"add_prio_{course_code}_{examino}",
                    )
                    if st.button("Προσθήκη", key=f"add_go_{course_code}_{examino}"):
                        chosen = hits_frame[hits_frame["book_id"] == picked].to_dict("records")
                        # Store what the search already told us, so the new book
                        # has a title in the list without a second round trip.
                        edb.upsert_books(chosen)
                        ok, message = edb.add_book(
                            working_year, course_code, examino, int(picked),
                            int(priority), user_email,
                        )
                        (st.success if ok else st.warning)(message)
                        if ok:
                            st.rerun()


# --------------------------------------------------------------------------
with tab_admin:
    if not coordinator:
        st.info("Η καρτέλα αυτή είναι διαθέσιμη στον Συντονιστή.")
    else:
        st.markdown("### Άνοιγμα νέου έτους")
        if working_year is not None:
            st.info(
                f"Ανοιχτό έτος: {edb.year_label(working_year)}. "
                "Κλείδωσέ το πριν ανοίξεις το επόμενο."
            )
        else:
            columns = st.columns(2)
            with columns[0]:
                baseline = st.selectbox(
                    "Έτος βάσης", years, format_func=edb.year_label, key="admin_baseline"
                )
            with columns[1]:
                new_year = st.number_input(
                    "Νέο έτος",
                    value=int(max(years)) + 1,
                    step=1,
                    key="admin_new_year",
                )
            st.caption(
                f"Το {edb.year_label(int(new_year))} θα ξεκινήσει ως αντίγραφο "
                f"του {edb.year_label(baseline)} και θα είναι επεξεργάσιμο."
            )
            if st.button("Άνοιγμα έτους", type="primary", key="admin_open"):
                st.info(edb.open_year(int(new_year), int(baseline), user_email))
                st.rerun()

        if working_year is not None:
            st.divider()
            st.markdown(f"### Μεταβολές του {edb.year_label(working_year)}")
            diff = edb.changes_vs_baseline(working_year)
            if diff.empty:
                st.caption("Καμία διαφορά από το έτος βάσης.")
            else:
                st.caption(
                    f"{len(diff)} μεταβολές σε "
                    f"{diff.groupby(['course_code', 'examino']).ngroups} μαθήματα, "
                    "σε σύγκριση με το έτος βάσης."
                )
                st.dataframe(
                    diff.rename(
                        columns={
                            "course_code": "Κωδικός",
                            "examino": "Εξάμηνο",
                            "book_id": "Κωδ. Ευδόξου",
                            "book_title": "Βιβλίο",
                            "authors": "Συγγραφείς",
                            "action": "Ενέργεια",
                            "priority_before": "Σειρά πριν",
                            "priority_after": "Σειρά μετά",
                        }
                    ),
                    width="stretch",
                    hide_index=True,
                )

            with st.expander("Ιστορικό ενεργειών"):
                log = edb.year_changes(working_year)
                if log.empty:
                    st.caption("Καμία καταγραφή.")
                else:
                    st.dataframe(log, width="stretch", hide_index=True)

            st.divider()
            st.markdown("### Κλείδωμα")
            st.caption(
                "Μετά το κλείδωμα το έτος γίνεται ιστορικό και δεν επεξεργάζεται."
            )
            confirm = st.checkbox(
                f"Επιβεβαιώνω το κλείδωμα του {edb.year_label(working_year)}",
                key="admin_confirm_lock",
            )
            if st.button("Κλείδωμα έτους", disabled=not confirm, key="admin_lock"):
                st.info(edb.lock_year(working_year, user_email))
                st.rerun()
