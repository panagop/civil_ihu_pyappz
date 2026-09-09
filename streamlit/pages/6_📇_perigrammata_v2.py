"""Περιγράμματα μαθημάτων v2 — served from Postgres, edited in the app.

Page 1 reads the Google Sheet; this one does not, ever. The sheet was exported
once into ``files/perigrammata/`` and the database is the master from then on,
so there is no "reload from Google Sheets" button here and no fallback to the
archive: with no database the page stops rather than showing data that has been
stale since the first edit landed.
"""

import sys
from datetime import date, datetime, time, timezone
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Περιγράμματα v2",
    layout="wide",
)

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import db  # noqa: E402
import perigrammata_db as pdb  # noqa: E402
import perigrammata_report as report  # noqa: E402
from auth import require_ihu_login  # noqa: E402

require_ihu_login()

# Also called from home.py, but Streamlit runs only the page you open: a
# visitor landing straight here would otherwise find the courses unloaded.
db.bootstrap()

LOCALE = "gr"  # Αγγλικά έπονται — see SEED_LOCALES in seed_perigrammata.py

st.markdown("## Περιγράμματα μαθημάτων")

if not db.is_available():
    st.error(
        "Η σελίδα λειτουργεί μόνο με σύνδεση στη βάση δεδομένων, "
        "η οποία είναι προσβάσιμη μόνο μέσα από το Railway."
    )
    st.caption(
        "Τα περιγράμματα τηρούνται πλέον στη βάση· το Google Sheet δεν "
        "χρησιμοποιείται. Για τοπική δοκιμή χρειάζεται DATABASE_URL."
    )
    st.stop()

curricula = pdb.stored_curricula(LOCALE)
if not curricula:
    st.warning("Δεν υπάρχουν αποθηκευμένα περιγράμματα στη βάση.")
    st.stop()

curriculum = st.radio(
    "Πρόγραμμα σπουδών",
    curricula,
    format_func=lambda year: f"Πρόγραμμα σπουδών {year}",
    horizontal=True,
)
editable = curriculum == pdb.EDITABLE_CURRICULUM
if not editable:
    st.caption(
        f"Το πρόγραμμα σπουδών {curriculum} τηρείται ως ιστορικό αρχείο "
        "και δεν είναι επεξεργάσιμο."
    )

# Not cached on purpose: ~100 rows in one query, against a form that writes to
# the same table. A cache here buys milliseconds and costs stale reads after
# every save.
courses = pdb.load_courses(curriculum, LOCALE)
user_email = getattr(st.user, "email", "") or ""
coordinator = db.is_coordinator(user_email)


def course_label(code: str) -> str:
    row = courses[courses["code"] == code]
    name = row["name"].iat[0] if not row.empty else ""
    return f"{code} — {name}" if name else str(code)


tab_table, tab_stats, tab_edit, tab_word, tab_reports = st.tabs(
    ["Πίνακας", "Στατιστικά", "Επεξεργασία", "Αρχείο Word", "Αναφορές"]
)


with tab_table:
    st.caption(f"{len(courses)} μαθήματα · πρόγραμμα σπουδών {curriculum}")
    st.dataframe(courses, use_container_width=True, hide_index=True)


with tab_stats:
    left, right = st.columns(2)
    with left:
        st.markdown("### Αριθμός μαθημάτων ανά εξάμηνο")
        st.bar_chart(courses["examino"].value_counts().sort_index())
    with right:
        st.markdown("### Τύπος μαθημάτων")
        st.bar_chart(courses["type"].value_counts())

    st.markdown("### Πρόσφατες ενημερώσεις")
    recent = courses[["code", "name", "updated_at", "updated_by"]].sort_values(
        "updated_at", ascending=False
    )
    st.dataframe(recent.head(15), use_container_width=True, hide_index=True)


with tab_edit:
    if not editable:
        st.info(
            f"Επεξεργασία γίνεται μόνο στο πρόγραμμα σπουδών "
            f"{pdb.EDITABLE_CURRICULUM}."
        )
    else:
        # Outside the form on purpose: a widget inside st.form does not rerun
        # until submit, so the fields below would keep showing the previously
        # selected course.
        edit_code = st.selectbox(
            "Μάθημα",
            courses["code"].tolist(),
            format_func=course_label,
            key="perigrammata_edit_code",
        )
        current = pdb.load_course(curriculum, LOCALE, edit_code)
        if current is None:
            st.error("Το μάθημα δεν βρέθηκε.")
            st.stop()

        stamp = current["updated_at"]
        by = current.get("updated_by") or "—"
        st.caption(f"Τελευταία ενημέρωση: {stamp:%d/%m/%Y %H:%M} · {by}")

        # Every widget key carries the course code: Streamlit keeps the stored
        # value of a key that has not changed, which would override the new
        # course's defaults with the previous one's text.
        with st.form(f"perigramma_form_{edit_code}"):
            values: dict[str, object] = {}
            for group, fields in pdb.FIELD_GROUPS:
                st.markdown(f"**{group}**")
                for field, label, kind in fields:
                    key = f"perigramma_{edit_code}_{field}"
                    value = current.get(field)
                    if kind == "area":
                        values[field] = st.text_area(
                            label, value="" if value is None else str(value),
                            key=key, height=160,
                        )
                    elif kind in ("num", "int"):
                        values[field] = st.number_input(
                            label,
                            value=None if value is None else float(value),
                            step=1.0,
                            format="%.0f" if kind == "int" else "%g",
                            key=key,
                        )
                    else:
                        values[field] = st.text_input(
                            label, value="" if value is None else str(value), key=key
                        )
                st.divider()

            note = st.text_input(
                "Σχόλιο αλλαγής (προαιρετικό)", key=f"perigramma_note_{edit_code}"
            )
            submitted = st.form_submit_button("Αποθήκευση", type="primary")

        if submitted:
            saved, message = pdb.save_course(
                curriculum,
                LOCALE,
                edit_code,
                values,
                editor=user_email,
                loaded_at=stamp,
                note=note,
            )
            if saved:
                st.success(message)
                st.rerun()
            else:
                st.warning(message)

        workload = [values.get(field) for field in pdb.WORKLOAD_COLUMNS]
        total = sum(float(hours) for hours in workload if hours is not None)
        declared = values.get("hours_sum")
        if declared is not None and total and float(declared) != total:
            st.warning(
                f"Ο δηλωμένος συνολικός φόρτος ({float(declared):g}) διαφέρει "
                f"από το άθροισμα των επιμέρους δραστηριοτήτων ({total:g})."
            )

        history = pdb.course_history(curriculum, LOCALE, edit_code)
        with st.expander(f"Ιστορικό αλλαγών ({len(history)} καταχωρήσεις)"):
            if history.empty:
                st.caption("Καμία καταγραφή.")
            else:
                for _, revision in history.iterrows():
                    changed = revision["changes"] or {}
                    labels = ", ".join(
                        pdb.FIELD_LABELS.get(field, field) for field in changed
                    )
                    stamp_text = pd.to_datetime(revision["edited_at"]).strftime(
                        "%d/%m/%Y %H:%M"
                    )
                    who = revision["edited_by"] or "—"
                    detail = labels or revision["note"] or "αρχική κατάσταση"
                    st.write(f"**{stamp_text}** · {who} — {detail}")


with tab_word:
    word_code = st.selectbox(
        "Μάθημα",
        courses["code"].tolist(),
        format_func=course_label,
        key="perigrammata_word_code",
    )
    row = pdb.load_course(curriculum, LOCALE, word_code)
    if row is None:
        st.error("Το μάθημα δεν βρέθηκε.")
    else:
        with st.expander("Στοιχεία μαθήματος (πλήρη)"):
            st.write(report.docx_context(row))
        st.download_button(
            "Λήψη περιγράμματος",
            data=report.render_course(row, LOCALE),
            file_name=f"Περίγραμμα-{word_code}-{curriculum}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        if any(row.get(field) for field in pdb.UNPRINTED_COLUMNS):
            st.caption(
                "Σημείωση: τα πεδία «Συναφή περιοδικά» αποθηκεύονται αλλά δεν "
                "υπάρχουν στο πρότυπο Word, οπότε δεν τυπώνονται."
            )


@st.cache_data(show_spinner=False)
def _full_report(curriculum: int, locale: str, version: str) -> bytes:
    """Cached on the newest updated_at, so an edit invalidates it and nothing else does."""
    rows = pdb.load_courses(curriculum, locale).to_dict("records")
    return report.build_full_report(rows, locale)


with tab_reports:
    if not coordinator:
        st.info("Οι συγκεντρωτικές αναφορές είναι διαθέσιμες στον Συντονιστή.")
    else:
        st.markdown("### Πλήρης αναφορά")
        st.caption(
            f"Όλα τα {len(courses)} περιγράμματα του προγράμματος σπουδών "
            f"{curriculum}, ένα ανά σελίδα."
        )
        if st.button("Δημιουργία πλήρους αναφοράς", key="build_full"):
            version = str(courses["updated_at"].max())
            with st.spinner("Δημιουργία εγγράφου…"):
                st.session_state["perigrammata_full_report"] = _full_report(
                    curriculum, LOCALE, version
                )
        if st.session_state.get("perigrammata_full_report"):
            st.download_button(
                "Λήψη πλήρους αναφοράς",
                data=st.session_state["perigrammata_full_report"],
                file_name=f"Περιγράμματα-{curriculum}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        st.divider()
        st.markdown("### Αναφορά μεταβολών")
        st.caption(
            "Τι άλλαξε σε μια περίοδο, από το ιστορικό των περιγραμμάτων. "
            "Η αρχική φόρτωση δεν προσμετράται."
        )
        left, right, scope_col = st.columns(3)
        with left:
            start_date = st.date_input(
                "Από", value=date(date.today().year, 1, 1), format="DD/MM/YYYY"
            )
        with right:
            end_date = st.date_input("Έως", value=date.today(), format="DD/MM/YYYY")
        with scope_col:
            all_curricula = "Όλα"
            scope = st.selectbox(
                "Πρόγραμμα σπουδών", [curriculum, all_curricula], key="changes_scope"
            )

        if start_date > end_date:
            st.warning("Η ημερομηνία έναρξης είναι μετά την ημερομηνία λήξης.")
        else:
            # The end date is inclusive to the reader, so the query runs to the
            # start of the next day.
            start_at = datetime.combine(start_date, time.min, tzinfo=timezone.utc)
            end_at = datetime.combine(end_date, time.max, tzinfo=timezone.utc)
            changes = pdb.changes_between(
                start_at,
                end_at,
                curriculum=None if scope == all_curricula else curriculum,
                locale=LOCALE,
            )
            if changes.empty:
                st.info("Καμία μεταβολή στη συγκεκριμένη περίοδο.")
            else:
                st.dataframe(
                    changes[["code", "label", "edited_by", "edited_at", "note"]],
                    use_container_width=True,
                    hide_index=True,
                )
                names = dict(zip(courses["code"], courses["name"]))
                st.download_button(
                    "Λήψη αναφοράς μεταβολών",
                    data=report.build_changes_report(
                        changes,
                        start_at,
                        end_at,
                        curriculum=None if scope == all_curricula else curriculum,
                        names=names,
                    ),
                    file_name=f"Μεταβολές-περιγραμμάτων-{start_date:%Y%m%d}-{end_date:%Y%m%d}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
