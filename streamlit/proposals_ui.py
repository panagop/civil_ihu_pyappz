"""The tab where members propose changes for the year under preparation.

Kept out of the page module both because page 5 is already long and because
this is the only part of it that writes anything. Everything it needs about
the registry is passed in, so it holds no loading logic of its own.

Actions are collected in a form below the table rather than as buttons on every
row: a subject has ~26 electors, and three inline buttons each would mean ~80
widgets rebuilt on every rerun, for a worse layout.
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

import db

ID_COL = "Κωδικός Χρήστη"

ACTION_LABELS = {
    db.REMOVE: "Αφαίρεση εκλέκτορα",
    db.RECHARACTERIZE: "Αλλαγή χαρακτηρισμού",
    db.REJUSTIFY: "Διόρθωση αιτιολόγησης",
}
PROPOSAL_MARKS = {db.ADD: "➕", db.REMOVE: "➖", db.RECHARACTERIZE: "🔄", db.REJUSTIFY: "✏️"}

BLOCKED_MARK = "🔴"
PROPOSED_COL = "Εκκρεμεί"
FLAG_COL = "Σήμανση"


def _person_label(registry_by_id: dict, elector_id: int) -> str:
    person = registry_by_id.get(elector_id)
    if person is None:
        return f"#{elector_id} (εκτός μητρώου)"
    return f"{person['Επώνυμο']} {person['Όνομα']} (#{elector_id})"


def _blocked_ids(registry: pd.DataFrame, blocking_cols: list[str]) -> set[int]:
    if not blocking_cols:
        return set()
    mask = pd.Series(False, index=registry.index)
    for col in blocking_cols:
        mask |= registry[col].astype(str).str.strip() == "ΝΑΙ"
    return set(registry.loc[mask, ID_COL].astype("int64"))


def _decorate(
    table: pd.DataFrame,
    registry_by_id: dict,
    blocked: set[int],
    pending: pd.DataFrame,
) -> pd.DataFrame:
    """The subject's table with names, a κώλυμα flag and pending-change marks."""
    marks: dict[int, list[str]] = {}
    for row in pending.itertuples(index=False):
        marks.setdefault(int(row.elector_id), []).append(
            f"{PROPOSAL_MARKS.get(row.action, '•')} {row.action}"
        )

    records = []
    for row in table.itertuples(index=False):
        elector = int(row.elector_id)
        person = registry_by_id.get(elector, {})
        records.append(
            {
                FLAG_COL: BLOCKED_MARK if elector in blocked else "",
                "Χαρακτηρισμός": row.characterization,
                ID_COL: elector,
                "Επώνυμο": person.get("Επώνυμο", ""),
                "Όνομα": person.get("Όνομα", ""),
                "Βαθμίδα": person.get("Βαθμίδα", ""),
                "Φορέας": person.get("Φορέας", ""),
                PROPOSED_COL: " · ".join(marks.get(elector, [])),
                "Αιτιολόγηση συνάφειας": row.reasoning,
            }
        )
    return pd.DataFrame(records)


def _open_year_block(year: int, baseline_year: int, user_email: str, coordinator: bool):
    st.info(f"Το έτος {year} δεν έχει ανοίξει ακόμη.")
    if not coordinator:
        st.caption("Μόνο ο συντονιστής μπορεί να το ανοίξει.")
        return
    st.write(
        f"Το άνοιγμα δεν αντιγράφει εγγραφές: ο πίνακας του {year} θα είναι "
        f"ο πίνακας του {baseline_year} συν τις εγκεκριμένες προτάσεις."
    )
    if st.button(f"Άνοιγμα έτους {year} με βάση το {baseline_year}", type="primary"):
        st.success(db.open_year(year, baseline_year, user_email))
        st.rerun()


def _propose_change(year: int, field_code: int, table: pd.DataFrame,
                    registry_by_id: dict, user_email: str) -> None:
    if table.empty:
        st.caption("Δεν υπάρχουν εκλέκτορες σε αυτό το αντικείμενο.")
        return

    options = {
        _person_label(registry_by_id, int(row.elector_id)): row
        for row in table.itertuples(index=False)
    }
    with st.form(f"change_{field_code}"):
        label = st.selectbox("Εκλέκτορας", list(options))
        action = st.radio(
            "Ενέργεια", list(ACTION_LABELS), format_func=ACTION_LABELS.get,
            horizontal=True,
        )
        current = options[label]
        characterization = st.radio(
            "Νέος χαρακτηρισμός", db.CHARACTERIZATIONS, horizontal=True,
            index=db.CHARACTERIZATIONS.index(current.characterization),
            help="Χρησιμοποιείται μόνο στην αλλαγή χαρακτηρισμού.",
        )
        reasoning = st.text_area(
            "Νέα αιτιολόγηση συνάφειας", value=current.reasoning,
            help="Χρησιμοποιείται μόνο στη διόρθωση αιτιολόγησης.",
        )
        note = st.text_area("Αιτιολόγηση της μεταβολής *", placeholder="Γιατί;")

        if st.form_submit_button("Καταχώρηση πρότασης"):
            message = db.add_proposal(
                year=year,
                field_code=field_code,
                elector_id=int(current.elector_id),
                action=action,
                note=note,
                author=user_email,
                characterization=(
                    characterization if action == db.RECHARACTERIZE else None
                ),
                reasoning=reasoning if action == db.REJUSTIFY else None,
            )
            if message.startswith("Η πρόταση καταχωρήθηκε"):
                st.success(message)
                st.rerun()
            else:
                st.error(message)


def _propose_addition(year: int, field_code: int, registry: pd.DataFrame,
                      current_ids: set[int], blocked: set[int],
                      fold, user_email: str) -> None:
    query = st.text_input(
        "Αναζήτηση στο μητρώο", placeholder="επώνυμο ή γνωστικό αντικείμενο",
        key=f"add_search_{field_code}",
    )
    if not query.strip():
        st.caption("Γράψτε κάτι για να βρείτε υποψήφιους.")
        return

    haystack = fold(registry["Επώνυμο"]) + " " + fold(registry["Γνωστικό Αντικείμενο"])
    hits = registry[haystack.str.contains(fold(pd.Series([query]))[0], regex=False)]
    hits = hits[~hits[ID_COL].astype("int64").isin(current_ids)]
    if hits.empty:
        st.caption("Καμία εγγραφή — ή είναι ήδη στο αντικείμενο.")
        return

    st.caption(f"{len(hits)} αποτελέσματα, εμφανίζονται τα πρώτα 25.")
    options = {
        f"{'🔴 ' if int(row[ID_COL]) in blocked else ''}"
        f"{row['Επώνυμο']} {row['Όνομα']} — {row['Βαθμίδα']}, {row['Φορέας']} (#{int(row[ID_COL])})": int(row[ID_COL])
        for _, row in hits.head(25).iterrows()
    }
    with st.form(f"add_{field_code}"):
        label = st.selectbox("Υποψήφιος", list(options))
        characterization = st.radio(
            "Χαρακτηρισμός", db.CHARACTERIZATIONS, horizontal=True
        )
        reasoning = st.text_area("Αιτιολόγηση συνάφειας *")
        note = st.text_area("Αιτιολόγηση της μεταβολής *", placeholder="Γιατί προστίθεται;")
        if st.form_submit_button("Πρόταση προσθήκης"):
            if not reasoning.strip():
                st.error("Η αιτιολόγηση συνάφειας είναι υποχρεωτική.")
            else:
                message = db.add_proposal(
                    year=year, field_code=field_code, elector_id=options[label],
                    action=db.ADD, note=note, author=user_email,
                    characterization=characterization, reasoning=reasoning.strip(),
                )
                if message.startswith("Η πρόταση καταχωρήθηκε"):
                    st.success(message)
                    st.rerun()
                else:
                    st.error(message)


def _my_proposals(year: int, user_email: str, registry_by_id: dict) -> None:
    mine = db.list_proposals(year)
    if not mine.empty:
        mine = mine[mine["author"] == user_email]
    if mine.empty:
        st.caption("Δεν έχετε καταχωρήσει προτάσεις.")
        return

    view = mine.assign(
        Εκλέκτορας=[
            _person_label(registry_by_id, int(i)) for i in mine["elector_id"]
        ]
    )[["id", "field_code", "Εκλέκτορας", "action", "status", "note", "created_at"]]
    st.dataframe(view, use_container_width=True, hide_index=True)

    pending = mine[mine["status"] == db.PENDING]
    if pending.empty:
        return
    with st.form("withdraw"):
        choice = st.selectbox(
            "Απόσυρση εκκρεμούς πρότασης",
            pending["id"].tolist(),
            format_func=lambda i: (
                f"#{i} — {pending.loc[pending['id'] == i, 'action'].iat[0]}"
            ),
        )
        if st.form_submit_button("Απόσυρση"):
            st.info(db.withdraw_proposal(int(choice), user_email))
            st.rerun()


def _coordinator_block(year: int, registry_by_id: dict, user_email: str) -> None:
    pending = db.list_proposals(year, status=db.PENDING)
    st.metric("Εκκρεμείς προτάσεις", len(pending))

    if not pending.empty:
        view = pending.assign(
            Εκλέκτορας=[
                _person_label(registry_by_id, int(i)) for i in pending["elector_id"]
            ]
        )[["id", "field_code", "Εκλέκτορας", "action", "characterization",
           "note", "author", "created_at"]]
        st.dataframe(view, use_container_width=True, hide_index=True)

        with st.form("decide"):
            choice = st.selectbox(
                "Πρόταση", pending["id"].tolist(),
                format_func=lambda i: (
                    f"#{i} — {pending.loc[pending['id'] == i, 'action'].iat[0]} "
                    f"({pending.loc[pending['id'] == i, 'author'].iat[0]})"
                ),
            )
            decision_note = st.text_input("Σχόλιο απόφασης (προαιρετικό)")
            accept, reject = st.columns(2)
            if accept.form_submit_button("Έγκριση", type="primary"):
                st.success(db.decide_proposal(
                    int(choice), db.ACCEPTED, user_email, decision_note))
                st.rerun()
            if reject.form_submit_button("Απόρριψη"):
                st.warning(db.decide_proposal(
                    int(choice), db.REJECTED, user_email, decision_note))
                st.rerun()

    st.divider()
    st.subheader("Οριστικοποίηση")
    st.caption(
        "Γράφει τον πίνακα στο μητρώο του έτους και κλειδώνει. "
        "Δεν επιτρέπονται άλλες αλλαγές — δεν αναιρείται από την εφαρμογή."
    )
    confirmed = st.checkbox(f"Επιβεβαιώνω την οριστικοποίηση του {year}")
    if st.button("Οριστικοποίηση έτους", disabled=not confirmed):
        message = db.finalize_year(year, user_email)
        (st.success if "οριστικοποιήθηκε" in message else st.error)(message)
        if "οριστικοποιήθηκε" in message:
            st.rerun()


def render(*, year: int, baseline_year: int, registry: pd.DataFrame,
           antikeimena: pd.DataFrame, blocking_cols: list[str],
           fold, user_email: str) -> None:
    """Draw the whole tab. `fold` is page 5's fold_greek_series."""
    if not db.is_available():
        st.warning(
            "Η βάση δεδομένων δεν είναι διαθέσιμη σε αυτό το περιβάλλον "
            "(είναι προσβάσιμη μόνο από το Railway)."
        )
        return

    coordinator = db.is_coordinator(user_email)
    state = db.year_state(year)
    if state is None:
        _open_year_block(year, baseline_year, user_email, coordinator)
        return

    locked = state["status"] == db.LOCKED
    if locked:
        st.success(
            f"Το έτος {year} είναι **κλειδωμένο** "
            f"(από {state['locked_by']}, {state['locked_at']:%d/%m/%Y})."
        )
    else:
        st.caption(
            f"Έτος {year} · ανοιχτό · βάση: {state['baseline_year']} · "
            f"συνδεδεμένος ως {user_email}"
            + (" · **συντονιστής**" if coordinator else "")
        )

    registry = registry.copy()
    registry[ID_COL] = registry[ID_COL].astype("int64")
    registry_by_id = registry.set_index(ID_COL).to_dict("index")
    blocked = _blocked_ids(registry, blocking_cols)

    table = db.working_electors(year)
    labels = {
        f"{row.Code} — {row.field}": int(row.Code)
        for row in antikeimena.sort_values("Code").itertuples(index=False)
    }
    label = st.selectbox("Γνωστικό αντικείμενο", list(labels), key="prop_field")
    field_code = labels[label]

    subject = table[table["field_code"] == field_code] if not table.empty else table
    pending = db.list_proposals(year, field_code=field_code, status=db.PENDING)

    counts = subject["characterization"].value_counts() if not subject.empty else {}
    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("Σύνολο", len(subject))
    col_b.metric("Ιδίου", int(counts.get("ΙΔΙΟΥ", 0)) if len(counts) else 0)
    col_c.metric("Συναφούς", int(counts.get("ΣΥΝΑΦΟΥΣ", 0)) if len(counts) else 0)
    blocked_here = (
        int(subject["elector_id"].astype("int64").isin(blocked).sum())
        if not subject.empty else 0
    )
    col_d.metric("Με κώλυμα", blocked_here, delta=None)

    if blocked_here:
        st.warning(
            f"{blocked_here} εκλέκτορες έχουν κώλυμα αποκλεισμού από τα μητρώα "
            "και πρέπει να αφαιρεθούν ή να τεκμηριωθεί η παραμονή τους."
        )

    st.dataframe(
        _decorate(subject, registry_by_id, blocked, pending),
        use_container_width=True, hide_index=True,
    )

    if locked:
        return

    tab_change, tab_add, tab_mine = st.tabs(
        ["Μεταβολή", "Προσθήκη", "Οι προτάσεις μου"]
    )
    with tab_change:
        _propose_change(year, field_code, subject, registry_by_id, user_email)
    with tab_add:
        current_ids = (
            set(subject["elector_id"].astype("int64")) if not subject.empty else set()
        )
        _propose_addition(
            year, field_code, registry, current_ids, blocked, fold, user_email
        )
    with tab_mine:
        _my_proposals(year, user_email, registry_by_id)

    if coordinator:
        st.divider()
        st.subheader("Συντονιστής")
        _coordinator_block(year, registry_by_id, user_email)
