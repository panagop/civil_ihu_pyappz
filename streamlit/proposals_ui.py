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
    db.MODIFY: "Μεταβολή χαρακτηρισμού / αιτιολόγησης",
    db.REMOVE: "Αφαίρεση εκλέκτορα",
    db.RECHARACTERIZE: "Αλλαγή χαρακτηρισμού",
    db.REJUSTIFY: "Διόρθωση αιτιολόγησης",
}
PROPOSAL_MARKS = {
    db.ADD: "➕", db.REMOVE: "➖", db.MODIFY: "🔄",
    db.RECHARACTERIZE: "🔄", db.REJUSTIFY: "✏️",
}

BLOCKED_MARK = "🔴"
CHANGED_MARK = "🟡"
PROPOSED_COL = "Εκκρεμεί"
FLAG_COL = "Σήμανση"
REGISTRY_CHANGE_COL = "Μεταβολές μητρώου"

# What is compared between the baseline year's registry and the current one, as
# (registry column, label). Κατηγορία Χρήστη is deliberately absent: the exports
# relabelled it, which would flag almost everybody.
TRACKED_FIELDS = [
    ("Βαθμίδα", "Βαθμίδα"),
    ("Γνωστικό Αντικείμενο", "Γνωστικό αντικείμενο"),
    ("Φορέας", "Φορέας"),
]


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


def registry_changes(
    baseline: pd.DataFrame, current: pd.DataFrame, fold
) -> dict[int, str]:
    """Per elector, a description of what moved between the two registries.

    Only electors present in both are compared; someone who left the registry
    is caught by the κώλυμα/absence checks instead.

    Values are compared through ``fold`` (fold_greek_series), not casefold:
    the exports re-typed several subjects in title case, and "ΔΥΝΑΜΙΚΗ" vs
    "Δυναμική" differs by an accent that casefold keeps. Those are not changes
    anyone needs to look at.
    """
    if baseline is None or baseline.empty or current.empty:
        return {}
    old = baseline.copy()
    old[ID_COL] = old[ID_COL].astype("int64")
    old = old.set_index(ID_COL)
    new = current.set_index(ID_COL)
    shared = old.index.intersection(new.index)

    changes: dict[int, list[str]] = {}
    for column, label in TRACKED_FIELDS:
        if column not in old.columns or column not in new.columns:
            continue
        before = old.loc[shared, column].fillna("").astype(str).str.strip()
        after = new.loc[shared, column].fillna("").astype(str).str.strip()
        differs = (fold(before) != fold(after)).to_numpy()
        for elector in shared[differs]:
            changes.setdefault(int(elector), []).append(
                f"{label}: {before[elector]} → {after[elector]}"
            )
    return {elector: " · ".join(items) for elector, items in changes.items()}


def _decorate(
    table: pd.DataFrame,
    registry_by_id: dict,
    blocked: set[int],
    pending: pd.DataFrame,
    changes: dict[int, str],
) -> pd.DataFrame:
    """The subject's table with names, flags, registry changes and pending marks."""
    marks: dict[int, list[str]] = {}
    for row in pending.itertuples(index=False):
        marks.setdefault(int(row.elector_id), []).append(
            f"{PROPOSAL_MARKS.get(row.action, '•')} {row.action}"
        )

    records = []
    for row in table.itertuples(index=False):
        elector = int(row.elector_id)
        person = registry_by_id.get(elector, {})
        change = changes.get(elector, "")
        flag = BLOCKED_MARK if elector in blocked else (CHANGED_MARK if change else "")
        records.append(
            {
                FLAG_COL: flag,
                "Χαρακτηρισμός": row.characterization,
                ID_COL: elector,
                "Επώνυμο": person.get("Επώνυμο", ""),
                "Όνομα": person.get("Όνομα", ""),
                "Βαθμίδα": person.get("Βαθμίδα", ""),
                "Φορέας": person.get("Φορέας", ""),
                REGISTRY_CHANGE_COL: change,
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
    # Deliberately OUTSIDE the form. A widget inside st.form does not rerun
    # until submit, so the fields below would keep showing the values of
    # whoever was selected before.
    label = st.selectbox("Εκλέκτορας", list(options), key=f"chg_who_{field_code}")
    current = options[label]
    elector = int(current.elector_id)

    action = st.radio(
        "Ενέργεια", [db.MODIFY, db.REMOVE], horizontal=True,
        format_func=ACTION_LABELS.get, key=f"chg_action_{field_code}",
    )

    # The elector is part of every key: Streamlit keeps the stored value of a
    # widget whose key is unchanged, which would defeat the new defaults.
    with st.form(f"change_{field_code}_{elector}_{action}"):
        characterization, reasoning = None, None
        if action == db.MODIFY:
            st.caption(
                "Αλλάξτε τον χαρακτηρισμό, την αιτιολόγηση ή και τα δύο. "
                "Καταχωρείται ως μία πρόταση, ώστε να εγκριθεί ενιαία."
            )
            characterization = st.radio(
                "Χαρακτηρισμός", db.CHARACTERIZATIONS, horizontal=True,
                index=db.CHARACTERIZATIONS.index(current.characterization),
                key=f"chg_char_{field_code}_{elector}",
            )
            reasoning = st.text_area(
                "Αιτιολόγηση συνάφειας", value=current.reasoning, height=140,
                key=f"chg_reason_{field_code}_{elector}",
            )
        else:
            st.warning(f"Πρόταση αφαίρεσης: **{label}**")

        note = st.text_area(
            "Αιτιολόγηση της μεταβολής *", placeholder="Γιατί;",
            key=f"chg_note_{field_code}_{elector}_{action}",
        )

        if st.form_submit_button("Καταχώρηση πρότασης"):
            unchanged = action == db.MODIFY and (
                characterization == current.characterization
                and (reasoning or "").strip() == (current.reasoning or "").strip()
            )
            if unchanged:
                st.error("Δεν αλλάξατε τίποτα.")
                return
            message = db.add_proposal(
                year=year,
                field_code=field_code,
                elector_id=elector,
                action=action,
                note=note,
                author=user_email,
                characterization=characterization,
                reasoning=(reasoning or "").strip() or None,
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


def _bulk_block(year: int, field_code: int, field_label: str, user_email: str) -> None:
    """Decide every pending proposal of the subject on screen, in one go."""
    here = db.list_proposals(year, field_code=field_code, status=db.PENDING)
    if here.empty:
        st.caption("Καμία εκκρεμής πρόταση σε αυτό το αντικείμενο.")
        return

    counts = here["action"].value_counts().to_dict()
    st.write(
        f"**{len(here)} εκκρεμείς** στο «{field_label}» — "
        + ", ".join(f"{count} × {action}" for action, count in counts.items())
    )
    note = st.text_input(
        "Σχόλιο απόφασης (προαιρετικό)", key=f"bulk_note_{field_code}"
    )
    confirmed = st.checkbox(
        f"Επιβεβαιώνω τη μαζική απόφαση για {len(here)} προτάσεις",
        key=f"bulk_ok_{field_code}",
    )
    accept, reject = st.columns(2)
    if accept.button(
        "Έγκριση όλων", type="primary", disabled=not confirmed,
        key=f"bulk_yes_{field_code}",
    ):
        st.success(db.decide_field_proposals(
            year, field_code, db.ACCEPTED, user_email, note))
        st.rerun()
    if reject.button(
        "Απόρριψη όλων", disabled=not confirmed, key=f"bulk_no_{field_code}"
    ):
        st.warning(db.decide_field_proposals(
            year, field_code, db.REJECTED, user_email, note))
        st.rerun()


def _coordinator_block(year: int, field_code: int, field_label: str,
                       registry_by_id: dict, user_email: str) -> None:
    pending = db.list_proposals(year, status=db.PENDING)
    st.metric("Εκκρεμείς προτάσεις (όλο το έτος)", len(pending))

    per_field = db.pending_by_field(year)
    if not per_field.empty:
        with st.expander(
            f"Εκκρεμότητες ανά αντικείμενο ({len(per_field)} αντικείμενα)"
        ):
            st.dataframe(
                per_field.rename(
                    columns={"field_code": "Κωδικός", "pending": "Εκκρεμείς"}
                ),
                use_container_width=True, hide_index=True,
            )

    st.markdown("##### Μαζική απόφαση για το τρέχον αντικείμενο")
    _bulk_block(year, field_code, field_label, user_email)
    st.markdown("##### Απόφαση ανά πρόταση")

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
           baseline_registry: pd.DataFrame | None, antikeimena: pd.DataFrame,
           blocking_cols: list[str], fold, user_email: str) -> None:
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
    changes = registry_changes(baseline_registry, registry, fold)

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
    col_d.metric("Με κώλυμα", blocked_here)

    if blocked_here:
        st.warning(
            f"{BLOCKED_MARK} {blocked_here} εκλέκτορες έχουν κώλυμα αποκλεισμού "
            "από τα μητρώα και πρέπει να αφαιρεθούν ή να τεκμηριωθεί η παραμονή τους."
        )
    changed_here = (
        int(subject["elector_id"].astype("int64").isin(changes).sum())
        if not subject.empty else 0
    )
    if changed_here:
        st.info(
            f"{CHANGED_MARK} {changed_here} εκλέκτορες άλλαξαν βαθμίδα, γνωστικό "
            f"αντικείμενο ή φορέα από το {baseline_year}. Δείτε τη στήλη "
            f"«{REGISTRY_CHANGE_COL}» — συνήθως δεν απαιτείται ενέργεια."
        )

    st.dataframe(
        _decorate(subject, registry_by_id, blocked, pending, changes),
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
        _coordinator_block(year, field_code, label, registry_by_id, user_email)
