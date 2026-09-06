"""Shape stored decisions into the submitted table's layout.

The database keeps only the decisions (χαρακτηρισμός + αιτιολόγηση); every
column describing a person is joined in from an ΑΠΕΛΛΑ export. This module owns
that join and the column order, so the tab that browses a finalised year, the
preview of a year being prepared, and the Word report all produce **the same
table** — the one that gets submitted to the department.

Which export to join against is the caller's decision: a finalised year uses
its own snapshot (the electors as they stood when it was submitted), while a
year under preparation uses the current one.
"""

from __future__ import annotations

import pandas as pd

ID_COL = "Κωδικός Χρήστη"
CHARAKTIRISMOS_COL = "Χαρακτηρισμός"
SUBJECT_COL = "Γνωστικό Αντικείμενο"
REASONING_COL = "Αιτιολόγηση συνάφειας"

# The registry export names three columns differently from the workbook
REGISTRY_TO_WORKBOOK = {
    "Φορέας": "Φορέας Χρήστη",
    "Σχολή": "Σχολή Χρήστη",
    "Τμήμα/Ινστιτούτο": "Τμήμα/Ινστιτούτο Χρήστη",
}

# The submitted workbooks' column order
WORKBOOK_COLUMNS = [
    "α/α",
    CHARAKTIRISMOS_COL,
    ID_COL,
    "Όνομα",
    "Επώνυμο",
    "Κατηγορία Χρήστη",
    "Φορέας Χρήστη",
    "Σχολή Χρήστη",
    "Τμήμα/Ινστιτούτο Χρήστη",
    "ΦΕΚ Διορισμού",
    SUBJECT_COL,
    "Βαθμίδα",
    REASONING_COL,
]


def prepare_registry(registry: pd.DataFrame | None) -> pd.DataFrame:
    """Registry export renamed to workbook column names, keyed by elector id."""
    if registry is None or registry.empty:
        return pd.DataFrame(columns=[ID_COL])
    people = registry.rename(columns=REGISTRY_TO_WORKBOOK).copy()
    people[ID_COL] = people[ID_COL].astype("int64")
    return people


def build(decisions: pd.DataFrame, people: pd.DataFrame, fold) -> pd.DataFrame:
    """One subject's decisions as the submitted table.

    ``decisions`` carries the database columns (elector_id, characterization,
    reasoning); ``people`` comes from :func:`prepare_registry`. Electors missing
    from the export keep their row with blank columns — dropping them would
    silently shrink the table.

    Rows are ordered ΙΔΙΟΥ first then by surname, folded so that accents and
    case do not scatter the alphabet, and ``α/α`` is numbered from that order
    rather than stored.
    """
    if decisions.empty:
        return pd.DataFrame(columns=WORKBOOK_COLUMNS)

    renamed = decisions.rename(
        columns={
            "characterization": CHARAKTIRISMOS_COL,
            "reasoning": REASONING_COL,
            "elector_id": ID_COL,
        }
    )
    renamed[ID_COL] = renamed[ID_COL].astype("int64")
    merged = renamed.merge(people, on=ID_COL, how="left")

    merged = merged.sort_values(
        [CHARAKTIRISMOS_COL, "Επώνυμο", "Όνομα"],
        key=lambda col: (
            col.ne("ΙΔΙΟΥ") if col.name == CHARAKTIRISMOS_COL else fold(col)
        ),
    ).reset_index(drop=True)
    merged.insert(0, "α/α", range(1, len(merged) + 1))

    columns = [name for name in WORKBOOK_COLUMNS if name in merged.columns]
    return merged[columns].fillna("")
