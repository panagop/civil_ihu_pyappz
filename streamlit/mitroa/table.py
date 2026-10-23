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

import io
import re
import zipfile

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

# The per-subject CSV layout (files/mitroa/csv_test/ holds a sample):
# id, surname, 1 for ΙΔΙΟΥ / 0 for ΣΥΝΑΦΟΥΣ.
CSV_HEADER = ["ΚΩΔΙΚΟΣ ΧΡΗΣΤΗ", "ΕΠΩΝΥΜΟ", "ΣΥΝΑΦΕΙΑ"]
CSV_SYNAFEIA = {"ΙΔΙΟΥ": 1, "ΣΥΝΑΦΟΥΣ": 0}
# Characters Windows refuses in a filename; subject titles carry ":" and "/"
UNSAFE_FILENAME_RE = re.compile(r'[\\/:*?"<>|]')


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


def subject_csv(df: pd.DataFrame) -> bytes:
    """One subject's table in the CSV layout: UTF-8 without BOM, CRLF, as the sample.

    Rows keep the table's order. A χαρακτηρισμός other than ΙΔΙΟΥ / ΣΥΝΑΦΟΥΣ
    raises rather than becoming a 0 nobody would notice.
    """
    synafeia = df[CHARAKTIRISMOS_COL].astype(str).str.strip().map(CSV_SYNAFEIA)
    if synafeia.isna().any():
        bad = sorted(set(df.loc[synafeia.isna(), CHARAKTIRISMOS_COL].astype(str)))
        raise ValueError(f"άγνωστος χαρακτηρισμός: {', '.join(bad)}")
    out = pd.DataFrame(
        {
            CSV_HEADER[0]: pd.to_numeric(df[ID_COL]).astype("int64"),
            CSV_HEADER[1]: df["Επώνυμο"].astype(str).str.strip(),
            CSV_HEADER[2]: synafeia.astype("int64"),
        }
    )
    return out.to_csv(index=False, lineterminator="\r\n").encode("utf-8")


def csv_filename(field: str) -> str:
    """The subject title as a filename, with the characters Windows refuses replaced."""
    name = UNSAFE_FILENAME_RE.sub("-", str(field)).strip().rstrip(".").strip()
    return f"{name}.csv"


def csv_archive(workbook: dict[str, dict]) -> bytes:
    """Every subject of a year as ``<title>.csv``, zipped.

    ``workbook`` is the structure the page builds from either source
    ({sheet: {code, field, domain, df}}). Raises ValueError naming the subject
    whose table cannot be expressed in the layout.
    """
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for entry in sorted(workbook.values(), key=lambda e: str(e["code"])):
            try:
                content = subject_csv(entry["df"])
            except (KeyError, ValueError) as exc:
                raise ValueError(f"{entry['code']} — {entry['field']}: {exc}") from exc
            archive.writestr(csv_filename(entry["field"]), content)
    return buffer.getvalue()
