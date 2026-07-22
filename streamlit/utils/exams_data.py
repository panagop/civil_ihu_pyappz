import re
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

EXAM_FILE_RE = re.compile(r"exams-(\d{4})-(\d{2})\.xlsm$", re.IGNORECASE)

GREEK_MONTHS_GENITIVE = {
    1: "Ιανουαρίου",
    2: "Φεβρουαρίου",
    3: "Μαρτίου",
    4: "Απριλίου",
    5: "Μαΐου",
    6: "Ιουνίου",
    7: "Ιουλίου",
    8: "Αυγούστου",
    9: "Σεπτεμβρίου",
    10: "Οκτωβρίου",
    11: "Νοεμβρίου",
    12: "Δεκεμβρίου",
}


def _period_name(month: int) -> str:
    """Greek name of the exam period implied by its month."""
    if month in (1, 2):
        return "Χειμερινό"
    if month in (6, 7):
        return "Εαρινό"
    if month in (8, 9):
        return "Επαναληπτική Σεπτεμβρίου"
    return "Εξεταστική"


def _academic_year(year: int, month: int) -> str:
    """Academic year string (e.g. "2025-2026"); the year rolls over in October.

    September (month 9) still belongs to the *previous* academic year: the
    "Επαναληπτική Σεπτεμβρίου" resit period is the final exam period of that
    year, held just before the new academic year starts in October.
    """
    start = year if month >= 10 else year - 1
    return f"{start}-{start + 1}"


@st.cache_data(show_spinner=False)
def _exam_date_range(input_excel_str: str) -> tuple[date | None, date | None]:
    """(min, max) exam_date in the file, read from the first sheet that has it."""
    path = Path(input_excel_str)
    try:
        excel_file = pd.ExcelFile(path)
    except Exception:
        return None, None

    sheets = list(dict.fromkeys(["ΔΙΠΑΕ", "ΤΕΙ", *excel_file.sheet_names]))
    for sheet in sheets:
        if sheet not in excel_file.sheet_names:
            continue
        try:
            df = pd.read_excel(path, sheet_name=sheet)
        except Exception:
            continue
        if "exam_date" not in df.columns:
            continue
        dates = pd.to_datetime(df["exam_date"], errors="coerce").dropna()
        if dates.empty:
            continue
        return dates.min().date(), dates.max().date()

    return None, None


def discover_exam_periods(exams_dir: Path) -> list[dict]:
    """Find ``exams-yyyy-mm.xlsm`` files and return metadata sorted chronologically."""
    periods: list[dict] = []
    for path in exams_dir.glob("exams-*.xlsm"):
        match = EXAM_FILE_RE.search(path.name)
        if not match:
            continue
        year, month = int(match.group(1)), int(match.group(2))
        start_date, end_date = _exam_date_range(str(path))
        name = _period_name(month)
        academic_year = _academic_year(year, month)
        periods.append(
            {
                "path": path,
                "year": year,
                "month": month,
                "name": name,
                "academic_year": academic_year,
                "label": f"{name} {academic_year} ({GREEK_MONTHS_GENITIVE[month]} {year})",
                "start_date": start_date,
                "end_date": end_date,
            }
        )

    periods.sort(key=lambda p: (p["year"], p["month"]))
    return periods


def default_period_index(periods: list[dict], today: date | None = None) -> int:
    """Index of the period to preselect.

    Preference order: the period currently in progress (today within its exam-date
    range), otherwise the next upcoming period, otherwise the most recent one.
    """
    if not periods:
        return 0
    today = today or date.today()

    # Period currently in progress.
    for idx, p in enumerate(periods):
        start, end = p["start_date"], p["end_date"]
        if start and end and start <= today <= end:
            return idx

    # Next upcoming period (earliest start date in the future).
    upcoming = [
        (p["start_date"], idx)
        for idx, p in enumerate(periods)
        if p["start_date"] and p["start_date"] > today
    ]
    if upcoming:
        return min(upcoming)[1]

    # Fallback: most recent period (list is sorted chronologically).
    return len(periods) - 1


def load_data(
    input_excel: Path,
    input_sheet: str,
    extra_column: str | None = None,
) -> pd.DataFrame:
    """Διαβάζει τα δεδομένα από το Excel (sheet επιλεγμένου προγράμματος σπουδών).

    Optionally keeps an extra sheet-specific column (e.g. "students_total" for ΔΙΠΑΕ
    or "φοιτΤΕΙ" for ΤΕΙ).
    """

    if not input_excel.exists():
        st.error(f"❌ Το αρχείο {input_excel} δεν βρέθηκε!")
        st.info(f"Αναζητούμενη διαδρομή: {input_excel.absolute()}")
        st.stop()

    try:
        excel_file = pd.ExcelFile(input_excel)
        available_sheets = excel_file.sheet_names

        if input_sheet not in available_sheets:
            st.error(f"❌ Το sheet '{input_sheet}' δεν βρέθηκε στο αρχείο!")
            st.info(f"Διαθέσιμα sheets: {', '.join(available_sheets)}")
            st.stop()

        df = pd.read_excel(input_excel, sheet_name=input_sheet)
    except Exception as e:
        st.error(f"❌ Σφάλμα κατά το άνοιγμα του αρχείου: {e}")
        st.stop()

    required_cols = [
        "course_id",
        "course_name",
        "semester",
        "instructor",
        "exam_date",
        "start_time",
        "room",
        "notes",
        "epitirites",
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Λείπουν οι στήλες: {missing}")

    if extra_column and extra_column not in df.columns:
        raise ValueError(f"Λείπει η στήλη '{extra_column}' στο sheet '{input_sheet}'")

    keep_cols = required_cols + ([extra_column] if extra_column else [])
    df = df[keep_cols]

    df["exam_date"] = pd.to_datetime(df["exam_date"]).dt.date
    df = df.dropna(subset=["exam_date"])

    day_names_greek = {
        0: 'Δευτέρα',
        1: 'Τρίτη',
        2: 'Τετάρτη',
        3: 'Πέμπτη',
        4: 'Παρασκευή',
        5: 'Σάββατο',
        6: 'Κυριακή',
    }
    df["day_of_week"] = pd.to_datetime(df["exam_date"]).dt.dayofweek.map(day_names_greek)

    df["start_time"] = df["start_time"].astype(str)

    # room / course_id can mix numeric codes (e.g. 101) and strings (e.g. "ΔΟΜ704");
    # normalize to string so pyarrow doesn't infer int64 and fail on the strings.
    def _to_str(v: object) -> str:
        if pd.isna(v):
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)

    df["room"] = df["room"].apply(_to_str)
    df["course_id"] = df["course_id"].apply(_to_str)

    df["start_dt"] = pd.to_datetime(
        df["exam_date"].astype(str) + " " + df["start_time"],
        errors="coerce",
    )

    df["iso_week_number"] = df["start_dt"].dt.isocalendar().week

    unique_weeks = sorted(df["iso_week_number"].unique())
    week_mapping = {iso_week: idx + 1 for idx, iso_week in enumerate(unique_weeks)}
    df["week_number"] = df["iso_week_number"].map(week_mapping)

    return df


def reload() -> None:
    """Clear cache to force reload."""
    st.cache_data.clear()
