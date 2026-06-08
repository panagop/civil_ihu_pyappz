from pathlib import Path

import pandas as pd
import streamlit as st


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
