from pathlib import Path

import pandas as pd
import streamlit as st


def load_data(input_excel: Path, sheet_name: str, teaching_period: str) -> pd.DataFrame:
    """Διαβάζει τα δεδομένα του εβδομαδιαίου προγράμματος από το Excel."""

    if not input_excel.exists():
        st.error(f"❌ Το αρχείο {input_excel} δεν βρέθηκε!")
        st.info(f"Αναζητούμενη διαδρομή: {input_excel.absolute()}")
        st.stop()

    try:
        excel_file = pd.ExcelFile(input_excel)
        available_sheets = excel_file.sheet_names

        if sheet_name not in available_sheets:
            st.error(f"❌ Το sheet '{sheet_name}' δεν βρέθηκε στο αρχείο!")
            st.info(f"Διαθέσιμα sheets: {', '.join(available_sheets)}")
            st.stop()

        df = pd.read_excel(input_excel, sheet_name=sheet_name)

        required_cols = [
            "course_id",
            "course_name",
            "class_name",
            "semester",
            "teaching_period",
            "instructors",
            "day",
            "start_time",
            "duration",
            "room",
            "notes",
        ]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            raise ValueError(f"Λείπουν οι στήλες: {missing}")

        df = df[df['teaching_period'] == teaching_period]

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

        df['full_class_name'] = df.apply(
            lambda row: f"{row['course_name']} - {row['class_name']}"
            if pd.notna(row['class_name']) else str(row['course_name']),
            axis=1,
        )

        df['start_hour'] = df['start_time'].apply(
            lambda x: x.hour if hasattr(x, 'hour') else int(x)
        )

        df['end_hour'] = df['start_hour'] + df['duration']
        df['end_time'] = df.apply(
            lambda row: f"{int(row['end_hour'])}:00",
            axis=1,
        )

    except Exception as e:
        st.error(f"❌ Σφάλμα κατά το άνοιγμα του αρχείου: {e}")
        st.stop()

    return df
