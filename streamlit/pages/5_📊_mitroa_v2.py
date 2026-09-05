import io
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Μητρώα v2",
    layout="wide",
)

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from auth import require_ihu_login  # noqa: E402

require_ihu_login()

ROOT = Path(__file__).resolve().parents[2]
PROFESSORS_DIR = ROOT / "files" / "mitroa" / "professors_tables"
ANTIKEIMENA_CSV = ROOT / "files" / "mitroa" / "antikeimena.csv"
BY_YEAR_DIR = ROOT / "files" / "mitroa" / "mitroa_by_year"

# Layout of every worksheet in the external_<year>.xlsx workbooks (0-based)
META_ROW = 5           # row holding Κωδικός / Γνωστικό αντικείμενο / Επιστημονικό πεδίο
META_CODE_COL = 4
META_FIELD_COL = 5
META_DOMAIN_COL = 10
HEADER_ROW = 8         # row holding the α/α ... Αιτιολόγηση συνάφειας table header
CHARAKTIRISMOS_COL = "Χαρακτη-ρισμός"
# The workbooks spell the two characterisations inconsistently
CHARAKTIRISMOS_ALIASES = {"ΙΔΙΟ": "ΙΔΙΟΥ", "ΣΥΝΑΦΕΣ": "ΣΥΝΑΦΟΥΣ"}

ID_COL = "Κωδικός Χρήστη"
SUBJECT_COL = "Γνωστικό Αντικείμενο"
# Fields compared between the year tables and the registry, as
# (column in the external tables, column in the registry, label)
COMPARED_FIELDS = [
    ("Βαθμίδα", "Βαθμίδα", "Βαθμίδα"),
    ("Φορέας Χρήστη", "Φορέας", "Φορέας"),
    ("Γνωστικό Αντικείμενο", "Γνωστικό Αντικείμενο", "Γνωστικό αντικείμενο"),
]
# Κατηγορία Χρήστη is deliberately NOT compared: the exports relabelled it
# ("Ημεδαπής" -> "Καθηγητής Ημεδαπής"), which would swamp the real findings.
# Only exclusion from the μητρώα blocks an elector. Exclusion from
# εκλεκτορικά/επιτροπές is reported for information but does not disqualify.
# ("Σε αναστολή" is the equivalent column in the 2024/2025 exports.)
BLOCKING_FLAG_KEYWORDS = ["αποκλεισμου απο μητρωα", "σε αναστολη"]
# How multiple keywords combine in the keyword search tab
MATCH_ANY = "Οποιαδήποτε λέξη (OR)"
MATCH_ALL = "Όλες οι λέξεις (AND)"
# Marks electors absent from the registry being compared against
NEW_COL = "Νέος"

STATUS_MISSING = "🔴 Εκτός μητρώου"
STATUS_KOLYMA = "🟠 Κώλυμα"
STATUS_CHANGED = "🟡 Μεταβολή"
STATUS_OK = "🟢 Χωρίς μεταβολή"
STATUS_ORDER = [STATUS_MISSING, STATUS_KOLYMA, STATUS_CHANGED, STATUS_OK]

# Columns searched by the free-text box (only those present in the file are used)
SEARCH_COLS = [
    "Επώνυμο",
    "Όνομα",
    "Φορέας",
    "Σχολή",
    "Τμήμα/Ινστιτούτο",
    "Γνωστικό Αντικείμενο",
]
# Columns offered as multiselect filters, in the order they appear in the UI
FILTER_COLS = ["Κατηγορία Χρήστη", "Βαθμίδα", "Φορέας"]


@st.cache_data
def list_professor_files() -> dict[str, Path]:
    """Map 'year' -> parquet path, newest year first.

    Filenames look like ``professors_export_20260904.parquet``; the first four
    digits of the date stamp are the year of the μητρώο.
    """
    files: dict[str, Path] = {}
    for path in sorted(PROFESSORS_DIR.glob("professors_export_*.parquet"), reverse=True):
        match = re.search(r"(\d{4})(\d{2})(\d{2})", path.stem)
        label = f"{match.group(1)} ({match.group(3)}/{match.group(2)})" if match else path.stem
        files[label] = path
    return files


@st.cache_data
def load_professors(path_str: str) -> pd.DataFrame:
    return pd.read_parquet(path_str)


@st.cache_data
def load_antikeimena() -> pd.DataFrame:
    return pd.read_csv(ANTIKEIMENA_CSV)


@st.cache_data
def list_external_files() -> dict[str, Path]:
    """Map year label -> external_<year>.xlsx path, newest first."""
    files: dict[str, Path] = {}
    for path in sorted(BY_YEAR_DIR.glob("external_*.xlsx"), reverse=True):
        files[path.stem.replace("external_", "")] = path
    return files


@st.cache_data
def load_external_workbook(path_str: str) -> dict[str, dict]:
    """Parse every worksheet into {sheet_name: {code, field, domain, df}}.

    Each sheet holds one γνωστικό αντικείμενο: a small metadata block on top and
    the electors table starting at ``HEADER_ROW``.
    """
    sheets = pd.read_excel(path_str, sheet_name=None, header=None)
    parsed: dict[str, dict] = {}
    for name, raw in sheets.items():
        header = raw.iloc[HEADER_ROW].tolist()
        df = raw.iloc[HEADER_ROW + 1:].copy()
        df.columns = header
        df = df.dropna(how="all").dropna(axis=1, how="all")
        # The α/α column is a spreadsheet formula; renumber it ourselves
        if "α/α" in df.columns:
            df = df.drop(columns="α/α")
        df = df.reset_index(drop=True)
        df.insert(0, "α/α", range(1, len(df) + 1))
        if CHARAKTIRISMOS_COL in df.columns:
            df[CHARAKTIRISMOS_COL] = (
                df[CHARAKTIRISMOS_COL]
                .astype(str)
                .str.strip()
                .replace(CHARAKTIRISMOS_ALIASES)
            )
        parsed[name] = {
            "code": raw.iat[META_ROW, META_CODE_COL],
            "field": raw.iat[META_ROW, META_FIELD_COL],
            "domain": raw.iat[META_ROW, META_DOMAIN_COL],
            "df": df.fillna(""),
        }
    return parsed


def normalize(series: pd.Series) -> pd.Series:
    """Casefolded, whitespace-collapsed text, for comparing values across years."""
    return (
        series.fillna("")
        .astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .str.casefold()
    )


def fold_greek(text: str) -> str:
    """Accent-, case- and final-sigma-insensitive form of a Greek string.

    Plain ``casefold()`` is not enough: "σκυροδέμ" would not match
    "Σκυρόδεμα" (the accent sits on a different vowel) and "ς" does not
    casefold to "σ".
    """
    decomposed = unicodedata.normalize("NFD", str(text))
    stripped = "".join(c for c in decomposed if not unicodedata.combining(c))
    return stripped.casefold().replace("ς", "σ")


# Built from code points rather than written as "̀-ͯ": the parquet
# columns are Arrow-backed and pandas runs their regexes through RE2, which
# rejects \u escapes.
COMBINING_MARKS_RE = f"[{chr(0x0300)}-{chr(0x036F)}]"


def fold_greek_series(series: pd.Series) -> pd.Series:
    """Vectorised :func:`fold_greek` for a whole column."""
    return (
        series.fillna("")
        .astype(str)
        .str.normalize("NFD")
        .str.replace(COMBINING_MARKS_RE, "", regex=True)
        .str.casefold()
        .str.replace("ς", "σ", regex=False)
    )


def split_flag_columns(df: pd.DataFrame) -> tuple[list[str], list[str]]:
    """Registry ΝΑΙ/ΟΧΙ flag columns, split into (blocking, informational).

    Flags are found by their values rather than their names, because the column
    names change between yearly exports. Whether a flag disqualifies an elector
    is then decided by name against BLOCKING_FLAG_KEYWORDS.
    """
    blocking, informational = [], []
    for col in df.columns:
        values = set(df[col].dropna().astype(str).str.strip().unique())
        if not values or not values <= {"ΝΑΙ", "ΟΧΙ"}:
            continue
        folded = fold_greek(col)
        if any(k in folded for k in BLOCKING_FLAG_KEYWORDS):
            blocking.append(col)
        else:
            informational.append(col)
    return blocking, informational


@st.cache_data
def registry_ids(registry_path_str: str) -> set[int]:
    """The Κωδικοί Χρήστη present in a registry export."""
    return set(load_professors(registry_path_str)[ID_COL])


@st.cache_data
def folded_subjects(registry_path_str: str) -> pd.Series:
    """The Γνωστικό Αντικείμενο column of a registry, folded for searching."""
    return fold_greek_series(load_professors(registry_path_str)[SUBJECT_COL])


@st.cache_data
def build_check(external_path_str: str, registry_path_str: str) -> pd.DataFrame:
    """Cross-check every elector of the year tables against a registry export.

    Returns one row per (γνωστικό αντικείμενο, elector) with a status, the
    κωλύματα that are active and the fields that changed.
    """
    workbook = load_external_workbook(external_path_str)
    registry = load_professors(registry_path_str)
    blocking_cols, info_cols = split_flag_columns(registry)

    frames = []
    for entry in workbook.values():
        df = entry["df"].copy()
        df["Κωδικός"] = entry["code"]
        df["Γνωστικό αντικείμενο"] = entry["field"]
        frames.append(df)
    ext = pd.concat(frames, ignore_index=True)
    ext[ID_COL] = pd.to_numeric(ext[ID_COL], errors="coerce")
    ext = ext[ext[ID_COL].notna()]
    ext[ID_COL] = ext[ID_COL].astype(int)

    reg = registry.drop_duplicates(subset=ID_COL).set_index(ID_COL)
    merged = ext.join(reg, on=ID_COL, rsuffix="_reg")
    present = merged[ID_COL].isin(reg.index)

    def reg_col(name: str) -> str:
        """Name the joined registry column took after the rsuffix collision."""
        return f"{name}_reg" if f"{name}_reg" in merged.columns else name

    def collect_flags(columns: list[str]) -> pd.Series:
        """Join the names of the flags that are ΝΑΙ for each row."""
        out = pd.Series([""] * len(merged), index=merged.index)
        for col in columns:
            hit = merged[reg_col(col)].astype(str).str.strip().eq("ΝΑΙ") & present
            labels = pd.Series([col] * len(merged), index=merged.index)
            out = out.where(~hit, out.str.cat(labels, sep=" | ").str.strip(" |"))
        return out

    kolymata = collect_flags(blocking_cols)      # disqualifying
    simanseis = collect_flags(info_cols)         # reported only

    # Changed fields, with the old/new pair kept side by side
    changes = pd.Series([""] * len(merged), index=merged.index)
    out_cols = {}
    year_old = Path(external_path_str).stem.replace("external_", "")
    year_new = re.search(r"(\d{4})", Path(registry_path_str).stem).group(1)
    for ext_col, reg_name, label in COMPARED_FIELDS:
        if ext_col not in merged.columns:
            continue
        old, new = merged[ext_col], merged[reg_col(reg_name)]
        differs = (normalize(old) != normalize(new)) & present
        changes = changes.where(~differs, changes.str.cat(pd.Series([label] * len(merged), index=merged.index), sep=" | ").str.strip(" |"))
        # Prefixed rather than bare years: the two sides can be the same year
        out_cols[f"{label} (πίνακες {year_old})"] = old
        out_cols[f"{label} (μητρώο {year_new})"] = new

    status = pd.Series(STATUS_OK, index=merged.index)
    status = status.mask(changes.ne(""), STATUS_CHANGED)
    status = status.mask(kolymata.ne(""), STATUS_KOLYMA)
    status = status.mask(~present, STATUS_MISSING)

    result = pd.DataFrame(
        {
            "Κωδικός": merged["Κωδικός"],
            "Γνωστικό αντικείμενο": merged["Γνωστικό αντικείμενο"],
            "Κατάσταση": status,
            CHARAKTIRISMOS_COL: merged.get(CHARAKTIRISMOS_COL, ""),
            ID_COL: merged[ID_COL],
            "Επώνυμο": merged["Επώνυμο"],
            "Όνομα": merged["Όνομα"],
            "Κωλύματα": kolymata,
            "Λοιπές σημάνσεις": simanseis,
            "Μεταβολές": changes,
            **out_cols,
        }
    )
    result["Κατάσταση"] = pd.Categorical(
        result["Κατάσταση"], categories=STATUS_ORDER, ordered=True
    )
    return result.sort_values(["Κατάσταση", "Κωδικός", "Επώνυμο"]).reset_index(drop=True)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    return buffer.getvalue()


def apply_text_search(df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Keep rows matching every whitespace-separated term.

    Text terms match as substrings of any searchable column. A purely numeric
    term matches the Κωδικός Χρήστη exactly, so searching "555" returns that
    elector instead of every code containing those digits.
    """
    cols = [c for c in SEARCH_COLS if c in df.columns]
    if not query.strip() or not (cols or ID_COL in df.columns):
        return df
    haystack = (
        df[cols].fillna("").astype(str).agg(" ".join, axis=1).str.casefold()
        if cols
        else pd.Series("", index=df.index)
    )
    codes = df[ID_COL].astype(str) if ID_COL in df.columns else None
    mask = pd.Series(True, index=df.index)
    for term in query.split():
        hit = haystack.str.contains(re.escape(term.casefold()), regex=True)
        if codes is not None and term.isdigit():
            hit |= codes == term
        mask &= hit
    return df[mask]


st.markdown("## Μητρώα γνωστικών αντικειμένων (v2)")

tab_eklektores, tab_antikeimena, tab_external, tab_check, tab_keywords = st.tabs(
    [
        "Σύνολο εκλεκτόρων",
        "Γνωστικά αντικείμενα",
        "Εξωτερικοί εκλέκτορες ανά αντικείμενο",
        "Έλεγχος εγκυρότητας",
        "Αναζήτηση με λέξεις-κλειδιά",
    ]
)


with tab_eklektores:
    files = list_professor_files()

    if not files:
        st.error(f"Δεν βρέθηκαν αρχεία parquet στον φάκελο {PROFESSORS_DIR}")
        st.stop()

    year_label = st.selectbox("Έτος μητρώου", list(files), index=0)
    selected_path = files[year_label]
    df = load_professors(str(selected_path))

    st.caption(f"Αρχείο: `{selected_path.name}` — {len(df):,} εγγραφές")

    query = st.text_input(
        "Αναζήτηση",
        placeholder="π.χ. επώνυμο, φορέας, γνωστικό αντικείμενο ή κωδικός χρήστη",
        help=(
            "Πολλοί όροι χωρισμένοι με κενό — εμφανίζονται μόνο οι εγγραφές που "
            "τους περιέχουν όλους. Αριθμητικός όρος αναζητείται ως ακριβής "
            "Κωδικός Χρήστη."
        ),
    )

    filtered = apply_text_search(df, query)

    filter_cols = [c for c in FILTER_COLS if c in df.columns]
    if filter_cols:
        for col, container in zip(filter_cols, st.columns(len(filter_cols))):
            options = sorted(filtered[col].dropna().unique())
            selected = container.multiselect(col, options, key=f"filter_{col}")
            if selected:
                filtered = filtered[filtered[col].isin(selected)]

    col_total, col_shown, col_foreis = st.columns(3)
    col_total.metric("Σύνολο εκλεκτόρων", f"{len(df):,}")
    col_shown.metric("Εμφανίζονται", f"{len(filtered):,}")
    if "Φορέας" in filtered.columns:
        col_foreis.metric("Φορείς", f"{filtered['Φορέας'].nunique():,}")

    st.dataframe(filtered, use_container_width=True, hide_index=True)

    col_xlsx, col_csv = st.columns(2)
    stem = f"eklektores_{selected_path.stem.replace('professors_export_', '')}"
    col_xlsx.download_button(
        "Λήψη Excel",
        data=to_excel_bytes(filtered),
        file_name=f"{stem}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    col_csv.download_button(
        "Λήψη CSV",
        data=filtered.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{stem}.csv",
        mime="text/csv",
    )

    with st.expander("Στατιστικά"):
        for col in filter_cols:
            st.markdown(f"**{col}**")
            st.bar_chart(filtered[col].value_counts().head(20))


with tab_antikeimena:
    df_ant = load_antikeimena()

    domains = sorted(df_ant["domain"].dropna().unique())
    selected_domains = st.multiselect("Τομέας", domains)
    query_ant = st.text_input(
        "Αναζήτηση αντικειμένου", placeholder="π.χ. σκυρόδεμα", key="search_antikeimena"
    )

    filtered_ant = df_ant
    if selected_domains:
        filtered_ant = filtered_ant[filtered_ant["domain"].isin(selected_domains)]
    if query_ant.strip():
        pattern = re.escape(query_ant.strip().casefold())
        filtered_ant = filtered_ant[
            filtered_ant["field"].str.casefold().str.contains(pattern, regex=True)
        ]

    col_ant_total, col_ant_shown, col_ant_domains = st.columns(3)
    col_ant_total.metric("Σύνολο αντικειμένων", len(df_ant))
    col_ant_shown.metric("Εμφανίζονται", len(filtered_ant))
    col_ant_domains.metric("Τομείς", filtered_ant["domain"].nunique())

    st.dataframe(
        filtered_ant,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Code": st.column_config.NumberColumn("Κωδικός", format="%d"),
            "field": st.column_config.TextColumn("Γνωστικό αντικείμενο", width="large"),
            "domain": st.column_config.TextColumn("Τομέας", width="medium"),
        },
    )

    st.download_button(
        "Λήψη Excel",
        data=to_excel_bytes(filtered_ant),
        file_name="antikeimena.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Αντικείμενα ανά τομέα"):
        st.bar_chart(filtered_ant["domain"].value_counts())


with tab_external:
    external_files = list_external_files()

    if not external_files:
        st.warning(
            f"Δεν βρέθηκαν αρχεία `external_<έτος>.xlsx` στον φάκελο {BY_YEAR_DIR}"
        )
    else:
        col_year, col_field = st.columns([1, 4])
        year = col_year.selectbox("Έτος", list(external_files), key="external_year")
        external_path = external_files[year]
        workbook = load_external_workbook(str(external_path))

        # Label each sheet with its own code + γνωστικό αντικείμενο, ordered by code
        labels = {
            f"{entry['code']} — {entry['field']}": sheet
            for sheet, entry in sorted(
                workbook.items(), key=lambda kv: str(kv[1]["code"])
            )
        }
        label = col_field.selectbox("Γνωστικό αντικείμενο", list(labels))
        entry = workbook[labels[label]]
        df_ext = entry["df"]

        st.markdown(f"### {entry['field']}")
        st.caption(
            f"Κωδικός {entry['code']} · Επιστημονικό πεδίο: {entry['domain']} · "
            f"αρχείο `{external_path.name}`"
        )

        if CHARAKTIRISMOS_COL in df_ext.columns:
            counts = df_ext[CHARAKTIRISMOS_COL].value_counts()
            col_all, col_idiou, col_synafous = st.columns(3)
            col_all.metric("Σύνολο εκλεκτόρων", len(df_ext))
            col_idiou.metric("Ιδίου", int(counts.get("ΙΔΙΟΥ", 0)))
            col_synafous.metric("Συναφούς", int(counts.get("ΣΥΝΑΦΟΥΣ", 0)))

            chosen = st.multiselect(
                "Χαρακτηρισμός", list(counts.index), key="external_charakt"
            )
            if chosen:
                df_ext = df_ext[df_ext[CHARAKTIRISMOS_COL].isin(chosen)]
        else:
            st.metric("Σύνολο εκλεκτόρων", len(df_ext))

        st.dataframe(df_ext, use_container_width=True, hide_index=True)

        st.download_button(
            "Λήψη Excel",
            data=to_excel_bytes(df_ext),
            file_name=f"external_{year}_{entry['code']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="external_download",
        )


with tab_check:
    external_files = list_external_files()
    registry_files = list_professor_files()

    if not external_files:
        st.warning(f"Δεν βρέθηκαν πίνακες έτους στον φάκελο {BY_YEAR_DIR}")
    else:
        col_a, col_b = st.columns(2)
        check_year = col_a.selectbox(
            "Πίνακες έτους", list(external_files), key="check_external_year"
        )
        registry_label = col_b.selectbox(
            "Έλεγχος έναντι μητρώου", list(registry_files), key="check_registry_year"
        )
        external_path = external_files[check_year]
        registry_path = registry_files[registry_label]

        check = build_check(str(external_path), str(registry_path))
        persons = check.drop_duplicates(ID_COL)
        blocking_cols, info_cols = split_flag_columns(load_professors(str(registry_path)))

        st.caption(
            f"Έλεγχος των εκλεκτόρων του `{external_path.name}` έναντι του "
            f"`{registry_path.name}` — {len(check)} εγγραφές, "
            f"{len(persons)} μοναδικά πρόσωπα."
        )
        st.caption(
            "Ως κώλυμα λογίζεται μόνο: "
            + (", ".join(f"«{c}»" for c in blocking_cols) or "—")
            + (
                ". Καταγράφονται χωρίς αποκλεισμό: "
                + ", ".join(f"«{c}»" for c in info_cols)
                if info_cols
                else ""
            )
        )

        cols = st.columns(4)
        for container, status in zip(cols, STATUS_ORDER):
            container.metric(status, int((persons["Κατάσταση"] == status).sum()))

        only_findings = st.toggle(
            "Εμφάνιση μόνο ευρημάτων", value=True, key="check_only_findings"
        )
        view = check[check["Κατάσταση"] != STATUS_OK] if only_findings else check

        st.markdown("### Σύνοψη ανά γνωστικό αντικείμενο")
        summary = (
            check.pivot_table(
                index=["Κωδικός", "Γνωστικό αντικείμενο"],
                columns="Κατάσταση",
                aggfunc="size",
                fill_value=0,
                observed=False,
            )
            .reset_index()
        )
        summary["Σύνολο"] = summary[
            [c for c in STATUS_ORDER if c in summary.columns]
        ].sum(axis=1)
        summary["Ευρήματα"] = summary["Σύνολο"] - summary.get(STATUS_OK, 0)
        summary = summary.sort_values(
            ["Ευρήματα", "Κωδικός"], ascending=[False, True]
        )
        if only_findings:
            summary = summary[summary["Ευρήματα"] > 0]
        st.dataframe(summary, use_container_width=True, hide_index=True)

        st.markdown("### Αναλυτικά")
        objects = ["(όλα)"] + [
            f"{code} — {field}"
            for code, field in check[["Κωδικός", "Γνωστικό αντικείμενο"]]
            .drop_duplicates()
            .sort_values("Κωδικός")
            .itertuples(index=False)
        ]
        chosen_object = st.selectbox(
            "Γνωστικό αντικείμενο", objects, key="check_object"
        )
        if chosen_object != "(όλα)":
            code = chosen_object.split(" — ")[0]
            view = view[view["Κωδικός"].astype(str) == code]

        if view.empty:
            st.success("Δεν βρέθηκαν ευρήματα για την επιλογή αυτή.")
        else:
            st.dataframe(view, use_container_width=True, hide_index=True)

        st.download_button(
            "Λήψη ευρημάτων (Excel)",
            data=to_excel_bytes(check[check["Κατάσταση"] != STATUS_OK]),
            file_name=f"elegxos_{check_year}_vs_{registry_path.stem.replace('professors_export_', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="check_download",
        )


with tab_keywords:
    registry_files = list_professor_files()
    labels = list(registry_files)

    col_reg, col_prev = st.columns(2)
    kw_registry_label = col_reg.selectbox("Μητρώο", labels, key="kw_registry_year")
    kw_registry_path = registry_files[kw_registry_label]
    kw_df = load_professors(str(kw_registry_path))

    # Default comparison is the next oldest registry, i.e. "the previous year"
    others = [lbl for lbl in labels if lbl != kw_registry_label]
    previous_label = (
        col_prev.selectbox(
            "Σύγκριση με",
            others,
            index=min(labels.index(kw_registry_label), len(others) - 1),
            key="kw_previous_year",
            help="Ως «νέοι» θεωρούνται όσοι δεν υπάρχουν στο μητρώο αυτό.",
        )
        if others
        else None
    )
    previous_ids = (
        registry_ids(str(registry_files[previous_label])) if previous_label else set()
    )

    raw_keywords = st.text_area(
        "Λέξεις-κλειδιά",
        placeholder="σκυροδ, γεωτεχν, οπλισμένο σκυρόδεμα",
        help=(
            "Μία λέξη ανά γραμμή ή χωρισμένες με κόμμα. Η αναζήτηση γίνεται "
            "στο Γνωστικό Αντικείμενο, αγνοεί πεζά/κεφαλαία και τόνους, και "
            "βρίσκει και τμήματα λέξεων."
        ),
        key="kw_input",
    )
    keywords = [k.strip() for k in re.split(r"[,\r\n]+", raw_keywords) if k.strip()]

    match_mode = st.radio(
        "Λογική συνδυασμού",
        [MATCH_ANY, MATCH_ALL],
        horizontal=True,
        key="kw_match_mode",
        help=(
            f"«{MATCH_ANY}»: αρκεί μία λέξη να βρεθεί. "
            f"«{MATCH_ALL}»: πρέπει να υπάρχουν όλες στο ίδιο αντικείμενο."
        ),
    )

    if not keywords:
        st.info("Δώστε τουλάχιστον μία λέξη-κλειδί για να γίνει αναζήτηση.")
    else:
        subjects = folded_subjects(str(kw_registry_path))
        # One column per keyword, so we can report which ones matched
        hits = {kw: subjects.str.contains(re.escape(fold_greek(kw))) for kw in keywords}
        matched = pd.DataFrame(hits, index=kw_df.index)
        selected = (
            matched.all(axis=1) if match_mode == MATCH_ALL else matched.any(axis=1)
        )

        results = kw_df[selected].copy()
        results.insert(
            0,
            "Λέξεις που ταιριάζουν",
            matched[selected].apply(
                lambda row: " | ".join(k for k in keywords if row[k]), axis=1
            ),
        )

        if previous_label:
            is_new = ~results[ID_COL].isin(previous_ids)
            results.insert(1, NEW_COL, is_new.map({True: "ΝΑΙ", False: ""}))
            show_only_new = st.toggle(
                f"Μόνο νέοι σε σχέση με το μητρώο {previous_label}",
                value=False,
                key="kw_only_new",
                help=(
                    "Εκλέκτορες που υπάρχουν στο επιλεγμένο μητρώο αλλά όχι "
                    "στο μητρώο σύγκρισης."
                ),
            )
            if show_only_new:
                results = results[is_new]

        blocking_cols, _ = split_flag_columns(kw_df)
        hide_blocked = st.toggle(
            "Απόκρυψη όσων έχουν κώλυμα αποκλεισμού από μητρώα",
            value=False,
            key="kw_hide_blocked",
        )
        if hide_blocked and blocking_cols:
            blocked = pd.Series(False, index=results.index)
            for col in blocking_cols:
                blocked |= results[col].astype(str).str.strip().eq("ΝΑΙ")
            results = results[~blocked]

        col_found, col_new, col_kw = st.columns(3)
        col_found.metric("Εκλέκτορες που βρέθηκαν", f"{len(results):,}")
        if previous_label:
            col_new.metric(
                f"Νέοι από το {previous_label}",
                f"{int(results[NEW_COL].eq('ΝΑΙ').sum()):,}",
            )
        col_kw.metric("Λέξεις-κλειδιά", len(keywords))

        per_keyword = pd.DataFrame(
            {
                "Λέξη-κλειδί": keywords,
                "Εκλέκτορες": [int(matched[k].sum()) for k in keywords],
            }
        )
        if (per_keyword["Εκλέκτορες"] == 0).any():
            missed = per_keyword[per_keyword["Εκλέκτορες"] == 0]["Λέξη-κλειδί"]
            warning = "Καμία εγγραφή για: " + ", ".join(missed)
            if match_mode == MATCH_ALL:
                # With AND a single unmatched keyword empties the whole result
                warning += " — με λογική AND το αποτέλεσμα είναι αναγκαστικά κενό."
            st.warning(warning)

        if results.empty:
            st.info("Δεν βρέθηκαν εκλέκτορες για τις λέξεις αυτές.")
        else:
            st.dataframe(results, use_container_width=True, hide_index=True)
            st.download_button(
                "Λήψη Excel",
                data=to_excel_bytes(results),
                file_name="anazitisi_lexeon.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="kw_download",
            )

        with st.expander("Πλήθος ανά λέξη-κλειδί"):
            st.dataframe(per_keyword, use_container_width=True, hide_index=True)
