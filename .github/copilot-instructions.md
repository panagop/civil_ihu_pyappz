# Civil IHU Python Apps - AI Agent Instructions

## Project Overview
Multi-page Streamlit application for the Civil Engineering Department at IHU (International Hellenic University). Manages course syllabi (περιγράμματα μαθημάτων), professor registries (μητρώα), and exam schedules. Bilingual support (Greek/English).

## Architecture

### Three Main Streamlit Pages
1. **[Περιγράμματα](streamlit/pages/1_📇_perigrammata.py)**: Course syllabi generation with Word export
2. **[Μητρώα](streamlit/pages/2_📊_mitroa.py)**: Professor registry and subject matter expertise tracking (eklektores)
3. **[Exams Schedule](streamlit/pages/3_⛱_exams-schedule.py)**: Exam calendar with interactive visualization

### Data Sources
- **Google Sheets**: Primary data source via `st.secrets['gsheet_perigrammata_id']` and `st.secrets['gsheet_mitroa_id']`
- **Local Excel**: Exam data in `files/exams/*.xlsm` with multiple semester sheets
- **JSON Archives**: Professor data in `files/mitroa/json2024/` and `json2025/` (structured as `{code}-info-{year}.json`)

### Key Patterns

#### Google Sheets Integration
All pages use `@st.cache_data` decorated functions to load Google Sheets as CSV:
```python
sheet_id = st.secrets['gsheet_id_name']
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
df = pd.read_csv(url, dtype_backend='pyarrow', index_col=0)
```

#### Session State for Language/Year
Pages use `st.session_state` to persist user choices:
- `st.session_state['lang']` for "Ελληνικά"/"Αγγλικά"
- `st.session_state['programma_spoudon']` for "Προγραμμα σπουδών 2025"/"2018"

#### Document Generation
- **Template Source**: Word templates fetched from GitHub raw URLs (e.g., `perigrammata-template-gr.docx`)
- **Library**: `docxtpl` (DocxTemplate) for mail-merge style rendering
- **Pattern**: Load template to BytesIO → render with dict → save to buffer → `st.download_button()`

#### Data Cleaning Helper
Standard pattern across codebase:
```python
def replace_none_with_empty_str(some_dict: dict) -> dict:
    return {k: ('' if v is None else v) for k, v in some_dict.items()}
```

## Development Workflow

### Environment Setup
- **Python**: 3.12 (specified in `runtime.txt`)
- **Package Manager**: Uses `pyproject.toml` with hatchling build system
- **Virtual Environment**: `.venv/` directory (activate via PowerShell: `& .venv\Scripts\Activate.ps1`)

### Running the App
```powershell
streamlit run streamlit/home.py
```

### Secrets Configuration
Required `.streamlit/secrets.toml`:
```toml
gsheet_perigrammata_id = "..."
gsheet_mitroa_id = "..."
```

### Dependencies
Core libraries: `streamlit`, `pandas`, `openpyxl`, `docxtpl`, `streamlit-calendar`, `pydantic`

## Project-Specific Conventions

### Greek Language Support
- UI elements, column names, and comments are in Greek
- Variables use transliterated Greek (e.g., `examino` for εξάμηνο, `eklektores` for εκλεκτόρες)
- Always preserve Greek strings exactly as written

### File Naming
- Streamlit pages use emoji prefixes: `1_📇_`, `2_📊_`, `3_⛱_`
- Old/deprecated files prefixed with `ooo_`
- JSON files follow pattern: `{code}-info-{year}.json` (codes 555-606)

### Data Processing Patterns
1. **Excel Exams Data**: Load with `pd.read_excel()`, convert `exam_date` to datetime, add `day_of_week` in Greek
2. **Mitroa JSON**: Contains `code`, `field_name`, `domain_name`, and lists of professor IDs (`eklektores_idiou`, `eklektores_synafous`)
3. **DataFrame Filtering**: Extensive use of boolean indexing with semester/instructor filters

### Tab-Based UI
All pages use `st.tabs()` for organization:
- Πίνακας (Table view)
- Στατιστικά (Statistics with `st.bar_chart()`)
- Download/Export functionality

## Jupyter Notebooks
Used for data preparation and prototyping:
- `jupyter/mitroa2025/`: Professor data processing for 2025
- `jupyter/programmata/`: Exam scheduling prototypes
- Notebooks are exploratory; Streamlit pages are production code

## Testing
Minimal test infrastructure in `tests/` - primarily manual testing via Streamlit UI.

## Common Tasks

### Adding a New Course Syllabus Sheet
1. Update Google Sheets with new sheet name (e.g., "gr_2026")
2. Modify [perigrammata.py](streamlit/pages/1_📇_perigrammata.py) `load_gsheet()` logic
3. Update `st.session_state['programma_spoudon']` options

### Updating Exam Schedules
1. Place new Excel in `files/exams/exams-YYYY-MM.xlsm`
2. Update `INPUT_EXCEL` and `INPUT_SHEET` in [exams-schedule.py](streamlit/pages/3_⛱_exams-schedule.py)
3. Ensure columns match: `course_id`, `course_name`, `semester`, `instructor`, `exam_date`, `start_time`, `room`, `notes`

### Modifying Word Templates
- Templates hosted on GitHub, not in repo
- Update GitHub URL in page file when template changes
- Test with `make_word_file()` function
