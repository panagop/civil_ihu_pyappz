# civil-ihu-pyappz

Streamlit multi-page application for the Civil Engineering department at the International Hellenic University (IHU). Manages course syllabi, course registries (μητρώα), exam scheduling, and weekly timetables.

## Running the app

```bash
uv run streamlit run streamlit/home.py
```

## Project structure

```
civil_ihu_pyappz/
├── streamlit/                        # Streamlit app (entry point + pages)
│   ├── home.py                       # Landing page / entry point
│   ├── pages/
│   │   ├── 1_📇_perigrammata.py      # Course syllabi — reads from Google Sheets, exports Word docs
│   │   ├── 2_📊_mitroa.py            # Course registries (eklektores/antikeimena) — Google Sheets
│   │   ├── 3_⛱_exams-schedule.py    # Exam schedule — reads files/exams/*.xlsm, exports Word/calendar
│   │   └── 4_📅_weekly_timetable.py  # Weekly timetable — reads files/timetables/*.xlsm, exports Word
│   └── .streamlit/
│       └── secrets.toml              # Google Sheets IDs + API credentials (NOT in git — create locally)
├── civil_ihu_pyappz/                 # Python package (legacy; perigrammata.py not used by the app)
├── files/
│   ├── exams/                        # Exam Excel files (.xlsm); active: exams-2026-06.xlsm
│   └── timetables/                   # Timetable Excel files (.xlsm); active: 2025-2026.xlsm
│   └── mitroa/                       # Registry JSON exports (json2024/, json2025/)
├── jupyter/                          # Exploration notebooks (not part of app)
├── tests/                            # Minimal tests (pytest)
└── pyproject.toml                    # Dependencies — managed with uv
```

## Dependencies

Uses `uv` as the package manager.

```bash
uv sync           # install all dependencies
uv sync --extra dev   # include dev tools (pytest, ruff, black)
```

Key libraries: `streamlit`, `pandas`, `openpyxl`, `python-docx`, `docxtpl`, `streamlit-calendar`, `pydantic`.

## Secrets / credentials

`streamlit/.streamlit/secrets.toml` is gitignored. On a new machine, create it manually with:

```toml
gsheets_id_perigrammata = "..."
gsheets_id_mitroa_eklektores = "..."
gsheets_id_mitroa_antikeimena = "..."
# add any other Sheet IDs used by the pages
```

The Google Sheets are accessed as public CSV exports (no OAuth needed, just the sheet IDs).

## Active data files

Update these paths inside the page files when switching academic year:

| Page | Active file |
|------|-------------|
| Exam schedule | `files/exams/exams-2026-06.xlsm` |
| Timetable | `files/timetables/2025-2026.xlsm` |

## Known improvement backlog

These are planned refactors (no functionality changes):

1. **Split large page files** — `3_⛱_exams-schedule.py` and `4_📅_weekly_timetable.py` are 750+ lines; extract data loading, Word export, and calendar logic into sibling modules.
2. **Remove legacy `civil_ihu_pyappz/perigrammata.py`** — dead code, duplicates page 1.
3. **Delete commented-out dead code** in exams-schedule.py.
4. **Guard `.iloc[0]` calls** — add `.empty` checks before row access to prevent crashes.
5. **Wrap Excel loading in try/except** — show `st.error()` if file is missing or locked.
6. **Guard `st.secrets` access** — use `.get()` with a friendly error if secrets are missing.
7. **Centralise semester color map** — currently duplicated in pages 3 and 4.
8. **Replace magic strings with constants** — column names, time slots, file paths.
9. **Add type hints** to key functions (especially Word export helpers).
10. **Extract Word export logic** into standalone functions (currently embedded in button handlers) so they can be unit-tested.
11. **Add smoke tests** for document generation.
12. **Move active file paths to a config section** so year updates are a single-line change.

## Notes

- `streamlit/_ooo_exams-schedule_old.py` is an archived previous version of page 3 — kept for reference, not loaded by Streamlit.
- Python 3.12 required (pinned in pyproject.toml and runtime.txt).
