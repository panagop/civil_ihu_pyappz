# Εβδομαδιαίο πρόγραμμα — μαθήματα με βάση μόνο τον κωδικό

Status: **implemented 2026-09-16**. The "Classes" section of CLAUDE.md is the
reference from here on. Differences from the plan below: `candidate_courses`
became `course_catalogue` + the pure `build_catalogue` / `courses_for_semester`
/ `programme_semesters` / `examina_label`; `seed_timetable.read_workbook` lost
its `year` argument (only the rule used it); and page 8's class selectbox label
dict was renamed from `labels`, which the staff tab rebinds (found by the new
AppTest).

## Why

During the transition from the 2018 to the 2025 programme some courses may move
between winter and spring, sooner or later, and the details are not known yet.
The rule in `timetable_db` that decides which programme each εξάμηνο follows in
each year (`NEW_CURRICULUM_EXAMINA`) tries to predict what we cannot predict.

Codes are enough to identify a course:

| comparison of the Greek περιγράμματα | count |
|---|---|
| codes in both programmes | 84 |
| only in 2018 (retired) | 18 |
| only in 2025 (new courses get new codes) | 12 |
| shared codes with a different name | 1 — ΣΥΓ017 |
| shared codes with a different εξάμηνο | 20 |
| … of which moved to the other period | 19 |

## Decisions (2026-09-16)

1. **A timetable row is a course code.** It no longer carries a programme.
2. **Drop `timetable_classes.curriculum`.** `perigrammata_courses.curriculum`
   stays: the περιγράμματα need it.
3. **Name = 2025 περίγραμμα if the code is there, else 2018.** ΣΥΓ017 keeps both
   names in `perigrammata_courses`; the timetable prints «Προγραμματισμός και
   Διαχείριση Τεχνικών Έργων» from now on, including in the locked 2025-26
   terms.
4. **Default course list for an εξάμηνο = every code either programme places in
   that εξάμηνο**, once each. A course that moved (ΔΟΜ007: 3rd in 2018, 4th in
   2025) is offered under both.
5. **Toggle «Δυνατότητα επιλογής από όλα τα εξάμηνα»**, off by default: offers
   every code of both programmes, both periods.
6. **The row's εξάμηνο is the timetable's**, i.e. the εξάμηνο being edited, not
   the programme's. This is what lets a course be taught in another period.

Exam schedules (page 3) read course names from the exams workbook, not from the
περιγράμματα, so no code changes there: the new ΣΥΓ017 name is typed into the
next `files/exams/*.xlsm`.

## Changes

### `streamlit/timetable_db.py`

- Remove `NEW_CURRICULUM_EXAMINA`, `OLD_CURRICULUM`/`NEW_CURRICULUM`,
  `curriculum_for`, `off_programme`, `retaggable`, `_retagged`,
  `retag_curricula`, and the re-stamping loop in `open_term` (the copy becomes
  a plain `INSERT … SELECT`; its log line loses the «στο πρόγραμμα 2025» part).
- `SCHEMA_SQL`: remove the column from `CREATE TABLE`, and append
  `ALTER TABLE timetable_classes DROP COLUMN IF EXISTS curriculum;`. It goes
  **in the timetable SQL after its CREATE**, not in `db.MIGRATIONS_SQL`: that
  runs before the timetable schema, and on a fresh database the table would not
  exist yet, failing the whole transaction.
- `_LOAD_TERM_SQL`: replace the join on `(curriculum, code)` with
  `LEFT JOIN LATERAL (SELECT name FROM perigrammata_courses WHERE locale = 'gr'
  AND code = c.course_code ORDER BY curriculum DESC LIMIT 1)`. Same rule in
  one place, as a named constant or a small SQL helper, so every lookup picks
  the newest programme's name.
- `candidate_courses(year, period, all_semesters=False)` → one row per code:
  `course_code`, `course_name` (newest programme), `examina` (e.g.
  `{2018: 3, 2025: 4}`, for the option hint), `in_term` (by code only).
  Default filter: the code is in the selected εξάμηνο in *either* programme.
  The period filter disappears when `all_semesters` is set.
- `add_class` / `update_class`: drop the `curriculum` parameter.
- Module docstring: rewrite the "curriculum stored on the row" paragraph.

### `streamlit/seed_timetable.py`

Drop `curriculum` from the insert and the 2025 → 2018 fallback. An unknown code
is still logged; `load_term` prints «(άγνωστο μάθημα)» for it.

### `streamlit/app_pages/8_🗓_timetable_v2.py`

- Remove the «Ενημέρωση προγράμματος σπουδών» block and the "not in this year's
  programme" warning.
- «Μαθήματα που λείπουν»: columns Κωδικός · Μάθημα · Εξάμηνο 2018 · Εξάμηνο 2025.
- «Νέα γραμμή»: see "Form layout" below.
- Course option label: `ΔΟΜ007 — Name (εξ. 3 στο 2018 · εξ. 4 στο 2025)`; a
  code already in the term keeps its current behaviour.

### Form layout

The toggle changes the options of the course selectbox, so it must trigger a
rerun and therefore **cannot be inside `st.form`** (a widget in a form sends
its value only on submit). Layout:

```python
@st.fragment
def add_class_section():
    all_semesters = st.toggle("Δυνατότητα επιλογής από όλα τα εξάμηνα", key="add_all_semesters")
    choice = st.selectbox("Μάθημα:", options=..., key="add_course")   # outside: drives the suffix default
    with st.form("add_class"):
        ...  # section, suffix, instructors, rooms, placement, notes
        if st.form_submit_button("Προσθήκη", type="primary"):
            ...
            st.rerun()  # scope="app" by default: the week view and table refresh
```

- The toggle and the course selectbox sit **above** the form; everything that is
  only read on submit stays **inside**.
- `st.fragment` so flipping the toggle reruns only this section, not the whole
  page. Page 8 caches nothing, so a full rerun repeats `load_term`, staff, rooms
  and the conflict check.
- The fragment is fed its inputs (term, εξάμηνο, staff, rooms) as arguments,
  computed once in the page body.

### Tests (`tests/test_timetable.py`)

- Delete `test_curriculum_rule` and the retag / `off_programme` assertions.
- Name resolution: ΣΥΓ017 resolves to the 2025 name; a 2018-only code resolves
  to its 2018 name; a code in neither → «(άγνωστο μάθημα)».
- `candidate_courses`: ΔΟΜ007 appears under both εξάμηνο 3 and 4; with
  `all_semesters=True` codes from the other period appear; no duplicate codes.
- Migration: a database created with the old schema loses the column on
  `get_engine()` and its rows still load.
- AppTest: flipping the toggle widens the course options.

### CLAUDE.md

Rewrite "Εβδομαδιαίο πρόγραμμα (page 8)" → "Classes": names by code (newest
programme wins), no curriculum rule, the toggle, and why it sits outside the
form in a fragment.

## Not reversible

Dropping the column loses which programme each existing row was stamped with.
Nothing reads it after this change, and for the seeded 2025-26 terms it can be
reconstructed from the old rule (εξάμηνα 1–2 → 2025, else 2018) in git history.
