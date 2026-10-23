# Εβδομαδιαίο πρόγραμμα (page 8)

Loaded by Claude Code when files in this folder are read. App-wide rules (deployment, database, settings, auth, navigation) are in the root [CLAUDE.md](../../CLAUDE.md).

## Εβδομαδιαίο πρόγραμμα (page 8, Postgres)

**Page 8 is now the weekly timetable** (renamed 2026-09-18, see "Navigation" in the root CLAUDE.md):
it shows five views from the database and adds the preparation of the next
semester. Page 4, which reads `files/timetables/2025-2026.xlsm`, is legacy —
kept for comparison, out of the navigation unless «Παλαιές σελίδες» is ticked,
and due for deletion. Viewing is public as it was on page 4; editing
is for `coordinator_emails` only (decided 2026-09-15), so the page checks the
role where it matters instead of gating itself with `require_ihu_login`.

Tables in [streamlit/timetable/data.py](../../streamlit/timetable/data.py):
`timetable_staff` · `timetable_staff_terms` · `timetable_rooms` ·
`timetable_terms` · `timetable_classes` · `timetable_class_instructors` ·
`timetable_class_rooms` · `timetable_changes` · `timetable_snapshots` ·
`timetable_snapshot_classes`.

### A term is a copy of last year's same period

A *term* is `(year, period)`: `(2026, 'Χειμερινό')` is the winter of 2026-27.
`open_term` **copies** the same period one year earlier — classes, links, and
who was active — and the coordinator rearranges; `lock_term` freezes it. One
term is open at a time. Same model as Εύδοξος, for the same reason: one owner,
so no proposals machinery.

### Εκδοχές: save points of the open term (added 2026-09-26)

A coordinator can freeze the open term as «Επιλογή Α», «Επιλογή Β»…, compare
any two of those and the working timetable (counts, conflicts, a
change-by-change table and two weeks side by side), export one to Word, and
restore one. They live in the Προετοιμασία tab only, for coordinators.

- **Save points, not parallel drafts.** The term stays the one working
  timetable that every view shows and every form edits. Editable drafts would
  need a "which draft?" on every form and a draft filter on every query in
  `timetable.data`, and one forgotten filter would leak a draft into the public
  timetable. So snapshots are separate tables, not a column on
  `timetable_classes`.
- Instructors and rooms are **arrays** in `timetable_snapshot_classes`: a
  frozen copy is never edited row by row. `load_snapshot` returns exactly
  `load_term`'s columns, so `conflicts`, the calendar and the Word export work
  on it unchanged. A snapshot row stores **every column of a class that
  `SNAPSHOT_FIELDS` names** — including `name_curriculum`, since
  `NEWEST_NAME_SQL` reads it — so a column added to `timetable_classes` must be
  added there and to `timetable_snapshot_classes`, or versions silently lose it.
- Rows carry `source_class_id`, the working row they came from; comparison
  (`compare_frames`, pure) and restore match on it, so a moved class reads as
  «Άλλαξε», not removed + added. `restore_snapshot` updates rows in place,
  deletes the extra ones, and re-inserts deleted ones. A re-inserted row gets a
  new id, and **every version of the term is re-pointed to it**. Without that,
  the version just restored would differ from the term it produced, and each
  restore would churn the row again.
- The page's restore saves the current state as a version first (checkbox, on
  by default), so a restore never loses work.
- `lock_term` deletes the term's snapshots, since they can no longer be
  restored. The `ΚΛΕΙΔΩΜΑ` log line gives the count. Saving, restoring and
  deleting a version are logged as `ΕΚΔΟΧΗ` / `ΕΠΑΝΑΦΟΡΑ`, which needed the
  `timetable_changes.action` CHECK re-created at the end of the timetable
  `SCHEMA_SQL`.

### Staff, not instructor

The site lists ΕΤΕΠ who never teach, and "instructor" is the role a person has
on one class. `timetable_staff` holds everyone (category ΔΕΠ / ΕΔΙΠ / ΕΤΕΠ /
ΕΚΤΑΚΤΟΣ / ΑΛΛΟ, rank, site URL) and prints `short_name` — «Σαφούρη Χρ.» vs
«Σαφούρη Γ.». Λιαλιαμπής and Παπαϊωάννου are former ΔΕΠ who still teach and
are filed as ΔΕΠ. «ΔΕΠ» is a `placeholder` row for the courses all faculty
teach (ΓΕΝ009, ΓΕΝ010); the conflict check ignores placeholders.

**Activity is per term** (`timetable_staff_terms`), not a column on the
person: a column per semester would need a migration every term, and a from/to
range does not fit έκτακτοι who come and go. No row means *unknown*, and the
seed marks as active only whoever taught in that term. The Προσωπικό tab edits
the flag for the selected term with a `data_editor`.

### Classes

One row per meeting, as a workbook row was: εξάμηνο, course, `section`
(Θ · Ε · Ε1…Ε9 · Φ), day 1–5, whole `start_hour` (08–20, **no half hours**),
`duration`. `day IS NULL` is a course listed for the term and not yet placed —
the 16 workbook rows with only a semester survive that way, and the
Προετοιμασία tab lists them. Instructors and rooms are link tables: both are
already many-to-one in the data (ΓΕΝ002 has two instructors, ΔΟΜ011 two rooms).

**A class is a course code, not a (programme, code) pair** (since
2026-09-16). New courses of the 2025 programme got new codes and carried-over
courses kept theirs, so the code identifies the course. Which programme an
εξάμηνο follows in a given year is deliberately **not modelled**: in the
transition courses may move between winter and spring, sooner or later, and
nobody can say in advance. The former rule (`NEW_CURRICULUM_EXAMINA`,
re-stamping on `open_term`, `off_programme`) is gone, and so is
`timetable_classes.curriculum` — dropped by an `ALTER TABLE … DROP COLUMN IF
EXISTS` at the end of the timetable `SCHEMA_SQL`, **not** in
`mdb.MIGRATIONS_SQL`, which runs before the timetable schema and would fail on a
fresh database. `perigrammata_courses.curriculum` stays: the περιγράμματα need it.

- **Names come from the περιγράμματα, by code, the newest programme winning**
  (`NEWEST_NAME_SQL`, a lateral join in `load_term`) — so the seed must run
  after the περιγράμματα one. Of the 84 shared codes only ΣΥΓ017 changed its
  name — and it is really a different course («Οργάνωση Εργοταξίου και Δομικές
  Μηχανές» in 2018, «Προγραμματισμός και Διαχείριση Τεχνικών Έργων» in 2025).
  So a row may carry `name_curriculum` (nullable, added 2026-09-23): the
  programme whose title it prints, NULL meaning the newest. `title_options`
  offers the choice only where the programmes disagree; «Νέα γραμμή» lists such
  a code once per title and the edit form gets a «Τίτλος» selectbox.
  `open_term` copies the column. (Page 3 reads names from the exams workbook,
  so there the name is typed into the file.)
- **The row's εξάμηνο is the timetable's**, not the programme's: a course added
  while editing the 3rd εξάμηνο sits in the 3rd whatever either programme says.
  20 shared codes changed εξάμηνο, 19 of them to the other period.
- **What the Προετοιμασία tab offers** is `course_catalogue` (one row per code:
  newest name, `examina` = `{2018: 3, 2025: 4}`, `in_term`), filtered by the pure
  `courses_for_semester`: by default every code **either** programme places in
  the selected εξάμηνο (ΔΟΜ007 appears under both 3 and 4); the toggle
  «Δυνατότητα επιλογής από όλα τα εξάμηνα» offers every code of both periods.
  Each option carries `examina_label` («εξ. 3 στο 2018 · εξ. 4 στο 2025»).
- **The toggle and the course selectbox sit above the «Νέα γραμμή» form, inside
  an `@st.fragment`.** Above the form because they change what the form offers
  and a widget in a form does not rerun until submit (the page-5 elector trap);
  a fragment because page 8 caches nothing, so a full rerun would repeat
  `load_term`, staff, rooms and the conflict check on every flip. A successful
  add calls `st.rerun()`, whose default scope is the whole app, so the table
  and the week above refresh.
- The seed stores codes only and lists any code no περίγραμμα has in its log
  line; `load_term` prints «(άγνωστο μάθημα)» for it.

`name_suffix` keeps the «ΔΥ, ΥΕ» marker the workbook wrote after elective
titles, and `display_name` prints it after the περίγραμμα name. It also drives
the conflict rule below.

**`shape_term` empties the nullable TEXT columns** (`OPTIONAL_TEXT_COLUMNS`:
`name_suffix`, `notes`). A NULL arrives from `read_sql` as `NaN`, and
`NaN or ""` keeps the `NaN` because **NaN is truthy** — which is how the
literal «nan» appeared inside the «Ένδειξη μετά τον τίτλο» text input on
2026-09-15. Emptying them once, where the frame is built, means no caller has
to remember; `add_class` / `update_class` turn `""` back into NULL.

`load_term` returns the workbook's column names (`course_id`, `class_name`,
`full_class_name`, `semester`, `instructors`, `day`, `start_time`, `duration`,
`room`…) so `utils/timetable_export.py` works unchanged on either source.

### Conflicts are warnings, not constraints

`conflicts(frame)` is pure and runs on what `load_term` returns. Two placed
classes on the same day with overlapping hours conflict when they share a
room, share a non-placeholder instructor, or belong to the same εξάμηνο —
**except** parallel lab groups of one course (Ε1 / Ε2 run together on
purpose) and electives whose stream markers share no letter (ΓΥ beside
«ΥΥ, ΔΕ» is fine; ΓΥ beside ΓΕ is not; a course without a marker is common to
all and conflicts with anything). Θ and Ε of the same course overlapping *is* a
conflict. The tab shows them as a table; saving is never blocked, because a
term being rearranged is allowed to be inconsistent. With the stream rule the
seeded 2025-26 still reports a handful, all genuinely in the workbook (ΔΟΜ001
Θ and ΔΟΜ002 Ε1 both in 301 on Tuesday 10:00; ΓΕΩ009 «ΓΥ» over ΥΔΡ006
«ΥΥ, ΓΕ»).

### Seed data

[streamlit/timetable/seed.py](../../streamlit/timetable/seed.py) loads
`files/timetables/staff.csv` and `rooms.csv` while those tables are empty, and
`2025-2026.xlsm` as two locked terms — once. `WORKBOOKS` is an explicit dict,
not a glob: **`2026-2027.xlsm` is a byte-identical copy of the 2025-26 file**
and must not be seeded as 2026-27 (it is kept until the database approach is
confirmed, then deleted). `ROOM_ALIASES` maps every room spelling the workbook
used to a code («Αίθ. Τεχν. Σχεδίου 1» and «Εργαστήριο Τεχνικού Σχεδίου Ι»
are the same room, ΤΣ1); an unmapped spelling **raises**, as does an unknown
instructor. The bare «Σαφούρη» on ΔΟΜ007 is Γεωργία (the drawing lab).

Rooms and staff are edited in the page from then on; the CSVs are the
historical record, like the μητρώα workbooks.

[tests/test_timetable.py](../../tests/test_timetable.py) covers the parser without
a database, and the seed, copy-open, edits, lock and the page script itself
(via `streamlit.testing.v1.AppTest`) against the embedded Postgres.
