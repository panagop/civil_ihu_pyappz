# Εκδοχές (saved versions) of the timetable in preparation — page 8

## Context

While preparing the next term on page 8 («Προετοιμασία» tab), the coordinator
rearranges classes directly in the one open term. There is no way to try an
arrangement, keep it, try another, and compare them, so a promising layout is
lost as soon as it is edited further. The goal is «Επιλογή Α», «Επιλογή Β»…
that can be saved, compared and restored before the term is locked.

**Is it easy, or too confusing?** It is moderate work and stays easy to follow
if the versions are **save points**, not parallel editable timetables (decided
with the user):

- There is still **one working timetable**, the one every public view shows
  and every edit form changes. Nothing that exists today changes meaning.
- «Αποθήκευση ως εκδοχή» freezes a copy under a name. A copy can be viewed,
  compared with the working timetable or with another copy, exported, restored
  into the working timetable, or deleted.
- **Coordinators only**, inside the Προετοιμασία tab's editable branch.

Parallel editable drafts were rejected: every form would need a "which draft?"
selector, it is easy to edit the wrong one, and every query in `timetable_db`
filters on `(year, period)` and would have to learn about drafts.

## Storage — two new tables, in `timetable_db.SCHEMA_SQL`

Separate tables rather than a `variant` column on `timetable_classes`: that
column would leak into `load_term`, `conflicts`, `open_term`'s copy,
`course_catalogue.in_term` and the seed, and a forgotten filter would mix a
saved version into the public timetable.

```sql
timetable_snapshots (
  id BIGSERIAL PK, year, period, name TEXT NOT NULL, note TEXT,
  created_by TEXT, created_at TIMESTAMPTZ DEFAULT now(),
  FOREIGN KEY (year, period) REFERENCES timetable_terms,
  UNIQUE (year, period, name))

timetable_snapshot_classes (
  snapshot_id BIGINT REFERENCES timetable_snapshots ON DELETE CASCADE,
  source_class_id BIGINT NOT NULL,        -- the working row it was copied from
  examino, course_code, section, name_suffix, day, start_hour, duration, notes,
  instructor_ids INTEGER[] NOT NULL,      -- in position order
  room_codes TEXT[] NOT NULL)
```

- Instructors/rooms as **arrays**, not link tables: a frozen copy is never
  edited row by row, and arrays keep save and restore one statement each.
- `source_class_id` is what lets a comparison say "moved" rather than
  "removed + added".
- New change-log actions `ΕΚΔΟΧΗ` (saved) and `ΕΠΑΝΑΦΟΡΑ` (restored) go into
  `ACTIONS`. The inline CHECK on `timetable_changes.action` is not altered by
  `CREATE TABLE IF NOT EXISTS`, so append to the migrations at the end of the
  timetable `SCHEMA_SQL` (next to the `DROP COLUMN curriculum` one):
  `DROP CONSTRAINT IF EXISTS timetable_changes_action_check` + `ADD CONSTRAINT`.
  Confirm the generated constraint name in the embedded-Postgres test.

## Functions — `streamlit/timetable_db.py`

- `save_snapshot(year, period, name, note, author) -> str` — refuses unless
  the term is open (`is_editable`), the name is blank, or already used. One
  `INSERT … SELECT` from `timetable_classes` with `array_agg` over the link
  tables (ordered by `position`). Logs `ΕΚΔΟΧΗ`.
- `list_snapshots(year, period) -> DataFrame` — name, note, who, when, row count.
- `load_snapshot(snapshot_id) -> DataFrame` — a `_LOAD_SNAPSHOT_SQL` returning
  **the same columns as `_LOAD_TERM_SQL`** (`id` = `source_class_id`,
  names via `unnest … WITH ORDINALITY` joined to staff, the same
  `NEWEST_NAME_SQL` lateral join, `conflict_ids` excluding placeholders), then
  `shape_term`. So `conflicts`, `render_week`, `to_display` and
  `create_weekly_timetable_document` work on a snapshot unchanged.
- `restore_snapshot(snapshot_id, author) -> str` — one transaction, refuses on
  a locked term. Matches rows by `source_class_id`: rows still in the working
  term are UPDATEd in place (and their instructor/room links reset); rows since
  deleted are re-INSERTed; working rows not in the snapshot are DELETEd. Keeping
  ids where possible keeps other snapshots comparable. One `ΕΠΑΝΑΦΟΡΑ` log line
  with the counts.
- `delete_snapshot(snapshot_id) -> str`.
- `compare_frames(a, b) -> DataFrame` — **pure**, on two `load_term`-shaped
  frames, keyed on `id`: `Προστέθηκε` / `Αφαιρέθηκε` / `Άλλαξε` with a
  «πριν → μετά» description of what moved (ημέρα/ώρα/διάρκεια, αίθουσα,
  διδάσκοντες, εξάμηνο, τμήμα). Unchanged rows omitted.
- `lock_term` deletes the term's snapshots in the same transaction (logged in
  the `ΚΛΕΙΔΩΜΑ` detail): once locked they can neither be restored nor matter.

## UI — `streamlit/app_pages/8_📅_weekly_timetable.py`

A new «#### Εκδοχές» section in the editable branch of `tab_prepare`, just
above «Κλείδωμα», as an `@st.fragment` (same reason as `add_class_section`:
switching the comparison selectors should not reload the whole page):

1. **Save**: small form: name (placeholder «Επιλογή Α»), optional note, button.
2. **List**: `list_snapshots` as a table, plus conflicts / χωρίς ώρα counts per
   version so the options can be ranked at a glance.
3. **Compare**: two selectboxes, each offering «Τρέχον πρόγραμμα» + every
   snapshot. Shows side-by-side metrics (γραμμές, χωρίς ώρα, συγκρούσεις), the
   `compare_frames` table, and for a chosen εξάμηνο σπουδών two `render_week`
   calendars in `st.columns(2)` (keys include both version ids).
4. **Restore / delete / export**: pick a snapshot; «Επαναφορά» behind a
   confirmation checkbox, with a checkbox (on by default) «Αποθήκευση του
   τρέχοντος ως εκδοχή πριν την επαναφορά», so a restore never loses work;
   «Διαγραφή»; and a Word download of that snapshot via
   `create_weekly_timetable_document`. Restore/delete call `st.rerun()`
   (whole app, so the table and calendar above refresh).

Public views, the term selector, and the per-row edit forms are untouched.

## Tests — `tests/test_timetable.py`

- Pure: `compare_frames` on `_frame(...)` specs (added, removed, moved, room
  change, identical → empty).
- Embedded Postgres (`database` fixture), extending the open-term flow:
  save «Α» → `load_snapshot` equals `load_term` on the compared columns;
  edit/add/delete in the working term → save «Β» → compare shows those;
  restore «Α» → `load_term` equals «Α» again; duplicate name refused; save on a
  locked term refused; `lock_term` removes the snapshots; the action CHECK
  accepts `ΕΚΔΟΧΗ` / `ΕΠΑΝΑΦΟΡΑ` after the migration runs on an existing table.
- `test_page_renders` still passes (AppTest).

## Docs

Short subsection in `CLAUDE.md` under «Εβδομαδιαίο πρόγραμμα»: save points,
not drafts (and why); arrays in the snapshot table; matching by
`source_class_id` (a row re-inserted by a restore gets a new id, so older
snapshots see it as removed + added); snapshots deleted on lock; add
`timetable_snapshots*` to the table list. Also copy this plan into
`plans/` (project convention).

## Verification

1. `uv run --extra dev pytest tests/test_timetable.py`
2. `uv run ruff check streamlit tests`
3. Locally, run the app against a local Postgres (`DATABASE_URL`, per backlog
   item 6) as a coordinator: open a term, save «Επιλογή Α», move two classes,
   save «Επιλογή Β», compare Α↔Β and Α↔current, restore Α, lock and check the
   snapshots are gone. After pushing, check the Railway deployment log for the
   bootstrap line listing the two new tables.
