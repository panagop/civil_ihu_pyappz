# Περιγράμματα (page 6)

Loaded by Claude Code when files in this folder are read. App-wide rules (deployment, database, settings, auth, navigation) are in the root [CLAUDE.md](../../CLAUDE.md).

## Περιγράμματα (page 6, Postgres)

Page 1 reads the Google Sheet and **keeps working** — leave it alone until
explicitly told to remove it. Page 6 is its replacement and never touches the
sheet.

The sheet was exported once, on 2026-09-09, into `files/perigrammata/` and the
database is the master from that point. That export is the *only* thing that
reads those CSVs — after the first successful seed they are the historical
record of what the sheet held at handover, exactly as `external_<year>.xlsx`
is for the μητρώα. **Page 6 does not fall back to them**: with no
`DATABASE_URL` it stops with a message, because a page quietly serving
pre-handover data is worse than a page that says it cannot run. That does mean
page 6 works **only on Railway** until a local Postgres exists (see the
backlog).

All three worksheets (`gr`, `gr_2025`, `eng`) carry identical column names, so
it is one table with `(curriculum, locale, code)` as the key. `eng` matched the
2018 curriculum exactly (102 of its 103 codes), so English is a *locale*, not a
separate programme.

### `perigrammata_courses`

One row per course. The 39 content columns are generated into the DDL from
`CONTENT_COLUMNS` in [streamlit/perigrammata/data.py](../../streamlit/perigrammata/data.py),
so the table, the edit form, the Word context and the seeder cannot drift apart
— add a column there and everything follows. `FIELD_GROUPS` in the same file
carries the Greek label and the widget kind for each, so a new column cannot be
added without deciding how it is edited.

- **`locale`, not `lang`.** `lang` is already a *content* column: the language
  the course is taught in («Ελληνική»). Two different things, and calling both
  `lang` would have been a bug waiting to happen.
- **Numerics are stored as numbers** (`examino` INTEGER, hours/ects NUMERIC) so
  εξάμηνο can be charted and the workload summed. Every value in both years is
  a whole number, and `perigrammata.report.format_value` prints them as such —
  page 1 rendered `4.0` into Word where the sheet said `4`.
- **`sort_order`** is the sheet's old `id`. It is row order within a worksheet,
  not an identity; it survives only so the full report prints in the order
  people are used to.
- **2018 is read-only** (`EDITABLE_CURRICULUM = 2025`), enforced in
  `save_course`, not just hidden in the UI.
- `refs_more_title` / `refs_more` exist in every worksheet but in **neither Word
  template**. They are stored and editable and simply do not print; the Word tab
  says so when a course has them. Dropping them on import would have lost data.

### `perigrammata_revisions`

Append-only, one row per save: `data` is the full new state, `changes` is only
the fields that moved. Both, deliberately — `data` makes a course recoverable,
`changes` makes the coordinator's report a read rather than a diff of
consecutive snapshots.

**No proposals workflow here, unlike μητρώα.** That machinery exists because 52
subjects are worked on by many members with no owner and Streamlit locks
nothing. A περίγραμμα has one natural owner per course, so anyone who passes
the login gate edits directly and the history carries the accountability. The
one concurrency guard is optimistic: the form remembers the `updated_at` it
rendered from and the UPDATE matches on it, so a second tab that loaded earlier
fails loudly instead of silently overwriting.

The seed writes an origin revision per course with an **empty `changes`
object**, so the initial import is recoverable but never shows up as 198
courses "changed" on the day the database was filled.

### Word output

[streamlit/perigrammata/report.py](../../streamlit/perigrammata/report.py), three
documents:

- `render_course` — one περίγραμμα, from the same `docxtpl` template as page 1
  but read from `files/` instead of fetched from GitHub on every click.
- `build_full_report` — every course, one per page, merged with `docxcompose`.
  ~10 s for the 96 courses of 2025, so it sits behind a button, a spinner, and a
  cache keyed on the newest `updated_at`. Each course is rendered separately and
  the results merged because the template is a whole-page form and `docxtpl`
  cannot repeat one.
- `build_changes_report` — what moved in a period, grouped by course, read
  straight out of `perigrammata_revisions`. Not derivable from the courses
  table, which only ever holds the present. Values are truncated to 400
  characters: `subject1` alone runs to 4.583, and printed in full one edit would
  bury every other change.

Column widths total 27.6 cm against the 27.7 cm usable on landscape A4 — the
same trap as the μητρώα report, where **Word ignores every width if the total
overflows**.

### English, deferred

`SEED_LOCALES = ("gr",)` in
[streamlit/perigrammata/seed.py](../../streamlit/perigrammata/seed.py). The English
sheet is exported and committed but not loaded: it writes εξάμηνο as `1st` /
`2nd`, which needs a mapping before it can enter an INTEGER column. Adding
`"eng"` there is the easy half.
