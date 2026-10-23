# Μητρώα (page 5)

Loaded by Claude Code when files in this folder are read. App-wide rules (deployment, database, settings, auth, navigation) are in the root [CLAUDE.md](../../CLAUDE.md).

## Database tables

### Table names (renamed 2026-09-15)

The three μητρώα tables were created as `external_electors`, `year_status`
and `proposals` and renamed to `mitroa_external_electors`,
`mitroa_year_status` and `mitroa_proposals` once the 2026 year was locked, so
that a future admin session's `\dt` shows the app's tables as three groups
(`mitroa_*`, `perigrammata_*`, `eudoxus_*`). `mdb._rename_legacy_tables` does it
on start, **before** `SCHEMA_SQL` and in the same transaction — run after it,
`CREATE TABLE IF NOT EXISTS` would create three empty tables under the new
names and the app would silently start against them. It also renames the
constraints, indexes and the `proposals_id_seq` sequence, which `ALTER TABLE …
RENAME` leaves alone, and prints `[db.get_engine] Μετονομάστηκαν πίνακες: …`
to the deployment log. Before renaming, each table is copied verbatim to
`mitroa_backup_20260915_<old name>`: nothing outside Railway can take a backup,
so the snapshot lives in the database until it has been downloaded.

**Backups leave the database through the app.** The coordinator section of
the «Προετοιμασία <έτους>» tab has «Αντίγραφο ασφαλείας της βάσης (CSV)»
(`mdb.backup_archive`): every `mitroa_*` table as a CSV in one zip. Commit the
download under `files/mitroa/db_backups/`. **Committing it is what removes
the copies**: on start, `mdb._drop_committed_backup_copies` drops every
`mitroa_backup_<date>_*` table for which a `mitroa_db_<date or later>-*.zip`
exists in that folder, and logs it. The archive includes the copies, so a zip
from that day or later holds them; a copy nobody has downloaded yet stays.
The 2026-09-15 snapshot is committed, so the first start after this code
landed drops its copies.

### `mitroa_external_electors`

One row per (year, γνωστικό αντικείμενο, elector) — see [streamlit/mitroa/data.py](../../streamlit/mitroa/data.py):

| Column | Notes |
| ------ | ----- |
| `year`, `field_code`, `elector_id` | composite primary key |
| `characterization` | `TEXT` + `CHECK IN ('ΙΔΙΟΥ','ΣΥΝΑΦΟΥΣ')` — not a boolean, so a third category costs no migration |
| `reasoning` | «Αιτιολόγηση συνάφειας» |
| `created_at` | |

Only the *decisions* live here. Name, φορέας, βαθμίδα, ΦΕΚ etc. are joined in
from an ΑΠΕΛΛΑ export at display time by
[streamlit/mitroa/table.py](../../streamlit/mitroa/table.py), which owns the join
and the column order so the browse tab, the preview of a year in preparation
and the Word report cannot drift apart. A finalised year joins **its own**
snapshot (the electors as they stood when submitted); a year being prepared
joins the current one. Nothing is duplicated, and a historical table still
renders as it was submitted. The year → export mapping is [files/mitroa/registry_snapshots.csv](../../files/mitroa/registry_snapshots.csv)
(`mdb.registry_file_for_year`) — a file, not a table, because it must also
resolve where there is no database, and because the export filenames are date
stamps rather than years. **Never delete an old parquet export**: without it
that year's table cannot be reconstructed.

The `α/α` column is deliberately not stored — it is derived on render. (Only
9 of the 52 sheets in `external_2025.xlsx` reproduce exactly by sorting ΙΔΙΟΥ
first then alphabetically; the rest carry hand-placed rows. That is noise, not
information.)

Seeding lives in [streamlit/mitroa/seed.py](../../streamlit/mitroa/seed.py),
which parses `external_<year>.xlsx` and skips any year that already has rows.
2025 loads as 1.476 rows / 52 αντικείμενα / 500 distinct electors.

**Direction of travel:** from 2026 on there is to be **no `external_<year>.xlsx`
at all** — the year is built and kept in the database, and the workbooks that
exist stay only as the historical record. Keep the file source working for those
years; do not build new features that require one.

`load_external_from_db` (page 5) rebuilds a stored year in exactly the shape
`load_external_workbook` returns — same keys, same column order — so the tab
renders either source with one code path. Electors missing from that year's
export are kept with blank columns and counted in a warning, never dropped.

### Preparing a new year (proposals workflow)

A year under preparation is **never stored as rows** while it is open. Its table
is computed on read as *baseline year + accepted proposals*
(`mdb.working_electors`); only `mdb.finalize_year` writes it into
`mitroa_external_electors`. So `mitroa_external_electors` always means "officially approved",
and the page-5 tab and the Word report need no notion of drafts.

- `mitroa_year_status(year, status, baseline_year, …)` — `ΑΝΟΙΧΤΟ` → `ΚΛΕΙΔΩΜΕΝΟ`.
  `mdb.open_year(2026, baseline_year=2025, …)` copies nothing; it just records
  the baseline.
- `mitroa_proposals` — one row per proposed change: `ΠΡΟΣΘΗΚΗ` / `ΑΦΑΙΡΕΣΗ` /
  `ΜΕΤΑΒΟΛΗ` (carries the new characterisation *and* the new justification, so
  changing both stays one decision a coordinator cannot half-accept; the older
  split `ΧΑΡΑΚΤΗΡΙΣΜΟΣ` / `ΑΙΤΙΟΛΟΓΗΣΗ` still replay), with a mandatory `note`,
  the `author` (from
  `st.user.email`, so it is asserted by Microsoft rather than typed) and
  `ΕΚΚΡΕΜΕΙ` / `ΕΓΚΡΙΘΗΚΕ` / `ΑΠΟΡΡΙΦΘΗΚΕ` / `ΑΠΟΣΥΡΘΗΚΕ`.

**Why proposals rather than editing the table directly:** several members work
on the same 52 subjects and Streamlit locks nothing. Direct edits would mean
last-write-wins, silently. Two proposals on the same elector simply coexist and
the coordinator resolves them.

`working_electors(year, include_pending=True)` replays the undecided proposals
on top of the accepted ones, projecting what the table would become if
everything proposed were approved. The tab renders it under "Ο πίνακας του
<έτους> μετά τις προτεινόμενες αλλαγές", between the proposal forms and the
coordinator section: removals are gone, additions and changes are listed, and
the person columns come from the current ΑΠΕΛΛΑ export. It is shaped by
`mitroa.table`, so it carries **the submitted layout** — this is the table
that goes to the department, not a review view; a checkbox adds a marker column
and can be cleared for the exact submitted form, and it downloads as Excel.
Accepted proposals are replayed before pending ones — they are already reality,
so a pending change to the same elector should win.

Replay rules (`working_electors`): proposals are applied in decision order, so a
later accepted one wins; `ΧΑΡΑΚΤΗΡΙΣΜΟΣ`/`ΑΙΤΙΟΛΟΓΗΣΗ` aimed at an elector who
has since been removed are **no-ops, not errors**. `finalize_year` refuses while
any proposal is still pending.

The tab lives in [streamlit/mitroa/ui.py](../../streamlit/mitroa/ui.py) — the
only part of the app that writes anything. Per-subject it shows the computed
table (🔴 on anyone with a κώλυμα in the current registry, and ➕➖🔄✏️ for
pending proposals), then sub-tabs for Μεταβολή / Προσθήκη / Οι προτάσεις μου.
Below the per-subject part come the year-wide sections — the projected table,
the new electors across all subjects, the Word report — and a coordinator-only
section to decide proposals and lock the year. The
coordinator can decide the current subject's pending proposals **in bulk**
(`mdb.decide_field_proposals` — one UPDATE, so a subject is decided whole or not
at all; a half-applied batch is hard to reason about when one elector has
several proposals), behind a confirmation checkbox, with
`mdb.pending_by_field` showing where the remaining work is. Actions
sit in a form under the table rather than as buttons on each row: ~26 electors
× 3 buttons would rebuild ~80 widgets per rerun for a worse layout.

🟡 marks anyone whose βαθμίδα, γνωστικό αντικείμενο or φορέας moved since the
baseline year, with the before → after in its own column
(`registry_changes`). It compares through `fold_greek_series`, **not**
`casefold`: the exports re-typed subjects in title case, and `ΔΥΝΑΜΙΚΗ` vs
`Δυναμική` differs by an accent casefold keeps — that alone was two false
findings out of 42.

The mandatory «Αιτιολόγηση της μεταβολής» offers ready-made reasons
(`REMOVAL_SUGGESTIONS`, `ADDITION_SUGGESTIONS`, and one built from the actual
characterisation change) through `_suggestion_pick`, which only *fills* the box
— the text stays editable and "Άλλο" clears it. The picker and the
characterisation radio both sit outside the form for the reason below, and the
note's key includes the chosen suggestion so a new choice replaces the default.

**The elector selectbox is outside `st.form` on purpose.** A widget inside a
form does not rerun until submit, so the fields below kept showing the previous
elector's justification. Widget keys also include the elector id, because
Streamlit keeps the stored value of a key that has not changed and would
override the new defaults.

**Losing eligibility in the ΑΠΕΛΛΑ registry removes an elector automatically.**
It is not a judgement anyone makes, so it is not a proposal: `working_electors`
takes `blocked_ids` and drops them at the source, so no view and no write can
forget. `finalize_year` passes the same set, so a locked year cannot contain
one. The tab keeps **two views** of the year: the filtered one drives the preview,
the report and finalisation, while the overview table renders the *unfiltered*
one so the removed electors stay visible in place with 🔴 (a caption says they
are not in the totals and do not reach the submitted table). Seeing who dropped
out, in position, is what makes the table readable at a glance — hiding them
was a regression on 2026-09-06.
`mdb.auto_removals` reconstructs who left, since there is no proposal row to
look at, and the report prints them with `AUTO_REMOVAL_NOTE` ("Διαγραφή λόγω μη
επιλεξιμότητας στο μητρώο του ΑΠΕΛΛΑ") and status `ΑΥΤΟΜΑΤΗ`. For 2026 that is
58 rows / 34 people across 32 subjects. Members still propose *other* removals
themselves, with their own justification.

The report's change table has **three sources**, because each leaves a different
trace: proposals (a row in `mitroa_proposals`), automatic removals (no row at all), and
**registry updates** — βαθμίδα, φορέας, γνωστικό αντικείμενο moving between the
baseline year's export and the current one (`registry_changes`, action
`ΕΝΗΜΕΡΩΣΗ`). The last are not changes to the list, but they *are* changes to
what the list prints, so a reader must see the values were updated rather than
wonder why they differ from last year. Entries are ordered
ΑΦΑΙΡΕΣΗ → ΠΡΟΣΘΗΚΗ → ΜΕΤΑΒΟΛΗ → ΕΝΗΜΕΡΩΣΗ within a subject.

A `ΜΕΤΑΒΟΛΗ` that leaves the characterisation where it was is **left out of the
report**: rewording a justification changes the text printed next to an elector,
not their standing in the μητρώο, and listing it buries the changes that matter.
It is compared against what held *before* — the baseline year plus any addition
already replayed — never against the proposal itself, which carries the new
value. The legacy `ΑΙΤΙΟΛΟΓΗΣΗ` action is reasoning-only by definition and is
dropped the same way. The preview and the tab still show these; only the
submitted document omits them. For 2026: 58
removals and 132 registry updates, so every one of the 52 subjects has a table.

Roles are two, and there is deliberately **no users table**: a coordinator is an
email listed in the `coordinator_emails` setting (the same list also governs
who edits the timetable on page 8) (same mechanism as
`allowed_emails`), everyone else who passes the login gate is a member. Members
may propose on any subject; only a coordinator decides proposals and locks a
year. A table would need an admin screen to manage and still need some way to
appoint the first admin.

### Νέοι εκλέκτορες, όλα τα αντικείμενα μαζί (added 2026-09-09)

Between the per-subject preview and the Word report, `_additions_block` lists
**every elector new since the baseline year, across all 52 subjects**, with a
filter for the still-pending ones and an Excel download. The per-subject
preview answers "what does this subject look like now"; this answers "who did
we take in this year", which is the question asked when the whole cycle is
reviewed — and 52 subjects is too many to answer by clicking through them.

- **"New" is a set difference, not a read of the `ΠΡΟΣΘΗΚΗ` proposals.** It
  compares `(field_code, elector_id)` between the projected year and the
  baseline table, so it reports what the year *is*: an addition that was later
  withdrawn or removed again does not appear, nor does one for somebody who has
  since lost eligibility (`working_electors` already filtered them). Reading
  the proposals instead would list decisions, several of which no longer hold.
- The key is the **pair**, so an elector moved to a different γνωστικό
  αντικείμενο is correctly new *there* — the same rule the rest of the μητρώα
  code follows.
- The proposal that caused each addition is looked up separately
  (`_addition_meta`) only for the κατάσταση / author / note columns. A row with
  no proposal behind it still renders, with those columns blank, rather than
  being dropped.
- The download follows the pending filter, so the file matches what is on
  screen.
- `render` computes `working_electors(..., include_pending=True)` and the
  baseline **once** and passes them to both this block and the report; they
  were being recomputed.

### Consolidated report

[streamlit/mitroa/report.py](../../streamlit/mitroa/report.py) builds one Word
document covering all 52 subjects — a summary table, then a landscape A4 page
each. It takes the same parsed structure either source produces, so the button
works identically for file and database. ~9 s for 1.476 rows, so it sits behind
a button and a spinner rather than being built on load.

`build_report(..., changes=..., draft=...)` puts, **before** each subject's
table, a short table of the additions, removals and changes with the reason for
each — what moved reads first, the list it produced second — and only
who / what / why, since the electors' full details follow underneath. The
`Κατάσταση` column disappears when nothing is `ΕΚΚΡΕΜΕΙ` (it would repeat the
same value on every row), and its width goes to the reason. Rejected and
withdrawn proposals are left out. `draft=True` stamps
"ΠΡΟΧΕΙΡΟ — περιλαμβάνει προτάσεις που δεν έχουν εγκριθεί ακόμη" under the
title, so a report of a year still in preparation cannot be mistaken for the
final one; the "Προετοιμασία <έτους>" tab generates exactly this from the
projected table, and drops the stamp once the year is locked.

**Word, not PDF, on purpose.** `docx2pdf` (already in `pyproject.toml`) shells
out to a real Microsoft Word install and is Windows-only — it cannot run in the
Railway container at all. PDF there would mean WeasyPrint/wkhtmltopdf and system
packages, for a document that gets edited before submission anyway.

Column widths in `COLUMN_WIDTHS` total 27.6 cm against the 27.7 cm usable on
landscape A4 at 1 cm margins. **Word silently ignores every width if the total
overflows the page**, so adding a column means taking the room from another.

**The database view is not byte-identical to the workbook, by design.** Of
13.284 compared cells for 2025, 475 differ: 40 are only capitalisation (the
workbook is hand-typed in caps), 334 are the known Κατηγορία Χρήστη relabelling
(`Ημεδαπής` → `Καθηγητής Ημεδαπής` — the workbook predates the export it is
joined to), and ~100 are real drift in Βαθμίδα, ΦΕΚ and Σχολή. The database view
shows the **official registry values**; the file view shows what was typed.

## Page 5 — μητρώα v2

Tabs: **Εκλέκτορες ΑΠΕΛΛΑ** (browse an annual export) · **Γνωστικά αντικείμενα
ΔΙΠΑΕ** (the 52 subjects) · **Εξωτερικοί εκλέκτορες ΔΙΠΑΕ** (one subject at a
time, from **either** the database **or** the submitted workbook — a radio picks
the source, defaulting to the database, which is where a year lives from 2026
on; the file option is the historical archive, and the database option appears
only where there are stored years) · **Έλεγχος εγκυρότητας** (cross-check a
submitted year against a registry export) · **Αναζήτηση με λέξεις-κλειδιά**
(find candidates by γνωστικό αντικείμενο, OR/AND, flag those new since a chosen
year) · **Προετοιμασία <έτους>** (propose, decide and finalise the year being
prepared — see below).

`WORKING_YEAR` / `BASELINE_YEAR` at the top of the page name the year being
prepared and the finalised one it starts from. Bump both when the next cycle
begins, and add the new registry export to `registry_snapshots.csv`.

Reads only local files — no secrets, no network.

### CSV per subject (added 2026-10-23)

The Εξωτερικοί εκλέκτορες tab downloads the selected year (either source) as a
zip of one `<subject title>.csv` per γνωστικό αντικείμενο:
`ΚΩΔΙΚΟΣ ΧΡΗΣΤΗ,ΕΠΩΝΥΜΟ,ΣΥΝΑΦΕΙΑ`, 1 = ΙΔΙΟΥ, 0 = ΣΥΝΑΦΟΥΣ, rows in the table's
order. The model is `files/mitroa/csv_test/Θεμελιώσεις - Αντιστηρίξεις.csv`
(the 2025 table of 575) — **UTF-8 without BOM, CRLF**, and the generated file
for 2025/575 is byte-identical to it; `tests/test_mitroa_csv.py` holds that.

- `mitroa.table.csv_archive` builds it; titles lose `\ / : * ? " < > |` and
  trailing dots (570 carries `:` and a final `.`, 606 `Η/Υ`).
- Built on every run, not passed to `download_button` as a callable: a callable
  runs on another thread, so an unknown χαρακτηρισμός would fail unseen there.
  It raises instead of exporting a 0, and the tab shows the error.

### Registry data files

- `files/mitroa/professors_tables/professors_export_<YYYYMMDD>.parquet` — annual
  ΑΠΕΛΛΑ export. **Schemas differ between years**: 2024/2025 have `Σε αναστολή`
  and `Παρ. 1, Άρ. 145, Ν. 4957/2022`; 2026 replaced these with `Ενεργή
  άδεια/κώλυμα αποκλεισμού από μητρώα` and `... από εκλεκτορικά/επιτροπές`. Code
  must tolerate this — never assume a fixed column set.
- `files/mitroa/mitroa_by_year/external_<year>.xlsx` — the submitted external
  electors. One worksheet per γνωστικό αντικείμενο, **named by code** (555–606);
  the readable title sits at row 6 (0-based `META_ROW`), the table header at row
  9 (`HEADER_ROW`). The `α/α` column holds spreadsheet formulas — renumber it.
- `files/mitroa/antikeimena.csv` — the 52 αντικείμενα (`Code, field, domain`).

Domain rules encoded in page 5:

- **Only `κώλυμα αποκλεισμού από μητρώα` disqualifies** (and legacy `Σε
  αναστολή`). Exclusion from εκλεκτορικά/επιτροπές is reported but does not
  block — see `BLOCKING_FLAG_KEYWORDS`. Flag columns are detected by their
  ΝΑΙ/ΟΧΙ *values*, so renamed columns keep working.
- **Κατηγορία Χρήστη is never compared across years** — the exports relabelled it
  (`Ημεδαπής` → `Καθηγητής Ημεδαπής`), which would produce 179 false findings.
- Χαρακτηρισμός is spelled inconsistently in the workbooks, in **both** the
  values (`ΙΔΙΟ`/`ΙΔΙΟΥ`, `ΣΥΝΑΦΕΣ`/`ΣΥΝΑΦΟΥΣ` — normalised via
  `CHARAKTIRISMOS_ALIASES`) and the column header itself
  (`Χαρακτηρισμός` vs the soft-hyphenated `Χαρακτη-ρισμός`). **Never match a
  workbook header literally** — `load_external_workbook` renames it to the
  canonical `CHARAKTIRISMOS_COL` using `fold_header` (fold_greek with
  non-letters dropped). A literal constant silently matched nothing and left
  the ΙΔΙΟΥ/ΣΥΝΑΦΟΥΣ metrics, filter and comparison column dead until
  2026-09-05. `mitroa/seed.py` matches its headers the same way.

### Greek text matching

Use `fold_greek` / `fold_greek_series`, not `casefold()`. Plain casefolding fails
twice over: `σκυροδέμ` does not match `Σκυρόδεμα` (the accent sits on a different
vowel) and `ς` does not fold to `σ`. The helpers strip combining accents, casefold
and unify final sigma.

The parquet columns are Arrow-backed, so pandas runs their regexes through
**RE2, which rejects `\u` escapes** — build character classes from `chr()` (see
`COMBINING_MARKS_RE`) rather than writing `"[̀-ͯ]"`.
