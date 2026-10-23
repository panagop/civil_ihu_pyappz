# Εύδοξος (page 7)

Loaded by Claude Code when files in this folder are read. App-wide rules (deployment, database, settings, auth, navigation) are in the root [CLAUDE.md](../../CLAUDE.md).

## Εύδοξος (page 7, Postgres)

The books offered per academic year, and the list for the year being prepared.
A year is named by the one it starts in: `2025` is «2025-26» (`year_label`).
Database-only, for the same reason as page 6.

Tables: `eudoxus_years` · `eudoxus_courses` · `eudoxus_selections` ·
`eudoxus_books` · `eudoxus_changes`, all in
[streamlit/eudoxus/data.py](../../streamlit/eudoxus/data.py).

### A new year is a copy, not a replay

The μητρώα tables compute an open year as *baseline + accepted proposals* and
store nothing until it is finalised. That machinery exists because 52 subjects
are worked on by many members with no owner. A book list has one teacher per
course, so `open_year` **copies** the baseline outright and people edit their
own courses directly, exactly as they edit περιγράμματα. What the coordinator
reviews is `changes_vs_baseline` — a FULL OUTER JOIN of the two years, so it
cannot go stale and shows the *net* change; `eudoxus_changes` is the separate
audit log of who did each edit, including the ones that cancelled out. One year
is open at a time, and `open_year` refuses a second.

### The εξάμηνο is part of the course key

**ΔΟΜ022 «Οικοδομική ΙΙ» is listed twice** in the 2025-26 export, once as 7th
εξάμηνο and once as 9th — the old programme running beside the new one — with
the same three books in a *different* priority order (0/1/2 vs 0/2/1). So
`eudoxus_courses` is keyed on `(year, course_code, examino)` and
`eudoxus_selections` on `(year, course_code, examino, book_id)`. Keying on the
course code alone silently loses one of the two orderings.

The list also spans **both** curricula: of its 101 courses, 83 are in the 2025
περιγράμματα and the other 18 only in the 2018 ones. Do not assume a book list
maps onto one πρόγραμμα σπουδών.

### `eudoxus_books` — the catalogue cache

Book metadata is **not** in `eudoxus_selections`: a book appears in several
courses (222 distinct books over 291 rows) and its availability changes on
Eudoxus' side, not ours. `found` distinguishes "withdrawn from the registry"
from "this code is not in the registry at all" — different problems for whoever
has to fix the list.

- **A book that has never been checked is not reported as a problem.** Absence
  of an answer is not a negative one; the page counts those separately.
- **`active` and `selectable` are different flags and both matter.** Of the 9
  problem books found on 2026-09-09, 4 are `active=True, selectable=False` —
  checking only `active` would have missed them.
- Type coercion in `_clean_book` is not decoration: ISBNs arrive as **integers**
  from both the JSON and pandas, and Postgres has no assignment cast from
  bigint to text, so an uncoerced one fails the INSERT. numpy scalars out of a
  DataFrame are the same class of problem — psycopg cannot adapt them at all.

### The Eudoxus client

[streamlit/eudoxus/client.py](../../streamlit/eudoxus/client.py) is an **unofficial,
reverse-engineered** client for the JSON endpoint behind the public
[Σύνθετη Αναζήτηση](https://service.eudoxus.gr/search/#/advanced) page. There is
no published contract:

- `PUT /search/rest/app/advanced-search`, and **unknown keys are rejected with
  HTTP 500** — send only fields in `_TEMPLATE`.
- `id` is an exact match on a single code and **there is no bulk endpoint**, so
  checking a year is inherently one request per book: ~1 s each, ~3.5 minutes
  for 222. That is why results are stored rather than fetched per page view,
  and why the check runs behind a button with a progress bar.
- `fetch_books` never raises — a dead code comes back `found=False` with the
  reason, so one bad book cannot abort a check of two hundred.

It grew out of [files/eudoxus/eudoxus.py](../../files/eudoxus/eudoxus.py), which stays
where it was written as a standalone script.

### Seed data

Two department exports, one per year, plus a catalogue dump:

| File | Year | Rows |
|------|------|------|
| `eudoxus_books_2025-26.xlsx` | 2025 | 291 → 102 offerings, 291 selections |
| `Συγγράμματα ΕΥΔΟΞΟΣ 2026-2027.csv` | 2026 | 305 → 105 offerings, 305 selections |
| `eudoxus_catalogue_<YYYYMMDD>.csv` | — | what Eudoxus said about 222 books |

The catalogue was fetched once so the browse tab shows titles rather than bare
numeric codes from the first run. It is loaded **only while the table is
empty** — after that the availability check owns it, and re-applying the dump
on every restart would replace a fresh answer with a stale one.

**The exports are named in `EXPORTS`, one file per year — not globbed**
(since 2026-09-18). Working copies get dropped beside the real file: that day
three sat next to the official CSV, one of them a different export with 18
rows that have no εξάμηνο, and a glob seeded whichever sorted first (Latin
before Greek), which failed every fresh seed on `int(NaN)`. Same reason
`timetable.seed.WORKBOOKS` is a dict. `export_year` still parses the year out
of a filename and the tests use it to check each listed file is under the
right key. `.csv` and `.xlsx` carry identical columns, so only the reader
differs (`read_export`). Exports are seeded in **year order** — a year needs
its baseline in the table before it arrives.

**The production 2026-27 was not seeded from the CSV.** The Railway log shows
2026 already existing when the CSV-capable seeder first ran (2026-09-17
21:12 UTC) and absent fifteen minutes earlier; nothing but the admin tab's
«Άνοιγμα έτους» can create a year without printing, so it was opened as a
copy of 2025-26. `eudoxus.data.delete_year` exists for exactly that: the admin
tab's coordinator-only «Διαγραφή του ανοιχτού έτους» removes the open year's
list, audit log and year row (a locked year is refused), then runs
`eudoxus.seed.seed_eudoxus` at once so the year comes back from its file without a
restart. The rows are gone for good; the deployment log carries the counts.

Seeded years are inserted as `ΚΛΕΙΔΩΜΕΝΟ` unless they are listed in
`OPEN_SEED_YEARS`, which is `{2026: 2025}`: 2026-27 was declared in Εύδοξος
before this app existed, so its list arrives as a file like 2025-26 did, but it
is the year people are still working on. It is seeded `ΑΝΟΙΧΤΟ` with
`baseline_year = 2025`, so page 7 edits it and the admin tab's «Μεταβολές»
compares it against 2025-26 exactly as if `open_year` had produced it (35
additions, 21 removals, 3 new course offerings). Two conditions withdraw that
and leave the year as ordinary history rather than failing the bootstrap: the
baseline must exist, and **only one year may be open** — a coordinator who has
already opened one in the app must not find a second appearing underneath them.

From 2027 on a list is opened as a copy and edited in the app; a new file is
needed only if the department hands over another export.

The 18 book codes new in 2026-27 are **not** in the catalogue dump, so they read
as «χωρίς έλεγχο» until the availability check is run — absence of an answer is
not a negative one. All 9 books the dump found unusable were dropped from the
2026-27 list already.

**`~` on a pandas column of booleans is only safe on `bool` dtype.** A year
holding a book the catalogue has never seen gives the LEFT JOIN a NULL, pandas
reads the column back as **object**, and `~` is then the bitwise inversion:
`~True` is `-2`, `~False` is `-1`, both truthy. With one clean bool column this
is invisible; the 18 unknown codes of 2026-27 turned 9 real problems into 210.
`eudoxus.data.unusable_mask` casts first and is the one place the three flags are
combined — page 7 calls it too.
