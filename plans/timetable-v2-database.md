# Εβδομαδιαίο πρόγραμμα v2 — από Excel σε Postgres

Status: **implemented 2026-09-15** (answers below). Page 4
(`4_📅_weekly_timetable.py`) stays untouched; page 8 is the database version.
The "Εβδομαδιαίο πρόγραμμα (page 8, Postgres)" section of CLAUDE.md is the
reference from here on.

Decisions taken with the answers: Λιαλιαμπής and Παπαϊωάννου are former ΔΕΠ
who still teach (filed as ΔΕΠ); «Εργαστήριο Τεχνικού Σχεδίου Ι» is the one
drawing lab; in 2026-27 εξάμηνα 1–4 follow the 2025 programme and 5–10 the
2018 one (corrected from 1–2 after the first review); Θ over Ε of the same course is a conflict; only `coordinator_emails`
edit; form-based editing; whole hours only; the 2025-26 workbook is seeded as
two locked terms and `2026-2027.xlsm` stays on disk unseeded until the
approach is confirmed. Added after seeing the data: electives of disjoint
κατευθύνσεις may overlap.

## What the workbook holds (2025-2026.xlsm)

Two sheets. `Instructors` is a plain list of 32 surnames (incl. the
placeholder «ΔΕΠ»). `timetable` has 140 rows over 101 course codes:

| fact | value |
|---|---|
| rows with a slot | 124 (16 courses are listed with an εξάμηνο and nothing else) |
| `class_name` values | Θ · Ε · Ε1 · Ε2 · Ε3 · Φροντιστηριακό |
| `instructors` | comma-separated surnames, 1–2 per row |
| `start_time` | whole hours, 09:00–19:00 |
| `duration` | 1 · 2 · 3 · 4 hours |
| distinct `room` strings | 18, two of them are pairs («207 & Αίθ. Τεχν. Σχεδίου 1») |
| `notes` | only «Μάθημα του 7ου/9ου εξαμήνου» on two rows placed in another εξάμηνο |
| codes in the 2025 περιγράμματα | 83 of 101; the other 18 are 2018-only (same split Εύδοξος found) |

`2026-2027.xlsm` is a byte-identical copy of `2025-2026.xlsm`.

## What the department site lists (fetched 2026-09-15)

| category | count | names |
|---|---|---|
| ΔΕΠ (`faculty_members`) | 12 | Κολιόπουλος, Βοζίκης, Γαλάνης, Κίρτας, Μιχαηλίδης, Δανιήλ, Δροσόπουλος, Καζαντζή, Μίκικη, Φωτοπούλου, Βλαχονάσιου, Παναγόπουλος |
| ΕΔΙΠ (`lab_staff`) | 2 | Ιωάννου, Πανταζής |
| ΕΤΕΠ (`tech_staff`) | 2 | Δημητρακάκης, Σαφούρη Χριστίνα |
| Έκτακτοι (`temp_staff`) | 15 | Αναστασιάδης, Αποστολάκη, Αυγέρης, Καλαμάκης, Καπαγιαννίδης, Κοκκαλά, Κόκκινος, Μαραγκός, Μπακάλης, Σαπίδης, Σαφούρη Γεωργία, Σωτηριάδης, Τσιαράπας, Τσοχατζίδης, Φαναραδέλλη |

Each has a profile URL (`/staff/<slug>/`); ranks are on the page for ΔΕΠ.
Mismatches against the workbook:

- In the workbook but not on the site: **Λιαλιαμπής**, **Παπαϊωάννου**, «ΔΕΠ».
- On the site but not in the workbook: Δημητρακάκης, Πανταζής.
- Rows 37–39 (ΔΟΜ007) say just «Σαφούρη» — the sheet has both «Σαφούρη Γ.» and «Σαφούρη Χρ.».

## Proposed tables (prefix `timetable_`, one group like `mitroa_*` / `perigrammata_*` / `eudoxus_*`)

### `timetable_staff` — the people

I recommend **staff**, not instructor. The site lists ΕΤΕΠ and technicians who
never appear in a timetable, and "instructor" is a *role* a person has on one
class, not what the person is. A single table for all four categories:

```
id            SERIAL PRIMARY KEY
last_name     TEXT NOT NULL
first_name    TEXT
short_name    TEXT NOT NULL UNIQUE   -- the label printed in the timetable: «Σαφούρη Χρ.»
category      TEXT CHECK IN ('ΔΕΠ','ΕΔΙΠ','ΕΤΕΠ','ΕΚΤΑΚΤΟΣ','ΑΛΛΟ')
rank          TEXT                   -- Καθηγητής / Αναπληρωτής / Επίκουρος / Λέκτορας
email         TEXT
website_url   TEXT
notes         TEXT
```

### `timetable_staff_terms` — active per semester

Not boolean columns on the staff row (one per semester would need a migration
every term) and not `active_from`/`active_to` (έκτακτοι are hired per semester
and come and go). One row per person per term:

```
staff_id, year, period ('Χειμερινό'|'Εαρινό'), active BOOLEAN, category TEXT
PRIMARY KEY (staff_id, year, period)
```

`category` repeated here because it moves over time (έκτακτος → ΔΕΠ). Absent
row = unknown, not inactive — same rule as Εύδοξος' unchecked books.

### `timetable_rooms`

```
code     TEXT PRIMARY KEY    -- '202', 'ΗΥ1', 'ΕΡΓ-ΓΕΩΔ'…
name     TEXT NOT NULL       -- «Εργαστήριο Γεωδαισίας»
kind     TEXT CHECK IN ('ΑΙΘΟΥΣΑ','ΕΡΓΑΣΤΗΡΙΟ','ΗΥ')
capacity INTEGER
active   BOOLEAN NOT NULL DEFAULT TRUE
notes    TEXT
```

### `timetable_terms` — a semester being viewed or prepared

Term = (year, period). Copy model, as in Εύδοξος: opening 2026/Χειμερινό
copies 2025/Χειμερινό and people rearrange. `ΑΝΟΙΧΤΟ` → `ΚΛΕΙΔΩΜΕΝΟ`.

### `timetable_classes` — one row per meeting (what a workbook row is)

```
id, year, period, examino, course_code, curriculum (2018|2025),
section ('Θ','Ε','Ε1','Ε2','Ε3','Φ'), day (1–5), start_hour, duration,
notes, updated_by, updated_at
```

`curriculum` is stored, so the course name is resolved from
`perigrammata_courses` deterministically instead of "try 2025, then 2018".
A row with `day IS NULL` is a course listed for the term and not yet placed —
that is how the 16 unscheduled workbook rows survive, and it is what the
arranging tab works through.

Instructors and rooms are **link tables** (`timetable_class_instructors`,
`timetable_class_rooms`) because both are many-to-one in the data already
(two instructors on ΓΕΝ002, two rooms on ΔΟΜ011).

### `timetable_changes` — audit log, as `eudoxus_changes`.

## Conflict rules for the arranging tab

Two meetings overlap when same term, same day, hour ranges intersect. Then:

1. same room → **conflict**
2. same instructor → **conflict**
3. same εξάμηνο → **conflict, except** parallel sections of the same course
   (Ε1 / Ε2 / Ε3 of one code), which are meant to run at the same time
4. Θ and Ε of the same course overlapping → conflict (a student attends both)

Shown as a checklist above the calendar, not enforced on save: a term in
preparation is allowed to be inconsistent while it is being moved around.

## The new page (`8_🗓_timetable_v2.py`)

Same five tabs as page 4, reading the database, plus:

- **Προετοιμασία <term>** — pick a term, list its courses (placed / unplaced),
  add or move a meeting in a form (course → section → instructors → rooms →
  day → hour → duration), calendar preview per εξάμηνο / room / instructor,
  conflict checklist, lock.
- **Προσωπικό** — the staff table with the per-term active toggle.
- **Αίθουσες** — the rooms table.

Seeding: `seed_timetable.py` loads `2025-2026.xlsm` once as two locked terms,
the site list as the initial staff, and the 18 room strings normalised to a
room list. Page 4 keeps reading the workbook.
