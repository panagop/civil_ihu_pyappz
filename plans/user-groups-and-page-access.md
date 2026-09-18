# Plan: user groups, and access per page and per tab

Status: **proposal, not built**. Written 2026-09-18 for review — the audience
tables in §4 and §5 are guesses about department practice and are the part to
edit first.

## 1. The blocker comes before the groups

> *"Not all έκτακτοι have an IHU Microsoft account."*

That is not an authorization problem, and no group list solves it.

`st.login()` is pointed at `login.microsoftonline.com/<TENANT_ID>/v2.0/…` — a
**tenant-specific** endpoint. Only accounts inside the IHU tenant can
authenticate at all. A colleague without one never reaches the point where
`allowed_emails`, `coordinator_emails` or any future `ektaktoi_emails` is
consulted: they are turned away by Microsoft, not by this app.

So the first decision is not where groups live. It is **how (or whether) those
colleagues sign in.** Four routes:

| Route | What it costs | Consequence |
|---|---|---|
| **(a) IHU accounts for them** | Nothing in code; an IT request per person | One identity provider, one domain rule, everything below stays simple |
| **(b) Entra B2B guest invitations** | Tenant admin must allow guests; one invitation per person | They keep their own address. **Code impact:** a guest's `email` claim is often the mangled UPN `someone_gmail.com#EXT#@ihu.gr`, so the allowlist must match the verified claim and may need to accept both forms |
| **(c) Widen to `organizations`/`common`** | One line of config | **Not recommended.** Any Microsoft account on earth reaches the allowlist, which today defaults to *"empty means allow every `@ihu.gr`"* — a misconfiguration stops being a lockout and becomes an open door |
| **(d) Add a second provider (Google)** | Streamlit 1.63 supports `[auth.<name>]` sections and `st.login("google")` (verified in `auth_util.validate_auth_credentials`); needs a Google OAuth client, a second redirect URI, and the flat `[auth]` block converted to named sections — which touches `scripts/write_secrets_toml.py` and the existing Microsoft call site | Works for anyone with a Gmail address, at the cost of the "one tenant, one rule" simplicity the current setup is built on |

**Recommendation: (a) where IT will do it, (b) for the rest.** Avoid (c).
Treat (d) as a fallback if the tenant forbids guests.

Everything below works unchanged whichever route is chosen — but until one is,
`ektaktoi_emails` would be a list of people who cannot log in.

## 2. The mechanism

One lookup, in `auth.py`, and pages never read a settings list directly:

```python
PUBLIC = "public"            # explicit, not "forgot to gate this page"
IHU    = "ihu"               # implicit for every @ihu.gr login

def groups(email: str | None = None) -> frozenset[str]:
    """Every group the user belongs to. The one place membership is resolved."""

def has_any(*wanted: str) -> bool:      # for a tab, a section, a button
def require(*wanted: str) -> None:      # for a page; stops with a message
```

Membership comes from `<name>_emails` settings, exactly like today's
`coordinator_emails`: `dep_emails`, `edip_emails`, `etep_emails`,
`ektaktoi_emails`, `secretariat_emails`, `guest_emails`.

Three properties worth keeping:

- **Any `@ihu.gr` login is implicitly in `ihu`.** So today's
  `require_ihu_login()` becomes `require(IHU)` with no behavioural change, and
  the migration in §6 can be a no-op commit before any policy is applied.
- **`allowed_emails` stays as the outer gate.** If set, nobody outside it gets
  in whatever their groups — a useful kill-switch while rolling this out.
- **`coordinator` becomes an ordinary group**, so `db.is_coordinator` collapses
  into `groups()` and the same list keeps governing μητρώα decisions and
  timetable edits.

## 3. Where membership lives — revised

The earlier recommendation was a hybrid: staff groups derived from
`timetable_staff.category`, coordinators and guests in settings. **The έκτακτοι
point weakens it enough that I would not start there.**

- `timetable_staff` has the right shape (34 people: ΔΕΠ 14, ΕΔΙΠ 2, ΕΤΕΠ 2,
  ΕΚΤΑΚΤΟΣ 15, ΑΛΛΟ 1), an `email` column, and an editing UI already.
- But **all 34 emails are NULL today**, secretaries and guests are not in the
  table at all, and the largest category — the 15 έκτακτοι — is exactly the
  group that may not be able to log in. Deriving groups from it would buy
  self-service maintenance for the people least able to use it.
- It would also make the login gate depend on `DATABASE_URL`, which is absent
  locally and on Streamlit Cloud. Pages 3, 4 and 8's public views must keep
  working there.

**Recommendation: settings-only to start**, with `groups()` as the single
lookup so the source can change later without touching a single page. Revisit
`timetable_staff` once the emails exist and §1 is settled.

## 4. Per-page audience — *proposal, please edit*

| Page | Today | Proposed | Note |
|---|---|---|---|
| 0 home | public | `PUBLIC` | login block; lists only the pages you can open |
| 1 περιγράμματα (legacy) | `@ihu.gr` | — | to be deleted |
| 3 εξεταστική | public | `PUBLIC` | no personal data |
| 4 εβδομαδιαίο (legacy) | public | `PUBLIC` | to be deleted once page 8 replaces it |
| 5 μητρώα v2 | `@ihu.gr` | `dep`, `coordinator` | ❓ do secretaries need read access? |
| 6 περιγράμματα v2 | `@ihu.gr` | `dep`, `edip`, `ektaktoi`, `secretariat` | anyone who teaches owns a περίγραμμα |
| 7 Εύδοξος | `@ihu.gr` | `dep`, `edip`, `ektaktoi` | one teacher per course picks the books |
| 8 πρόγραμμα v2 | public view, coordinator edit | `PUBLIC` view | unchanged; edits per §5 |

`st.navigation` takes a per-user page list, so a page someone cannot open
simply is not in their sidebar — better than a page that greets them with a
refusal.

## 5. Per-tab access — *proposal, please edit*

| Page | Tab | Proposed |
|---|---|---|
| 5 μητρώα | Εκλέκτορες ΑΠΕΛΛΑ · Γνωστικά αντικείμενα · Εξωτερικοί εκλέκτορες · Έλεγχος · Αναζήτηση | page audience |
| 5 μητρώα | Προετοιμασία \<έτους\> | `dep`; the decide/lock section stays `coordinator` (already is) |
| 6 περιγράμματα | Πίνακας · Στατιστικά · Αρχείο Word | page audience |
| 6 περιγράμματα | Επεξεργασία | teaching groups |
| 6 περιγράμματα | Αναφορές | ❓ `coordinator` only? |
| 7 Εύδοξος | Βιβλία ανά έτος · Έλεγχος διαθεσιμότητας | page audience |
| 7 Εύδοξος | Επεξεργασία | teaching groups |
| 7 Εύδοξος | Συντονιστής | `coordinator` (already) |
| 8 πρόγραμμα | Εβδομαδιαία προβολή · Πίνακας · Αιθουσιολόγιο · Ανά διδάσκοντα · Εξαγωγή Word | `PUBLIC` |
| 8 πρόγραμμα | Προετοιμασία · Προσωπικό · Αίθουσες | `coordinator` ❓ + `secretariat`? |

Two rules that matter more than the table:

- **`st.tabs` executes every tab's body on every rerun**, whatever is
  selected. A restricted tab must therefore not *build* its content — guarding
  only the label leaks data into the page and pays for the query as well.
- **Hiding UI is not a control.** Every write keeps its own check:
  `proposals_ui` decisions, `eudoxus_db.delete_year` / `lock_year`, the
  timetable edits. The group check in the UI is there to keep the page
  readable, not to enforce anything.

## 6. Order of work

1. `auth.groups` / `has_any` / `require`, with the implicit `ihu` group and
   `coordinator` folded in. Tests for the resolution rules, including
   "logged in but in no group" and a non-`@ihu.gr` address in `guest_emails`.
2. Convert the existing call sites to `require(IHU)` and
   `has_any("coordinator")` — **no behaviour change**, so it can land and be
   verified on its own.
3. Apply §4 and §5 once you have edited them, page by page.
4. Filter the `st.navigation` page list by group.
5. Only then consider moving staff groups into `timetable_staff`.

## 7. What I need from you

1. **§1: which login route** for colleagues without an IHU account? Everything
   else waits on this.
2. **§4 and §5**: correct my guesses, especially the ❓ rows — whether the
   secretariat edits the timetable, and whether μητρώα is ΔΕΠ-only.
3. Are ΕΔΙΠ/ΕΤΕΠ worth separate groups, or is one `staff` group enough beside
   `dep`? Every group is another env var to keep current on Railway.
