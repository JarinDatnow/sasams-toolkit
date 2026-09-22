# SASAMS Toolkit

**Automation scripts for South African schools running SASAMS and D6.**

Built primarily as a portable, version-controlled backup of my own end-of-term
workflow — if my machine dies or I change schools, setup is a git clone away
instead of a manual rebuild. Shared publicly in case it helps other SA
schools running SASAMS.

## What is this?

Two generators that query your SASAMS Access database directly and produce
the finished workbook, plus one shared local web UI to run either of them:

- **[`Top 10 automation/`](Top%2010%20automation/)** — Generates the
  school's "Top 10" achiever workbook (one sheet per class/grade, correctly
  ordered, ties handled), replacing the manual process of exporting reports
  from D6 one at a time.
- **[`Annexure K automation/`](Annexure%20K%20automation/)** — Generates the
  DBE Annexure K (achievement-level distribution) workbook for both the
  "8-12" and "R-7" groupings, filling the DBE's own template instead of
  building the level counts by hand.
- **`Top 10 automation/webapp/`** — A local-only web page (no network access
  needed, nothing leaves your PC) that runs either generator: pick TOP 10 or
  ANNEXURE K at the top of the page, fill in the fields, click GENERATE.
  Launch it by double-clicking `Top 10 automation/Launch Top 10 Club.vbs`.

## What is SASAMS?

SASAMS (South African School Administration and Management System) is the
Department of Basic Education's system for managing learner data, marks,
attendance, and promotion across South African schools. Schools typically
interact with it through an Access `.mdb` database file that syncs with the
provincial system.

## Requirements

- **Windows** (required for the Access database driver)
- **Python 3 (64-bit)**
- The 64-bit **Microsoft Access Database Engine** ODBC driver (if you get a
  driver error, install the free "Microsoft Access Database Engine 2016
  Redistributable", 64-bit version, from Microsoft)
- Read access to your school's SASAMS `.mdb` database file

### Python packages

```bash
pip install pyodbc openpyxl
```

## Setup

1. **Clone this repo**
   ```bash
   git clone https://github.com/JarinDatnow/sasams-toolkit.git
   ```

2. **Add your database**
   Copy the term's Access database (`.mdb`) into the relevant folder (or
   point `DB_PATH` at the top of the generator script at it). Always use a
   **copy**, never the live file.

3. **Create your `db_secret.py`**
   Each generator folder needs its own real database password, kept out of
   git. In each of `Top 10 automation/` and `Annexure K automation/`:
   ```bash
   copy db_secret.example.py db_secret.py
   ```
   then edit `db_secret.py` and fill in the real password.

4. **Add your template(s)**
   - Top 10: `TEMPLATE  top 10.xlsx` ships in `Top 10 automation/` already.
   - Annexure K: `Top 10 automation`'s sibling folder ships the school's own
     "8-12" templates; you'll need to supply your term's DBE "R-7" template
     yourself (it's a large province-wide file, not committed here) and
     point `TEMPLATES["R-7"]` at it.

5. **Update TERM / YEAR** at the top of whichever generator script(s) you're
   running.

## Usage

### Web UI (recommended)
Double-click `Top 10 automation/Launch Top 10 Club.vbs`. It starts a local
server and opens your browser to a themed page with a TOP 10 / ANNEXURE K
toggle at the top — pick one, fill in the database/term/output fields (and
groupings/templates for Annexure K), click GENERATE. The launcher always
stops any earlier copy of the server before starting a fresh one, so it's
safe to double-click again any time.

### Command line
```bash
cd "Top 10 automation"
python generate_top10.py
```
```bash
cd "Annexure K automation"
python generate_annexure_k.py
```
Each prints a summary to the screen (learner/row counts, ties, unmatched
subjects, warnings) — read it before submitting the workbook.

## How the Top 10 percentage is calculated

It's tempting to average `ReportMarks.Mark` (SASAMS's own already-rounded,
whole-number mark per subject) directly. **This is wrong** — it can land a
full percentage point off what D6's own "Top Achievers" report actually
prints. D6 averages the **unrounded** per-subject percentage instead, built
from the continuous-assessment task tables (`LearnerCass` +
`SubjectCriteria`), with tasks matched to a term via the explicit
`SubjectCriteria.SubHeading` label (`'Term1'`/`'Term2'`/...), not by date.
See the `unrounded_subject_pct()` docstring in `generate_top10.py` for the
exact formula. Verified against real D6-exported Top Achiever data with zero
mismatches once this was in place.

## How Annexure K is calculated

The Composite Mark block is the current term's own marks, banded into the 7
CAPS achievement levels (cut-offs read from the database itself, not
hardcoded). The Exam Mark block is *not* the current term's own exam — it's
one specific assessment task from the last term that actually had a written
exam. See [`Annexure K automation/README.md`](Annexure%20K%20automation/README.md)
for the full rule set, confirmed exceptions, and known open items — this is
the more intricate of the two generators and worth reading before trusting
its output on a school you haven't verified it against yet.

## Security note

`db_secret.py` (in both generator folders) holds your real database
password and is gitignored — never commit it. The `.gitignore` here also
excludes the `.mdb` database itself, generated `Output/` workbooks, and any
Annexure K workbook that isn't a blank school-specific template, since those
contain real learner records. Ensure you comply with your school's data
policies and POPIA requirements when handling learner information.

## Reference

`SASAMS_map.json` at the repo root is a full pre-dumped schema (every table
and column SASAMS exposes) — useful if you want to write your own queries
against the database.

## Contributing

Found a bug? Got a better query? Open an issue or PR. This was built for one
school but should work for any SA school running SASAMS.

## Disclaimer

This toolkit is provided as-is for educational and administrative purposes.
It only reads data from your local SASAMS database — it does not modify any
records.

## License

MIT — do whatever you want with it.
