# Annexure K automation

Generates the DBE Annexure K (level-distribution) workbook straight from the
SA-SAMS Access database, for both the "8-12" and "R-7" groupings.

## How to run it

1. Copy the new term's Access database (`.mdb`) into this folder (or point
   `DB_PATH` in `generate_annexure_k.py` at it). Always use a **copy**, never
   the live file.
2. Update `TERM` and `DATA_YEAR` at the top of `generate_annexure_k.py`.
3. Point `TEMPLATES["8-12"]` and `TEMPLATES["R-7"]` at that term's Annexure K
   workbook. It's fine if it still has old data in it (e.g. a leftover
   Exam-mark block from the last exam term) - the script overwrites every
   matching row's Composite **and** Exam-mark blocks itself; it never reads
   the template's existing numbers.
4. Run:
   ```
   python generate_annexure_k.py
   ```
5. The finished workbook(s) land in `Output\`. Read the printed summary - it
   lists any subject label the script couldn't match to a database subject
   (that row is left blank, same as "no marks captured yet").

Requires Python 3 (64-bit), `pyodbc`, `openpyxl`, and the 64-bit "Microsoft
Access Database Engine" ODBC driver (same one `generate_top10.py` needs).

## What it calculates, and why

**Composite Mark block:** the current term's own `ReportMarks.Mark`, banded
into the 7 CAPS achievement levels. Verified count-for-count against real
submitted Annexure K workbooks for Term 2 and Term 3, grades 4, 7, 8 and 10.

**Exam Mark block:** *not* the current term's own exam - it's the mark from
the last term that actually had a written exam (Term 2 or Term 4). A Term 1
run looks back to the previous year's Term 4; a Term 3 run looks back to
that year's Term 2. Within that exam term, it isn't the subject's overall
Mark either - it's one specific assessment task:
- a task whose description contains "exam" wins outright (this is how FET's
  literal "Mid-year examination" task gets picked), regardless of weighting;
- otherwise, the highest-weighted task whose description contains "test" or
  "practical" is used;
- otherwise there's no separate exam for that subject - the block mirrors
  that exam term's own Composite-style average instead.

Level cut-offs (30/40/50/60/70/80) are read from the database's own
`Ana2012EvaluationLevels` table at startup (see `load_level_cutoffs()`), not
hardcoded.

### Confirmed exceptions

- **Grade 8/9 Natural Sciences and Social Sciences** (Senior Phase, "8-12"
  workbook only - **not** Grade 7, which lives in the "R-7" workbook and
  follows the normal isolate-a-task rule): no separate exam at this school
  despite an internally-named "Control test" CASS task - Exam-mark mirrors
  Composite. Confirmed with the school 2026-09-21.
- **Grade 12 Life Orientation**: confirmed blank - no Exam-mark data at all,
  not even a Composite mirror. Grade 10/11 Life Orientation isolates its own
  "Mid-year examination" task normally, same as every other FET subject.

### Known open item

**Grade 9 Life Orientation's Exam-mark block doesn't match.** Its task
structure is identical to Grade 8's (a "Control test" + a "PET" task), and
Grade 8 isolates the Control test task with an exact match - but the same
approach for Grade 9 doesn't match the historical answer key, and no other
formula tried (mirroring Composite, isolating the other task, weighted
combinations) matches either. Left on the normal isolate-a-task rule as the
best-evidenced default; flag this row for a manual check each run.

### Expected small discrepancies (not bugs)

- A handful of rows differ from the historical answer-key files by one or
  two learners in one band. The database has moved on since those files
  were frozen (new marks captured, roster changes) - this is drift, not a
  calculation error. `Mathematics (Gr 10)`'s Exam block is a recurring
  example (short by ~5 students in the lowest band - those students simply
  aren't in this database's Term 2 record for that subject at all).
- The historical **Term 3 2026 8-12** answer key has Grade 10-12 Composite
  entirely blank; this script produces real numbers for those rows because
  the live database now has Term 3 FET marks that weren't captured yet when
  that answer key was frozen. Worth a quick sanity check against the actual
  current roster before submitting, since it's new territory this script
  hasn't been checked against a real answer key for.
- The historical **Term 3 2026 R-7** answer key has Grade 7 Social Sciences
  and Sepedi First Additional Language swapped with each other (each row's
  numbers exactly match the *other* subject's answer) - a data-entry bug in
  that one historical file, not something this script reproduces.

## Files

- `generate_annexure_k.py` - the generator (this script)
- `Annexure K Term X {year} {grouping} Template.xlsx` - blank/prior-term
  workbook the script fills in
- `Output\` - generated workbooks land here, timestamped by run date
