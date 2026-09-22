"""
Generates the school's Annexure K (level-distribution) workbook straight from
the Access database, filling the DBE's own template instead of building the
level counts by hand.

Requires: Python 3 (64-bit), pyodbc, openpyxl, and the 64-bit
"Microsoft Access Database Engine" ODBC driver (same one generate_top10.py
in the "Top 10 automation" folder needs - if you get a driver error on a
different PC, install the free "Microsoft Access Database Engine 2016
Redistributable", 64-bit version, from Microsoft).

HOW TO RUN NEXT TERM
1. Copy the new term's Access database (.mdb) into this folder (or point
   DB_PATH at it). Always use a COPY, never the live file.
2. Update TERM and YEAR below.
3. Get a blank-ish copy of that term's DBE Annexure K workbook for each
   grouping ("8-12" and "R-7") and point TEMPLATES at them. It's fine if the
   template still has old data in it (e.g. a leftover Exam-mark block from
   the last exam term) - this script overwrites every matching row's
   Composite AND Exam-mark blocks itself, it doesn't rely on what was
   already there.
4. Run:  python generate_annexure_k.py
   The finished workbook(s) appear in OUTPUT_FOLDER. Read the summary
   printed to the screen - it lists any subject labels it couldn't match to
   a database subject (those rows are left blank, same as "no marks yet").

WHAT THIS DOES AND WHY (see the STEP 1/2 verification this was built from)
----------------------------------------------------------------------------
Composite Mark block: the CURRENT term's own ReportMarks.Mark, banded into
the 7 CAPS achievement levels. Verified count-for-count against real DBE
submissions for Terms 2 and 3, grades 4 and 8.

Exam Mark block ("Exam marks Mark"): NOT this term's own exam - it's the
mark from the last term that actually had a written exam (Term 2 or Term 4).
On a T1/T3 run this pulls the prior T2/T4 report cycle instead of the
current one. Within that exam term, it isn't the subject's overall Mark
either - it's ONE specific assessment task:
  - if a task's description contains "exam" (e.g. FET's literal "Mid-year
    examination" task), that task wins outright, regardless of weighting;
  - otherwise, among tasks whose description contains "test" or
    "practical", the highest-weighted one is used;
  - if neither exists, there's no separate exam for this subject - the
    Exam-mark block mirrors that exam term's own Composite-style average.
This was verified task-by-task against real answer keys for grades 4, 8 and
10 (Intermediate, Senior and FET phases).

Known Senior Phase (Gr 7-9) exception, confirmed with the school: Natural
Sciences and Social Sciences have an internal "Control test" task but no
separate formal exam at this school in Senior Phase - their Exam-mark block
always mirrors the Composite, same as the other non-core Senior subjects
(EMS, Life Orientation, Technology, Creative Arts) which reach the same
result naturally because they only have one CASS task that term.

Level cut-offs (30/40/50/60/70/80) are read from the database's own
Ana2012EvaluationLevels table at startup, not hardcoded - see
load_level_cutoffs() below.
"""

import datetime
import os
import re

import pyodbc
from openpyxl import load_workbook

try:
    # The real password lives in db_secret.py, which is gitignored and never
    # committed. Copy db_secret.example.py to db_secret.py and fill it in.
    from db_secret import DB_PASSWORD
except ImportError:
    raise RuntimeError(
        "db_secret.py not found. Copy db_secret.example.py to db_secret.py "
        "in this folder and put the real database password in it."
    )

# ---------------------------------------------------------------------------
# SETTINGS - change these each term
# ---------------------------------------------------------------------------
DB_PATH = (
    r"C:\Users\user\Desktop\Jarin Werk\2026\Claude playground\Annexure K automation"
    r"\700400012 Balmoral College 22 Sept\700400012 Balmoral College.mdb"
)
DATA_YEAR = 2026
TERM = 3
GROUPINGS = ["8-12", "R-7"]  # generate all; or e.g. ["R-7"] for one

TEMPLATES = {
    "8-12": (
        r"C:\Users\user\Desktop\Jarin Werk\2026\Claude playground\Annexure K automation"
        r"\Annexure K Term 3 2026 8-12 Template.xlsx"
    ),
    "R-7": (
        r"C:\Users\user\Desktop\Jarin Werk\2026\Claude playground\Annexure K automation"
        r"\Annexure K Term 3 2026 R-7 Template.xlsx"
    ),
}
OUTPUT_FOLDER = (
    r"C:\Users\user\Desktop\Jarin Werk\2026\Claude playground\Annexure K automation\Output"
)

# EMIS number of our own school - the R-7 template ships as a province-wide
# "Level data all Grades" dump with every school in it; we only touch rows
# whose EMIS column matches this, and leave every other school's rows alone.
SCHOOL_EMIS = "700400012"

# ---------------------------------------------------------------------------
# Fixed layout - matches the DBE Annexure K template column positions,
# should not normally change. Verified directly against real submitted
# workbooks; the "8-12" 3-block layout is 37 columns wide, "R-7" is 38
# (it has one extra duplicated "check difference" column between its blank
# middle block and its Exam-mark block).
#   grade_col / subject_col : where to read the row's grade and subject
#   composite / exam        : (avg_col, first_level_col, total_col) for
#                              each block - the check-difference column is
#                              deliberately never touched, it's left as
#                              whatever formula/value the template shipped
#                              with
#   subject_has_grade_suffix: whether the template's own Subject cell
#                              already reads e.g. "Mathematics (Gr 08)"
#                              (8-12) or just "Mathematics" (R-7)
# ---------------------------------------------------------------------------
LAYOUTS = {
    "8-12": dict(
        grade_col=5, subject_col=7,
        composite=(9, 10, 17),
        exam=(28, 29, 36),
        subject_has_grade_suffix=True,
        first_data_row=3,
    ),
    "R-7": dict(
        grade_col=5, subject_col=6,
        composite=(8, 9, 16),
        exam=(29, 30, 37),
        subject_has_grade_suffix=False,
        first_data_row=3,
    ),
}

# Grade -> report-cycle phase, per the CAPS phase structure this school uses
# (same mapping generate_top10.py uses).
PHASE_BY_GRADE = {
    0: "Foundation", 1: "Foundation", 2: "Foundation", 3: "Foundation",
    4: "Intermediate", 5: "Intermediate", 6: "Intermediate",
    7: "Senior", 8: "Senior", 9: "Senior",
    10: "FET", 11: "FET", 12: "FET",
}
FOUNDATION_GRADES = {0, 1, 2, 3}  # no Exam-mark block at all for these

# Grade 8/9 subjects confirmed (2026-09-21, with the school) to have no
# separate formal exam despite an internally-named "Control test" CASS task -
# their Exam-mark block always mirrors the Composite block. This does NOT
# extend to grade 7, even though it's also Senior Phase curriculum-wise:
# grade 7 lives in the R-7 workbook and was verified to follow the
# Intermediate-style "isolate the highest-weighted task" rule instead
# (exact match on both subjects once isolated) - so this set is applied by
# grade (8/9 only), not by curriculum phase.
GR_8_9_MIRROR_SUBJECTS = {"natural sciences", "social sciences"}

# Grade 12 subjects confirmed to carry NO Exam-mark data at all (left
# blank, not even mirroring Composite) - Life Orientation isn't part of the
# NSC exit-year exam/APS framework, unlike Gr 10/11 where its own
# "Mid-year examination" task IS the exam mark (verified exact match, same
# as every other FET subject). Note: Life Orientation Gr 8 isolates its
# "Control test" task cleanly via the normal rule below (exact match
# verified); Gr 9 does NOT match under the same rule despite an identical
# task structure, and no alternative formula tried matches either - that
# one grade/subject combination is a known, unresolved discrepancy flagged
# in the README rather than papered over with a guessed rule.
GR_12_BLANK_EXAM_SUBJECTS = {"life orientation"}


def connect():
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={DB_PATH};"
        rf"PWD={DB_PASSWORD};"
        r"ReadOnly=True;"
    )
    return pyodbc.connect(conn_str, readonly=True)


def load_level_cutoffs(cursor):
    """Read the 6 achievement-level boundaries (Level 2..7 lower bounds)
    from the database's own Ana2012EvaluationLevels table, current version,
    rather than assuming 30/40/50/60/70/80."""
    cursor.execute(
        "SELECT MarkFrom FROM Ana2012EvaluationLevels "
        "WHERE EvalVer = 3 AND GradeFrom = 7 AND GradeTo = 12 "
        "AND Code BETWEEN 2 AND 7 ORDER BY Code"
    )
    edges = [float(r[0]) for r in cursor.fetchall()]
    if len(edges) != 6:
        raise RuntimeError(
            "Could not read the 6 level cut-offs from Ana2012EvaluationLevels - "
            "check the table hasn't changed."
        )
    return edges


def band_index(mark, edges):
    for i, edge in enumerate(edges):
        if mark < edge:
            return i
    return 6


def band_counts(values, edges):
    levels = [0] * 7
    for v in values:
        levels[band_index(v, edges)] += 1
    return levels


def normalize_grade(raw):
    """'R' / 'r' / 0 / '0' -> 0 (Reception); everything else -> int."""
    s = str(raw).strip().upper()
    return 0 if s == "R" else int(s)


def grade_phase(grade_num):
    return PHASE_BY_GRADE[grade_num]


# ---------------------------------------------------------------------------
# Subject-name matching: the template's subject labels don't always spell
# things exactly like Subjects.Name does (e.g. "Natural Science &
# Technology" in the template vs "Natural Sciences and Technology (Gr 04)"
# in the database - singular/plural and "&"/"and" both differ). Build a
# normalised index once and match loosely rather than exactly.
# ---------------------------------------------------------------------------
def _norm(s):
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _norm_loose(s):
    return _norm(s).replace("s", "")  # tolerate singular/plural mismatches


def build_subject_index(cursor):
    cursor.execute("SELECT Id, Name FROM Subjects")
    exact, loose = {}, {}
    for sid, name in cursor.fetchall():
        exact.setdefault(_norm(name), sid)
        loose.setdefault(_norm_loose(name), sid)
    return exact, loose


def find_subject_id(subject_index, base_label, grade_num):
    exact_idx, loose_idx = subject_index
    grade_tag = "Gr R" if grade_num == 0 else f"Gr {grade_num:02d}"
    candidates = [base_label, base_label.replace("&", "and"), base_label.replace(" and ", " & ")]
    for cand in candidates:
        key = _norm(f"{cand} ({grade_tag})")
        if key in exact_idx:
            return exact_idx[key]
    for cand in candidates:
        key = _norm_loose(f"{cand} ({grade_tag})")
        if key in loose_idx:
            return loose_idx[key]
    return None


def strip_grade_suffix(label):
    return re.sub(r"\s*\(Gr\s*\d+\)\s*$", "", label).strip()


# ---------------------------------------------------------------------------
# Report-cycle helpers
# ---------------------------------------------------------------------------
def report_id_for(cursor, cache, term, phase, year):
    key = (term, phase, year)
    if key not in cache:
        cursor.execute(
            "SELECT CycleId FROM ReportCycles WHERE Datayear = ? AND Term = ? AND Phase = ?",
            str(year), term, phase,
        )
        row = cursor.fetchone()
        cache[key] = row[0] if row else None
    return cache[key]


def exam_term_for(term, year):
    """Which (term, year) actually has exam data. T2/T4 have their own;
    T1 looks back to the previous year's T4, T3 looks back to this year's T2."""
    if term in (2, 4):
        return term, year
    if term == 1:
        return 4, year - 1
    if term == 3:
        return 2, year
    raise ValueError(f"Unexpected term {term}")


# ---------------------------------------------------------------------------
# Mark calculation
# ---------------------------------------------------------------------------
def composite_for_report(cursor, edges, report_id, subject_id):
    cursor.execute(
        "SELECT Mark FROM ReportMarks WHERE ReportId = ? AND SubjectId = ?",
        report_id, subject_id,
    )
    marks = [float(r.Mark) for r in cursor.fetchall() if r.Mark is not None]
    if not marks:
        return None, [0] * 7
    return sum(marks) / len(marks), band_counts(marks, edges)


def exam_task_criterion_id(cursor, subject_id, exam_year, exam_term):
    """Pick the task that represents the formal exam for this subject that
    term: an explicit "exam"-named task wins outright; otherwise the
    highest-weighted "test"/"practical" task; otherwise None (no exam)."""
    cursor.execute(
        "SELECT CriterionID, Description, Weighting FROM SubjectCriteria "
        "WHERE Subjectid = ? AND DataYear = ? AND SubHeading = ?",
        subject_id, str(exam_year), f"Term{exam_term}",
    )
    tasks = cursor.fetchall()
    exam_tasks = [t for t in tasks if re.search(r"exam", t.Description, re.I)]
    pool = exam_tasks or [t for t in tasks if re.search(r"test|practical", t.Description, re.I)]
    if not pool:
        return None
    pool.sort(key=lambda t: (-t.Weighting, -t.CriterionID))
    return pool[0].CriterionID


def exam_block_for(cursor, edges, subject_id, base_label, grade_num, phase,
                    exam_year, exam_term, exam_report_id):
    label_norm = _norm(base_label).rstrip("s")

    if grade_num == 12 and label_norm in {_norm(s).rstrip("s") for s in GR_12_BLANK_EXAM_SUBJECTS}:
        return None, [0] * 7  # confirmed: no Exam-mark data for these in Gr 12

    if grade_num in (8, 9) and label_norm in {_norm(s).rstrip("s") for s in GR_8_9_MIRROR_SUBJECTS}:
        return composite_for_report(cursor, edges, exam_report_id, subject_id)

    criterion_id = exam_task_criterion_id(cursor, subject_id, exam_year, exam_term)
    if criterion_id is None:
        return composite_for_report(cursor, edges, exam_report_id, subject_id)

    cursor.execute(
        "SELECT Mark, Criterionscore FROM LearnerCass WHERE Subjectid = ? AND CriterionId = ?",
        subject_id, criterion_id,
    )
    vals = [(float(r.Mark) / float(r.Criterionscore)) * 100
            for r in cursor.fetchall() if r.Criterionscore]
    if not vals:
        return composite_for_report(cursor, edges, exam_report_id, subject_id)
    return sum(vals) / len(vals), band_counts(vals, edges)


# ---------------------------------------------------------------------------
# Sheet writing
# ---------------------------------------------------------------------------
def write_block(ws, row, block, avg, levels):
    avg_col, lvl_col, total_col = block
    ws.cell(row, avg_col, avg if avg is not None else None)
    for i, v in enumerate(levels):
        ws.cell(row, lvl_col + i, v if v else None)
    ws.cell(row, total_col, sum(levels))


def process_row(cursor, ws, row, layout, grade_raw, subject_label,
                 subject_index, report_id_cache, edges, unmatched):
    try:
        grade_num = normalize_grade(grade_raw)
    except ValueError:
        return
    phase = grade_phase(grade_num)

    base_label = (strip_grade_suffix(subject_label)
                  if layout["subject_has_grade_suffix"] else subject_label)
    subject_id = find_subject_id(subject_index, base_label, grade_num)
    if subject_id is None:
        unmatched.append((row, subject_label, grade_raw))
        return

    comp_report_id = report_id_for(cursor, report_id_cache, TERM, phase, DATA_YEAR)
    if comp_report_id is None:
        return
    comp_avg, comp_levels = composite_for_report(cursor, edges, comp_report_id, subject_id)
    write_block(ws, row, layout["composite"], comp_avg, comp_levels)

    if grade_num in FOUNDATION_GRADES:
        return  # Foundation Phase never gets an Exam-mark block

    exam_term, exam_year = exam_term_for(TERM, DATA_YEAR)
    exam_report_id = report_id_for(cursor, report_id_cache, exam_term, phase, exam_year)
    if exam_report_id is None:
        return
    exam_avg, exam_levels = exam_block_for(
        cursor, edges, subject_id, base_label, grade_num, phase,
        exam_year, exam_term, exam_report_id,
    )
    write_block(ws, row, layout["exam"], exam_avg, exam_levels)


def process_8_12(cursor, subject_index, report_id_cache, edges):
    layout = LAYOUTS["8-12"]
    wb = load_workbook(TEMPLATES["8-12"])
    ws = wb.active
    unmatched = []
    rows_done = 0
    for row in range(layout["first_data_row"], ws.max_row + 1):
        grade_raw = ws.cell(row, layout["grade_col"]).value
        subject_label = ws.cell(row, layout["subject_col"]).value
        if grade_raw is None or not subject_label:
            continue
        process_row(cursor, ws, row, layout, grade_raw, subject_label,
                    subject_index, report_id_cache, edges, unmatched)
        rows_done += 1
    return wb, rows_done, unmatched


def process_r7(cursor, subject_index, report_id_cache, edges):
    layout = LAYOUTS["R-7"]
    wb = load_workbook(TEMPLATES["R-7"])
    ws = wb.active
    unmatched = []
    rows_done = 0
    for row in range(layout["first_data_row"], ws.max_row + 1):
        emis = ws.cell(row, 3).value
        if emis is None or str(emis).strip() != SCHOOL_EMIS:
            continue
        grade_raw = ws.cell(row, layout["grade_col"]).value
        subject_label = ws.cell(row, layout["subject_col"]).value
        if grade_raw is None or not subject_label:
            continue
        process_row(cursor, ws, row, layout, grade_raw, subject_label,
                    subject_index, report_id_cache, edges, unmatched)
        rows_done += 1
    return wb, rows_done, unmatched


PROCESSORS = {"8-12": process_8_12, "R-7": process_r7}


def generate(db_path=None, db_password=None, data_year=None, term=None,
             groupings=None, templates=None, output_folder=None):
    """Run a full Annexure K generation. Any argument left as None falls back
    to the module-level SETTINGS above. Returns a list of saved workbook
    paths. Used both by main() (CLI run with the settings at the top of this
    file) and by the web UI (Top 10 automation/webapp/server.py), which
    passes its own values in."""
    global DB_PATH, DB_PASSWORD, DATA_YEAR, TERM, GROUPINGS, TEMPLATES, OUTPUT_FOLDER
    if db_path is not None:
        DB_PATH = db_path
    if db_password is not None:
        DB_PASSWORD = db_password
    if data_year is not None:
        DATA_YEAR = data_year
    if term is not None:
        TERM = term
    if groupings is not None:
        GROUPINGS = groupings
    if templates is not None:
        TEMPLATES = templates
    if output_folder is not None:
        OUTPUT_FOLDER = output_folder

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    cnxn = connect()
    cursor = cnxn.cursor()

    edges = load_level_cutoffs(cursor)
    print(f"Level cut-offs read from the database: {edges}")
    subject_index = build_subject_index(cursor)
    report_id_cache = {}

    exam_term, exam_year = exam_term_for(TERM, DATA_YEAR)
    print(f"Term {TERM} {DATA_YEAR}: Composite = Term {TERM} {DATA_YEAR}'s own marks; "
          f"Exam-mark block = Term {exam_term} {exam_year}\n")

    stamp = datetime.date.today().isoformat()
    out_paths = []
    for grouping in GROUPINGS:
        if grouping not in TEMPLATES:
            print(f"Skipping {grouping}: no TEMPLATES entry configured.")
            continue
        wb, rows_done, unmatched = PROCESSORS[grouping](cursor, subject_index, report_id_cache, edges)

        out_name = f"Annexure K Term {TERM} {DATA_YEAR} {grouping} - generated {stamp}.xlsx"
        out_path = os.path.join(OUTPUT_FOLDER, out_name)
        wb.save(out_path)
        out_paths.append(out_path)

        print(f"[{grouping}] {rows_done} row(s) processed -> {out_path}")
        if unmatched:
            print(f"  WARNING: {len(unmatched)} row(s) had a subject label that didn't match "
                  f"any database subject and were left blank:")
            for row, label, grade in unmatched:
                print(f"    row {row}: '{label}' (grade {grade})")
        print()

    cnxn.close()
    print(f"Run finished: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")

    return out_paths


def main():
    generate()


if __name__ == "__main__":
    main()
