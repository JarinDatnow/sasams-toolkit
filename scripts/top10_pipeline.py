"""
=============================================================================
  SASAMS TOP 10 — FULL PIPELINE
  Query SASAMS -> Generate one Top 10 workbook -> Print it

  One command. One workbook, one sheet per class/grade. Done.
=============================================================================
  Setup:
    1. pip install pyodbc openpyxl pywin32
    2. Copy config.example.py to config.py and fill in your details
    3. Put TEMPLATE__top_10.xlsx in the templates/ folder
    4. Run:  python scripts/top10_pipeline.py

  Flags:
    --no-print     Generate the workbook only, don't print
    --discover     Show database tables and columns

  How the PERCENTAGE is calculated (read this before changing the query)
  -------------------------------------------------------------------------
  It is tempting to just average ReportMarks.Mark (one already-rounded
  whole-number mark per subject) or LearnerPromotion.LearnerAverage (a
  precomputed promotion average). Both were tried and both are WRONG - they
  can be off by a percentage point from what SASAMS/D6 actually prints,
  because SASAMS rounds each subject's mark to a whole number for display
  *before* ReportMarks ever sees it, then D6's own Top Achievers report
  averages the UNROUNDED per-subject percentage, not the rounded one.

  The unrounded per-subject percentage lives one level deeper, in the
  continuous-assessment task tables:
    - LearnerCass          one row per task (test, assignment, exam...)
                            per learner per subject, with Mark / Criterionscore
    - SubjectCriteria       defines each task's Weighting toward the subject
                            total, and tags it to a term via SubHeading
                            ('Term1' / 'Term2' / 'Term3' / 'Term4')
  Subject % = sum(Mark/Criterionscore * Weighting) / sum(Weighting) * 100,
  and the learner's term % = round-half-up of the average of those
  unrounded subject percentages. Verified against 306 real D6-exported
  Top Achievers rows for two different terms with zero mismatches - see
  the README for how this was confirmed.
=============================================================================
"""

import copy
import datetime
import os
import sys
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pyodbc

# ── Load config ─────────────────────────────────────────────────────────────
# config.py lives at the repo root (one level up from this script), so make
# sure it's importable regardless of which directory this script is run from.
SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

try:
    from config import DB_PATH, DB_PASSWORD, DATA_YEAR, TERM, OUTPUT_FOLDER
except ImportError:
    print("[FAIL] config.py not found.")
    print("       Copy config.example.py to config.py and fill in your values.")
    sys.exit(1)

TEMPLATE_NAME = "TEMPLATE__top_10.xlsx"
AUTO_PRINT = True


def parse_term_number(term_setting):
    """Accepts TERM as an int (3) or a string ('TERM 3', 'Term3', '3')."""
    if isinstance(term_setting, int):
        return term_setting
    digits = "".join(ch for ch in str(term_setting) if ch.isdigit())
    if not digits:
        raise ValueError(
            f"Could not read a term number out of TERM={term_setting!r} in config.py"
        )
    return int(digits)


YEAR = str(DATA_YEAR)
TERM_NUM = parse_term_number(TERM)
TERM_LABEL = f"TERM {TERM_NUM}"

# ---------------------------------------------------------------------------
# Fixed layout - matches TEMPLATE__top_10.xlsx, should not normally change
# ---------------------------------------------------------------------------
TEMPLATE_SHEET_NAME = "Worksheet"
FIRST_DATA_ROW = 13
LAST_TEMPLATE_DATA_ROW = 22  # template ships 10 pre-styled rows: 13-22
DATA_ROW_HEIGHT = 28.5
COLUMNS = {"nr": "A", "learner_no": "B", "surname": "C", "name": "D", "pct": "E"}
TOP_N = 10

# Grade -> report-cycle phase, per the CAPS phase structure (adjust here if
# your school's SASAMS setup groups phases differently)
PHASE_BY_GRADE = {
    0: "Foundation", 1: "Foundation", 2: "Foundation", 3: "Foundation",
    4: "Intermediate", 5: "Intermediate", 6: "Intermediate",
    7: "Senior", 8: "Senior", 9: "Senior",
    10: "FET", 11: "FET",
}

# Long-name shrink threshold: names longer than this (in the SURNAME or
# PREFERRED NAME column) get shrink-to-fit turned on so they don't clip
# when printed.
SHRINK_NAME_LEN = 11

# Sanity-check window for task dates, see check_task_dates() below.
TASK_DATE_WARNING_BUFFER_DAYS = 14


def connect():
    driver = "{Microsoft Access Driver (*.mdb, *.accdb)}"
    conn_str = f"DRIVER={driver};DBQ={os.path.abspath(DB_PATH)};PWD={DB_PASSWORD};ReadOnly=True;"
    try:
        conn = pyodbc.connect(conn_str, readonly=True)
        print("[OK] Connected to database (read-only)")
        return conn
    except (pyodbc.Error, UnicodeDecodeError) as e:
        print(f"[FAIL] Could not connect: {e}")
        sys.exit(1)


def discover_schema(conn):
    cursor = conn.cursor()
    print("\n" + "=" * 60)
    print("  DATABASE SCHEMA DISCOVERY")
    print("=" * 60)
    tables_of_interest = [
        "Learner_Info", "Classes", "ReportMarks", "ReportCycles",
        "SchoolTerms", "LearnerCass", "SubjectCriteria", "Subjects",
    ]
    for table in cursor.tables(tableType="TABLE"):
        tname = table.table_name
        if tname.startswith("MSys"):
            continue
        marker = "  <<<" if tname in tables_of_interest else ""
        print(f"\n  TABLE: {tname}{marker}")
        try:
            cursor.execute(f"SELECT TOP 1 * FROM [{tname}]")
            for d in cursor.description:
                print(f"    {d[0]:30s} {d[1].__name__}")
        except pyodbc.Error as e:
            print(f"    (could not read columns: {e})")
    print("\n" + "=" * 60)
    print("A full pre-generated dump of the whole schema also lives in")
    print("SASAMS_map.json at the repo root if you'd rather grep than query.")
    print("=" * 60)


def round_half_up(value):
    return int(Decimal(repr(value)).quantize(0, rounding=ROUND_HALF_UP))


def get_report_id_map(cursor):
    """phase name -> ReportId (CycleId) for the configured TERM/DATA_YEAR."""
    cursor.execute(
        "SELECT Phase, CycleId FROM ReportCycles WHERE Datayear = ? AND Term = ?",
        YEAR, TERM_NUM,
    )
    mapping = {row.Phase: row.CycleId for row in cursor.fetchall()}
    missing = set(PHASE_BY_GRADE.values()) - set(mapping)
    if missing:
        raise RuntimeError(
            f"No ReportCycles row found for Term {TERM_NUM} {YEAR}, phase(s): {missing}. "
            "Check TERM/DATA_YEAR in config.py and that this report cycle exists."
        )
    return mapping


def get_classes(cursor):
    """(grade, classid, classname) for grades R(0)-6, ordered for output."""
    cursor.execute(
        "SELECT ClassId, Grade, ClassName FROM Classes WHERE Grade BETWEEN 0 AND 6 "
        "ORDER BY Grade, ClassName"
    )
    return [(row.Grade, row.ClassId, row.ClassName) for row in cursor.fetchall()]


def get_term_dates(cursor):
    """(StartDate, EndDate) for the configured TERM/DATA_YEAR, from SchoolTerms.
    Used only for the check_task_dates() sanity check below - the actual
    mark calculation matches tasks to the term via SubjectCriteria.SubHeading
    (see unrounded_subject_pct), not by date."""
    cursor.execute(
        "SELECT StartDate, EndDate FROM SchoolTerms WHERE CurrentYear = ? AND Term = ?",
        YEAR, TERM_NUM,
    )
    row = cursor.fetchone()
    if row is None:
        raise RuntimeError(
            f"No SchoolTerms row found for Term {TERM_NUM} {YEAR}. "
            "Check TERM/DATA_YEAR in config.py."
        )
    return row.StartDate, row.EndDate


def check_task_dates(cursor, term_start, term_end):
    """
    Sanity check, not part of the mark calculation itself: list every
    assessment task tagged as this term (SubjectCriteria.SubHeading =
    'TermN') whose own DateAdded falls more than TASK_DATE_WARNING_BUFFER_DAYS
    outside the term's official [StartDate, EndDate] window (from
    SchoolTerms). A task like that isn't excluded from the calculation -
    SubHeading is trusted as the authoritative term link - but a mismatch
    this large usually means either the task was captured very late/early,
    or (as with a "SBA Year Mark" rollup row) it's a summary task rather
    than a real per-term assessment, which is worth a human look.
    """
    buffer = datetime.timedelta(days=TASK_DATE_WARNING_BUFFER_DAYS)
    lower, upper = term_start - buffer, term_end + buffer
    cursor.execute(
        """
        SELECT sc.Subjectid, s.Name, sc.CriterionID, sc.Description, sc.DateAdded, sc.Weighting
        FROM SubjectCriteria sc LEFT JOIN Subjects s ON sc.Subjectid = s.Id
        WHERE sc.DataYear = ? AND sc.SubHeading = ?
          AND (sc.DateAdded < ? OR sc.DateAdded > ?)
        ORDER BY sc.DateAdded
        """,
        YEAR, f"Term{TERM_NUM}", lower, upper,
    )
    return cursor.fetchall()


def unrounded_subject_pct(cursor, learner_id, subject_id):
    """See the module docstring for why this exists instead of just
    averaging ReportMarks.Mark. Returns None if no task breakdown exists
    for this subject this term (falls back to the rounded Mark)."""
    cursor.execute(
        """
        SELECT lc.Mark, lc.Criterionscore, sc.Weighting
        FROM LearnerCass lc
        INNER JOIN SubjectCriteria sc
            ON lc.CriterionId = sc.CriterionID
           AND lc.Subjectid = sc.Subjectid
           AND lc.Datayear = sc.DataYear
        WHERE lc.Learnerid = ? AND lc.Subjectid = ? AND lc.Datayear = ?
          AND sc.SubHeading = ?
        """,
        learner_id, subject_id, YEAR, f"Term{TERM_NUM}",
    )
    rows = cursor.fetchall()
    if not rows:
        return None
    total_weight = sum(r.Weighting for r in rows)
    if not total_weight:
        return None
    weighted = sum(
        (r.Mark / r.Criterionscore) * r.Weighting for r in rows if r.Criterionscore
    )
    return weighted / total_weight * 100


def fetch_learners(cursor, grade, class_id, report_id):
    """Return sorted list of learner dicts for one class/grade report."""
    if class_id is None:
        cursor.execute(
            "SELECT ID, AccessionNo, SName, NickName, FName FROM Learner_Info "
            "WHERE Grade = ? AND Status = 'C'",
            grade,
        )
    else:
        cursor.execute(
            "SELECT ID, AccessionNo, SName, NickName, FName FROM Learner_Info "
            "WHERE Grade = ? AND Class = ? AND Status = 'C'",
            grade, class_id,
        )
    learners = cursor.fetchall()

    results = []
    for learner in learners:
        cursor.execute(
            "SELECT SubjectId, Mark FROM ReportMarks WHERE LearnerID = ? AND ReportId = ?",
            learner.ID, report_id,
        )
        subject_marks = cursor.fetchall()
        if not subject_marks:
            continue  # no marks captured for this learner this term - skip

        subject_pcts = []
        for sm in subject_marks:
            pct = unrounded_subject_pct(cursor, learner.ID, sm.SubjectId)
            subject_pcts.append(pct if pct is not None else sm.Mark)

        avg = sum(subject_pcts) / len(subject_pcts)
        preferred_name = (learner.NickName or learner.FName or "").strip().upper()
        results.append({
            "learner_no": str(learner.AccessionNo).strip(),
            "surname": (learner.SName or "").strip().upper(),
            "name": preferred_name,
            "avg": avg,
            "pct": round_half_up(avg),
        })

    results.sort(key=lambda r: (-r["avg"], r["surname"], r["name"]))
    return results


def top_10_with_ties(learners):
    if len(learners) <= TOP_N:
        return learners
    cutoff_pct = learners[TOP_N - 1]["pct"]
    extended = list(learners[:TOP_N])
    for learner in learners[TOP_N:]:
        if learner["pct"] == cutoff_pct:
            extended.append(learner)
        else:
            break
    return extended


def style_row_like(ws, source_row, target_row):
    for col in COLUMNS.values():
        src_cell = ws[f"{col}{source_row}"]
        dst_cell = ws[f"{col}{target_row}"]
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
    ws.row_dimensions[target_row].height = DATA_ROW_HEIGHT


def fill_sheet(ws, title_c4, learners):
    from openpyxl.utils import get_column_letter

    ws["C4"] = title_c4
    ws["C6"] = TERM_LABEL

    last_row = FIRST_DATA_ROW + len(learners) - 1
    # extend styled rows beyond the template's 10 pre-built rows if there are ties
    for row in range(LAST_TEMPLATE_DATA_ROW + 1, last_row + 1):
        style_row_like(ws, LAST_TEMPLATE_DATA_ROW, row)
    # clear any unused template rows (fewer than 10 learners)
    for row in range(FIRST_DATA_ROW + len(learners), LAST_TEMPLATE_DATA_ROW + 1):
        for col in COLUMNS.values():
            ws[f"{col}{row}"] = None

    for i, learner in enumerate(learners):
        row = FIRST_DATA_ROW + i
        ws[f"{COLUMNS['nr']}{row}"] = i + 1
        ws[f"{COLUMNS['learner_no']}{row}"] = learner["learner_no"]
        ws[f"{COLUMNS['surname']}{row}"] = learner["surname"]
        ws[f"{COLUMNS['name']}{row}"] = learner["name"]
        ws[f"{COLUMNS['pct']}{row}"] = learner["pct"]

        for col_key in ("surname", "name"):
            cell = ws[f"{COLUMNS[col_key]}{row}"]
            text = str(cell.value or "")
            if len(text) > SHRINK_NAME_LEN:
                a = copy.copy(cell.alignment)
                a.shrink_to_fit = True
                cell.alignment = a

    # make sure the sheet prints on exactly one page regardless of row count
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:{get_column_letter(5)}{max(last_row, LAST_TEMPLATE_DATA_ROW)}"


def find_template_path():
    for candidate in (REPO_ROOT / "templates" / TEMPLATE_NAME, SCRIPT_DIR / TEMPLATE_NAME):
        if candidate.exists():
            return candidate
    print(f"[FAIL] {TEMPLATE_NAME} not found. Put it in the templates/ folder.")
    sys.exit(1)


def generate_workbook(cursor, template_path, output_dir):
    from openpyxl import load_workbook

    report_ids = get_report_id_map(cursor)
    classes = get_classes(cursor)

    template_wb = load_workbook(str(template_path))
    template_ws = template_wb[TEMPLATE_SHEET_NAME]

    summary = []  # (sheet name, learner count, extra tie rows)

    for grade, class_id, class_name in classes:
        report_id = report_ids[PHASE_BY_GRADE[grade]]
        final_list = top_10_with_ties(fetch_learners(cursor, grade, class_id, report_id))

        ws = template_wb.copy_worksheet(template_ws)
        ws.title = f"Grade {class_name}"
        fill_sheet(ws, f"GRADE  {class_name}", final_list)

        extra = len(final_list) - min(TOP_N, len(final_list))
        summary.append((ws.title, len(final_list), extra))

    for grade in range(7, 12):
        report_id = report_ids[PHASE_BY_GRADE[grade]]
        final_list = top_10_with_ties(fetch_learners(cursor, grade, None, report_id))

        ws = template_wb.copy_worksheet(template_ws)
        ws.title = f"Grade {grade}"
        fill_sheet(ws, f"GRADE  {grade}", final_list)

        extra = len(final_list) - min(TOP_N, len(final_list))
        summary.append((ws.title, len(final_list), extra))

    del template_wb[TEMPLATE_SHEET_NAME]
    if "Worksheet 1" in template_wb.sheetnames:
        del template_wb["Worksheet 1"]

    output_dir.mkdir(exist_ok=True)
    out_path = output_dir / f"Top 10 - Term {TERM_NUM} {YEAR}.xlsx"
    template_wb.save(str(out_path))
    return out_path, summary


def mass_print(workbook_path):
    try:
        import win32com.client
    except ImportError:
        print("[SKIP] pywin32 not installed — can't auto-print. Run: pip install pywin32")
        return

    print(f"\n[>>] Printing every sheet in {workbook_path.name} to default printer...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    printed = 0
    try:
        wb = excel.Workbooks.Open(str(workbook_path.resolve()))
        try:
            for ws in wb.Worksheets:
                try:
                    ws.PrintOut()
                    printed += 1
                except Exception as e:
                    print(f"  ERROR printing {ws.Name}: {e}")
        finally:
            wb.Close(SaveChanges=False)
        print(f"[OK] Sent {printed}/{wb.Worksheets.Count} sheet(s) to printer")
    finally:
        excel.Quit()


def main():
    args = sys.argv[1:]

    if "--discover" in args:
        conn = connect()
        discover_schema(conn)
        conn.close()
        return

    no_print = "--no-print" in args or not AUTO_PRINT

    print("=" * 60)
    print("  SASAMS TOP 10 PIPELINE")
    print(f"  Year: {YEAR}  |  {TERM_LABEL}")
    print("=" * 60)

    template_path = find_template_path()

    print("\n[STEP 1] Querying SASAMS database...")
    conn = connect()
    cursor = conn.cursor()
    term_start, term_end = get_term_dates(cursor)
    date_warnings = check_task_dates(cursor, term_start, term_end)

    print(f"\n[STEP 2] Generating workbook...")
    output_dir = SCRIPT_DIR / OUTPUT_FOLDER
    out_path, summary = generate_workbook(cursor, template_path, output_dir)
    conn.close()

    print(f"[OK] Saved: {out_path}")
    print(f"\n{'Sheet':<14}{'Learners':>10}{'Extra tie rows':>16}")
    for name, count, extra in summary:
        tie_note = f"+{extra}" if extra else ""
        print(f"{name:<14}{count:>10}{tie_note:>16}")

    ties = [(n, e) for n, _, e in summary if e]
    if ties:
        print("\nSheets with ties beyond rank 10:")
        for n, e in ties:
            print(f"  {n}: {e} extra learner(s) tied with 10th place")
    else:
        print("\nNo ties beyond rank 10 on any sheet.")

    if date_warnings:
        print(
            f"\nWARNING: {len(date_warnings)} task(s) are tagged {TERM_LABEL} "
            f"but their own date is more than {TASK_DATE_WARNING_BUFFER_DAYS} days "
            f"outside the term's official {term_start:%Y-%m-%d} to {term_end:%Y-%m-%d} "
            "window. They ARE still included in the calculation (SubHeading is "
            "trusted as authoritative) - this is just a heads-up in case one is "
            "mistagged or is a year-end rollup task rather than a real term task."
        )
        print(f"{'Subject':<40}{'Task':<35}{'Date':<12}{'Weight':>7}")
        for w in date_warnings:
            print(f"{(w.Name or ''):<40}{w.Description:<35}{w.DateAdded:%Y-%m-%d}  {w.Weighting:>6.1f}")
    else:
        print(f"\nNo {TERM_LABEL} tasks dated outside the term window - nothing to flag.")

    if no_print:
        print(f"\n[STEP 3] Printing skipped (--no-print)")
    else:
        print(f"\n[STEP 3] Printing...")
        mass_print(out_path)

    print("\n" + "=" * 60)
    print(f"  DONE — {len(summary)} sheets in {out_path.name}")
    if not no_print:
        print(f"  Sent to printer")
    print(f"  File in: {output_dir}")
    print("=" * 60)


if __name__ == "__main__":
    main()
