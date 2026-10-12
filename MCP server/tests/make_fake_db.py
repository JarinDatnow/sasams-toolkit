"""
Builds a small fake SASAMS database (SQLite) with made-up learners, so the
MCP server can be tested anywhere without real school data or the Access
driver. Same table and column names as SASAMS for the parts the server uses.
"""

import random
import sqlite3
import sys
from pathlib import Path

SUBJECTS = ["Mathematics", "English Home Language", "Afrikaans First Additional Language",
            "Life Orientation", "Physical Sciences", "Life Sciences"]


def build(path):
    path = Path(path)
    if path.exists():
        path.unlink()
    rng = random.Random(42)
    db = sqlite3.connect(path)
    db.executescript("""
        CREATE TABLE Classes (ClassId INTEGER, Grade INTEGER, ClassName TEXT);
        CREATE TABLE Learner_Info (ID INTEGER, AccessionNo TEXT, SName TEXT, FName TEXT, NickName TEXT,
                                   IDNo TEXT, Grade INTEGER, Class INTEGER, Status TEXT, Gender TEXT);
        CREATE TABLE Subjects (Id INTEGER, Name TEXT);
        CREATE TABLE ReportCycles (CycleId INTEGER, Name TEXT, Phase TEXT, Datayear TEXT, Term INTEGER);
        CREATE TABLE ReportMarks (Id INTEGER, LearnerID INTEGER, SubjectId INTEGER, Mark INTEGER,
                                  ReportId INTEGER, Datayear TEXT, Comment1 TEXT);
        CREATE TABLE Ana2012EvaluationLevels (Id INTEGER, EvalVer INTEGER, GradeFrom INTEGER, GradeTo INTEGER,
                                              Code INTEGER, MarkFrom REAL, MarkTo REAL);
        CREATE TABLE SchoolTerms (id INTEGER, StartDate TEXT, EndDate TEXT, CurrentYear TEXT, Term INTEGER);
        CREATE TABLE Absentees (Id INTEGER, Learnerid INTEGER, DateAbsent TEXT, Grade INTEGER, Class INTEGER,
                                DataYear TEXT, ReasonOther TEXT);
    """)
    db.executemany("INSERT INTO Classes VALUES (?,?,?)", [(1, 10, "10A"), (2, 10, "10B"), (3, 9, "9A")])
    for code, lo in zip(range(1, 8), [0, 30, 40, 50, 60, 70, 80]):
        db.execute("INSERT INTO Ana2012EvaluationLevels VALUES (?,3,7,12,?,?,?)", (code, code, lo, lo + 9))
    for i, name in enumerate(SUBJECTS, start=1):
        db.execute("INSERT INTO Subjects VALUES (?,?)", (i, f"{name} (Gr 10)"))
    cycles = [(11, "FET Term 1", "FET", 1), (12, "FET Term 2", "FET", 2), (13, "FET Term 3", "FET", 3),
              (21, "Senior Term 1", "Senior", 1)]
    for cid, name, phase, term in cycles:
        db.execute("INSERT INTO ReportCycles VALUES (?,?,?,'2026',?)", (cid, name, phase, term))
    db.executemany("INSERT INTO SchoolTerms VALUES (?,?,?,'2026',?)", [
        (1, "2026-01-14", "2026-03-27", 1), (2, "2026-04-08", "2026-06-26", 2), (3, "2026-07-21", "2026-10-02", 3)])

    mark_id = absent_id = 0
    for lid in range(1, 61):
        cls = 1 if lid <= 30 else 2
        status = "L" if lid == 60 else "C"
        db.execute("INSERT INTO Learner_Info VALUES (?,?,?,?,?,?,?,?,?,?)",
                   (lid, f"A{1000 + lid}", f"TESTSURNAME{lid}", f"TESTNAME{lid}", None,
                    f"0801015{lid:06d}", 10, cls, status, "M" if lid % 2 else "F"))
        ability = rng.gauss(55, 15)
        for sid in range(1, len(SUBJECTS) + 1):
            prev = None
            for cid, term in ((11, 1), (12, 2), (13, 3)):
                mark = ability + rng.gauss(0, 8) + (sid == 1) * -8
                if lid in (5, 6, 7) and term == 3:
                    mark -= 25
                mark = max(0, min(100, round(mark)))
                mark_id += 1
                db.execute("INSERT INTO ReportMarks VALUES (?,?,?,?,?,'2026',?)",
                           (mark_id, lid, sid, mark, cid, "private teacher note"))
        for _ in range(rng.choice([0, 0, 1, 2, 8])):
            absent_id += 1
            day = rng.randint(21, 31)
            db.execute("INSERT INTO Absentees VALUES (?,?,?,10,?,'2026','sick')",
                       (absent_id, lid, f"2026-07-{day:02d}", cls))
    for lid in range(61, 71):
        db.execute("INSERT INTO Learner_Info VALUES (?,?,?,?,?,?,9,3,'C','F')",
                   (lid, f"A{1000 + lid}", f"TESTSURNAME{lid}", f"TESTNAME{lid}", None, "x"))
    db.commit()
    db.close()


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "fake_sasams.db")
