"""
End-to-end test against a fake SASAMS database: starts the real MCP server
over stdio, lists its tools, calls each data tool, and checks that no learner
name ever comes back to the client.

Run from the "MCP server" folder:  python tests/test_server.py
"""

import asyncio
import json
import os
import sys
import tempfile
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

HERE = Path(__file__).resolve().parent
SERVER = HERE.parent / "sasams_mcp.py"
sys.path.insert(0, str(HERE))
from make_fake_db import build  # noqa: E402


def text_of(result):
    if getattr(result, "structuredContent", None):
        return json.dumps(result.structuredContent)
    return "".join(getattr(c, "text", "") for c in result.content)


async def main():
    tmp = Path(tempfile.mkdtemp())
    db = tmp / "fake_sasams.db"
    build(db)
    env = {**os.environ, "SASAMS_SQLITE_PATH": str(db), "SASAMS_YEAR": "2026", "SASAMS_OUTPUT_DIR": str(tmp / "out")}
    params = StdioServerParameters(command=sys.executable, args=[str(SERVER)], env=env, cwd=str(SERVER.parent))

    failures = []

    def check(label, condition, detail=""):
        print(("PASS " if condition else "FAIL ") + label + ("" if condition else f"  {detail}"))
        if not condition:
            failures.append(label)

    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = {t.name for t in (await session.list_tools()).tools}
            expected = {"school_overview", "list_report_cycles", "subject_results", "compare_terms",
                        "at_risk_learners", "learner_marks", "attendance_summary", "export_learner_list",
                        "generate_top10", "generate_annexure_k", "search_schema", "describe_table", "run_select"}
            check("all tools listed", expected <= tools, sorted(expected - tools))

            async def call(name, **args):
                res = await session.call_tool(name, args)
                out = text_of(res)
                check(f"{name} never returns names", "TESTSURNAME" not in out and "TESTNAME" not in out, out[:200])
                return res, out

            _, out = await call("school_overview")
            data = json.loads(out)
            g10 = next(g for g in data["grades"] if g["grade"] == 10)
            check("overview counts current learners only", g10["learners"] == 59, g10)

            _, out = await call("subject_results", grade="10", term=3)
            data = json.loads(out)
            check("subject_results covers 6 subjects", len(data["subjects"]) == 6, data)
            check("levels add up", all(sum(s["levels_1_to_7"]) == s["learners"] for s in data["subjects"]))
            check("weakest subject first is maths", data["subjects"][0]["subject"].startswith("Mathematics"), data["subjects"][0])

            _, out = await call("subject_results", grade="10", term=3, class_name="10b", subject_contains="math")
            data = json.loads(out)
            check("class + subject filter", len(data["subjects"]) == 1 and data["learners_in_selection"] == 29, data)

            _, out = await call("compare_terms", grade="10", term_from=2, term_to=3, drop_threshold=15)
            data = json.loads(out)
            dropped = {f["learner_id"] for f in data["flagged"]}
            check("compare_terms flags the planted drops", {5, 6, 7} <= dropped, sorted(dropped)[:20])

            _, out = await call("at_risk_learners", grade="10", term=3, below_mark=40, min_subjects=3)
            data = json.loads(out)
            check("at_risk returns IDs", data["at_risk_total"] >= 1 and all("learner_id" in r for r in data["learners"]), data)

            _, out = await call("learner_marks", learner_id=5)
            data = json.loads(out)
            check("learner_marks has 3 terms", all({"term_1", "term_2", "term_3"} <= set(m) for m in data["marks"]), data)

            _, out = await call("attendance_summary", grade="10", term=3, flag_days=5)
            data = json.loads(out)
            check("attendance per class", len(data["by_class"]) == 2, data)

            res, out = await call("export_learner_list", learner_ids=[5, 6, 7], title="Term 3 maths drop")
            data = json.loads(out)
            check("export writes a file", Path(data["file"]).exists() and data["learners_written"] == 3, data)
            from openpyxl import load_workbook
            ws = load_workbook(data["file"]).active
            check("export file holds the names locally", ws["B5"].value == "TESTSURNAME5", ws["B5"].value)

            _, out = await call("run_select", sql="SELECT SubjectId, AVG(Mark) AS avg_mark FROM ReportMarks WHERE ReportId = 13 GROUP BY SubjectId")
            data = json.loads(out)
            check("run_select works", data.get("row_count") == 6, data)

            for bad in ["SELECT SName FROM Learner_Info", "SELECT * FROM ReportMarks",
                        "DELETE FROM ReportMarks", "SELECT Comment1 FROM ReportMarks",
                        "SELECT SName || FName AS x FROM Learner_Info", "SELECT Name FROM Educators",
                        "SELECT DLookUp('SName', 'Learner_Info', 'ID = 5') AS x FROM Classes",
                        "SELECT x FROM (SELECT [IDNo] AS x FROM Learner_Info)",
                        "SELECT ID FROM Learner_Info WHERE SName LIKE 'A%'"]:
                _, out = await call("run_select", sql=bad)
                check(f"run_select rejects: {bad}", "error" in json.loads(out), out[:150])

            _, out = await call("run_select", sql="SELECT ID, Class FROM Learner_Info WHERE Grade = 10 AND Status = 'C' ORDER BY ID", max_rows=5)
            data = json.loads(out)
            check("run_select row cap + truncation flag", data["row_count"] == 5 and data["truncated"], data)

            _, out = await call("describe_table", table="learner_info")
            data = json.loads(out)
            check("describe hides personal columns", "SName" not in data["columns"] and "Grade" in data["columns"], data)

    print()
    print("ALL PASSED" if not failures else f"{len(failures)} FAILED: {failures}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
