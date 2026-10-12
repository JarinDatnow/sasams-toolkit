# SASAMS MCP server

Ask questions about your SASAMS data in plain words, in Claude Code, instead of
writing a new script each time:

- "How did Grade 10 do in Term 3, weakest subject first?"
- "Which Grade 10 learners dropped 15+ points in maths from Term 2 to Term 3? Make 10B's teacher a list."
- "How many Grade 9 learners are failing 3 or more subjects?"
- "Which class had the most absences this term?"
- "Generate the Top 10 and Annexure K for Term 3."

Claude runs the queries itself, sees the results or the errors, and fixes its
own SQL, so there's no copy-pasting queries back and forth.

## Safety rules built in

| Rule | How |
| --- | --- |
| Never changes SASAMS | Opens the database read-only, and only single `SELECT` statements are allowed. Point it at a **copy** of the `.mdb`, never the live file. |
| No names leave the PC | Learners are identified by SASAMS's internal row number (`Learner_Info.ID`). Names, ID numbers, contact, medical, grant and comment columns are blocked in every query, and any result column that looks personal is blanked. |
| Teachers still get names | `export_learner_list` writes an Excel list with names to `MCP server/Output/` on this PC. Only the file path goes back to Claude. |
| Only academic data | Free-form queries can only use the marks, subjects, classes, report cycles, assessment and attendance tables. Staff, parent and finance tables are off-limits. |

Marks, averages and counts **are** sent to Anthropic when Claude reads them.
Get the principal's go-ahead before using this with real data (POPIA).

## Tools

| Tool | What it answers |
| --- | --- |
| `school_overview` | Grades, classes and learner counts |
| `list_report_cycles` | Which terms have marks captured |
| `subject_results` | Per-subject average, median, CAPS levels 1–7, share below level 2 and at level 4+ (filter by class or subject) |
| `compare_terms` | Mark changes between two terms per subject, plus learners who dropped by a threshold |
| `at_risk_learners` | Learners with N or more subjects below a mark |
| `learner_marks` | One learner's marks across the year (by ID) |
| `attendance_summary` | Days absent per class for a year or term, plus learners over a threshold |
| `export_learner_list` | Named Excel list for a teacher, saved locally |
| `generate_top10`, `generate_annexure_k` | Run the existing generators for a term |
| `search_schema`, `describe_table` | Find tables and columns (personal columns hidden) |
| `run_select` | Any other read-only question, with the safety rules above |

## Setup (the Windows PC with the Access driver)

1. Install the Python packages (64-bit Python 3.10 or newer, same PC as the other generators):
   ```
   pip install -r "MCP server/requirements.txt"
   ```
2. Put the database password in `MCP server/db_secret.py` (copy `db_secret.example.py`), or reuse the
   `db_secret.py` already in `Top 10 automation`; the server finds either.
3. Copy the term's `.mdb` somewhere outside the repo, for example
   `C:\SASAMS copy\school.mdb`.
4. Register the server with Claude Code (one line, adjust both paths; `-s user` makes it
   available whichever folder you start Claude in):
   ```
   claude mcp add sasams -s user --env SASAMS_DB_PATH="C:\SASAMS copy\school.mdb" -- python "C:\path\to\sasams-toolkit\MCP server\sasams_mcp.py"
   ```
   `claude mcp get sasams` should then show `Connected`.
5. Start `claude` and ask away. Each new term: replace the copy of the `.mdb` with the new one.

Optional settings (add more `--env` flags): `SASAMS_YEAR` (default year, otherwise the
current year) and `SASAMS_OUTPUT_DIR` (where exported lists go).

## Testing without real data

`tests/make_fake_db.py` builds a small fake SASAMS database with made-up learners, and
`tests/test_server.py` starts the server, calls every data tool, and checks that no
names come back and that unsafe queries are refused:

```
cd "MCP server"
python tests/test_server.py
```

The fake database is SQLite, so the test runs on any OS. The real server talks to
Access through the same queries the Top 10 and Annexure K generators use.
