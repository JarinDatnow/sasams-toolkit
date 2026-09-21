# SASAMS Toolkit

**Automation scripts for South African schools running SASAMS and D6.**

If you've ever spent hours manually exporting top achiever reports, copying data into templates, and printing 30+ spreadsheets one by one — this toolkit does all of that in seconds.
Built primarily as a portable, version-controlled backup of my own end-of-term workflow — if my machine dies or I change schools, setup is a git clone away instead of a manual rebuild. Shared publicly in case it helps other SA schools running SASAMS.
---

## What is this?

A collection of Python scripts that automate common school admin tasks by querying the SASAMS database directly:

- **`top10_pipeline.py`** — The main event. Queries your SASAMS database, generates one formatted Top 10 workbook (a sheet per class/grade, correctly ordered, ties handled), and prints it. One command, done. Verified match against 306 real D6-exported Top Achiever rows across two different terms — see "How the percentage is calculated" below before touching the query.
- **`sasams_query.py`** — Generic query runner. Paste any SQL query, get a CSV on your desktop. Useful for custom reports.
- **`make_top10.py`** — Standalone version that works with D6 exports (if you prefer exporting from D6 rather than querying the database directly).
- **`mass_print.py`** — Prints every `.xlsx` in a folder to your default printer. Works with any spreadsheets, not just Top 10s.

## What is SASAMS?

SASAMS (South African School Administration and Management System) is the Department of Basic Education's system for managing learner data, marks, attendance, and promotion across South African schools. Schools typically interact with it through an Access `.mdb` database file that syncs with the provincial system.

## Requirements

- **Windows** (required for the Access database driver and Excel printing)
- **Python 3.8+**
- **Microsoft Excel** (for the mass print feature)
- **Microsoft Access Database Engine** (usually already installed if you use SASAMS)
- Read access to your school's SASAMS `.mdb` database file

### Python packages

```bash
pip install pyodbc openpyxl pywin32
```

## Setup

1. **Clone this repo**
   ```bash
   git clone https://github.com/JarinDatnow/sasams-toolkit.git
   cd sasams-toolkit
   ```

2. **Create your config**
   ```bash
   copy config.example.py config.py
   ```
   Then edit `config.py` with your database path, password, current academic year (`DATA_YEAR`), and current term (`TERM` — just the number, e.g. `3`).

3. **Add your template**  
   Place your `TEMPLATE__top_10.xlsx` in the `templates/` folder. This is the formatted spreadsheet template that gets populated with data. It must have `Nr`, `LEARNER #`, `SURNAME`, `NAME` and `%` columns starting at row 13 (columns A–E) — see the layout constants at the top of `top10_pipeline.py` if your template differs.

4. **Run it**
   ```bash
   python scripts/top10_pipeline.py
   ```
   (works from either the repo root or from inside `scripts/` — `config.py` is found either way)

## Usage

### Full pipeline (query → generate → print)
```bash
python scripts/top10_pipeline.py
```

### Generate only (no printing)
```bash
python scripts/top10_pipeline.py --no-print
```

### Explore your database structure
```bash
python scripts/top10_pipeline.py --discover
```
This dumps all tables and columns so you can write custom queries.

### Run a custom query
Edit the `QUERY` variable in `sasams_query.py`, then:
```bash
python scripts/sasams_query.py
```

### Process D6 exports instead
If you export top achiever reports from D6:
1. Put `TEMPLATE__top_10.xlsx` and all your export `.xlsx` files in one folder
2. Copy `make_top10.py` into that folder
3. Run:
   ```bash
   python make_top10.py
   ```

### Mass print any folder of spreadsheets
Copy `mass_print.py` into a folder of `.xlsx` files and run:
```bash
python mass_print.py
```

## How it works

The pipeline queries the SASAMS database for each learner's term percentage, groups them by class (Grades R–6) or by grade (Grades 7–11), takes the top 10 per group (plus anyone tied with 10th place), writes everything into one Excel workbook — a sheet per group, in the order `Grade RA, RB, 1A…6C, 7…11` — and optionally sends every sheet to your default printer.

### Grouping logic
- **Grades R–6**: Top 10 per **class** (RA, RB, 1A, 1B, 2A, etc.) using the `Classes` table
- **Grades 7–11**: Top 10 per **grade** (all classes combined). Grade 12 is intentionally excluded (final-year learners typically get a different kind of report, not a mid-year Top 10).

### Ties
If the 10th-place learner shares their percentage with anyone below them, those learners are added as extra rows (11, 12, …) styled identically to the rest, instead of being cut off arbitrarily.

### How the percentage is calculated (read this before changing the query)

It's tempting to average `ReportMarks.Mark` (SASAMS's own already-rounded, whole-number mark per subject) or read `LearnerPromotion.LearnerAverage` (a precomputed promotion average) directly. **Both were tried and both are wrong** — they can land a full percentage point off what D6's own "Top Achievers" report actually prints, for two reasons:

1. Each subject's mark is rounded to a whole number for display *before* it's stored in `ReportMarks` — but D6 averages the **unrounded** per-subject percentage, not the rounded one. That unrounded value lives one level deeper, in the continuous-assessment task tables:
   - `LearnerCass` — one row per task (test, assignment, exam...) per learner per subject, with `Mark` out of `Criterionscore`
   - `SubjectCriteria` — defines each task's `Weighting` toward the subject total for the term

   Subject % = `sum(Mark/Criterionscore * Weighting) / sum(Weighting) * 100`, and the learner's term % is round-half-up of the *average of those unrounded subject percentages*.

2. Tasks are matched to a specific term via `SubjectCriteria.SubHeading` (an explicit `'Term1'`/`'Term2'`/`'Term3'`/`'Term4'` label) — **not** by date. A task's own `DateAdded` is close to its term but not reliable enough to filter on by itself (e.g. a placeholder "SBA Year Mark" rollup task can carry a nonsense date far outside the term). The script still does a sanity pass: it flags (without excluding) any task tagged for the target term whose date falls more than 14 days outside that term's official window, so you can eyeball anything that looks mistagged.

This was verified by reproducing 306 real learners' percentages from two D6-exported "Top Achievers" terms with **zero mismatches** once both fixes were in place. If you're pulling a different metric than Top 10 percentages, this detail may not matter to you — but if numbers ever look "off by one" versus what D6 prints, this is almost certainly why.

## Project structure

```
sasams-toolkit/
├── config.example.py      # Template config — copy to config.py
├── .gitignore              # Keeps credentials and data files out of git
├── README.md
├── scripts/
│   ├── top10_pipeline.py   # Full pipeline: query → one xlsx workbook → print
│   ├── sasams_query.py     # Generic query runner → CSV
│   ├── make_top10.py       # D6 export version
│   └── mass_print.py       # Mass print .xlsx files
└── templates/
    └── TEMPLATE__top_10.xlsx  # Your formatted template (add your own)
```

## Database access

These scripts require read access to your school's SASAMS `.mdb` database file. The database is typically exported from D6/SASAMS with a password. You will need to obtain this password through your school's IT administrator or SASAMS coordinator.

**Important**: Your `config.py` file contains database credentials and is excluded from git via `.gitignore`. Never commit credentials to a public repository.

## Customisation

- Edit the SQL queries in `top10_pipeline.py` (`fetch_learners`, `unrounded_subject_pct`) to pull different data
- Modify the template to change the spreadsheet layout — update the layout constants (`FIRST_DATA_ROW`, `COLUMNS`, etc.) at the top of `top10_pipeline.py` to match
- Adjust `PHASE_BY_GRADE` and the `range(7, 12)` grouping in `top10_pipeline.py` if your school phases or group grades differently
- Use `python scripts/top10_pipeline.py --discover` (or `sasams_query.py`) to explore what other data is available — there's also a full pre-dumped schema in `SASAMS_map.json` at the repo root

## Contributing

Found a bug? Got a better query? Open an issue or PR. This was built for one school but should work for any SA school running SASAMS.

## Disclaimer

This toolkit is provided as-is for educational and administrative purposes. It reads data from your local SASAMS database — it does not modify any records. Ensure you comply with your school's data policies and POPIA requirements when handling learner information.

## License

MIT — do whatever you want with it.
