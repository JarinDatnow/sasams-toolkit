# SASAMS Toolkit

**Automation scripts for South African schools running SASAMS and D6.**

If you've ever spent hours manually exporting top achiever reports, copying data into templates, and printing 30+ spreadsheets one by one — this toolkit does all of that in seconds.

---

## What is this?

A collection of Python scripts that automate common school admin tasks by querying the SASAMS database directly:

- **`top10_pipeline.py`** — The main event. Queries your SASAMS database, generates formatted Top 10 achiever spreadsheets per class/grade, and mass-prints them. One command, done.
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
   git clone https://github.com/YOUR_USERNAME/sasams-toolkit.git
   cd sasams-toolkit
   ```

2. **Create your config**
   ```bash
   copy config.example.py config.py
   ```
   Then edit `config.py` with your database path, password, current year, and term.

3. **Add your template**  
   Place your `TEMPLATE__top_10.xlsx` in the `templates/` folder. This is the formatted spreadsheet template that gets populated with data.

4. **Run it**
   ```bash
   cd scripts
   python top10_pipeline.py
   ```

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

The pipeline queries the SASAMS database for learner averages, groups them by class (Grades R–6) or by grade (Grades 7–12), takes the top 10 per group, populates an Excel template with the results, and optionally sends them all to your default printer.

### Grouping logic
- **Grades R–6**: Top 10 per **class** (RA, RB, 1A, 1B, 2A, etc.) using the `ClassName` field
- **Grades 7–12**: Top 10 per **grade** (all classes combined)

## Project structure

```
sasams-toolkit/
├── config.example.py      # Template config — copy to config.py
├── .gitignore              # Keeps credentials and data files out of git
├── README.md
├── scripts/
│   ├── top10_pipeline.py   # Full pipeline: query → xlsx → print
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

- Edit the SQL query in `top10_pipeline.py` to pull different data
- Modify the template to change the spreadsheet layout
- Adjust `PER_GRADE_GRADES` in the script if your school groups grades differently
- Use `sasams_query.py` with the `--discover` flag to explore what other data is available

## Contributing

Found a bug? Got a better query? Open an issue or PR. This was built for one school but should work for any SA school running SASAMS.

## Disclaimer

This toolkit is provided as-is for educational and administrative purposes. It reads data from your local SASAMS database — it does not modify any records. Ensure you comply with your school's data policies and POPIA requirements when handling learner information.

## License

MIT — do whatever you want with it.
