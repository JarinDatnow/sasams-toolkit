"""
SASAMS QUERY RUNNER
====================
Paste a SQL query, get a CSV. That's it.

1. Make sure config.py is set up (copy from config.example.py)
2. Paste your query into the QUERY variable below
3. Run:  python sasams_query.py
4. CSV appears in the output folder
"""

import pyodbc
import csv
import os
import sys
from datetime import datetime

try:
    from config import DB_PATH, DB_PASSWORD
except ImportError:
    print("[FAIL] config.py not found.")
    print("       Copy config.example.py to config.py and fill in your values.")
    sys.exit(1)

OUTPUT_FOLDER = os.path.expanduser("~\\Desktop")

# ── PASTE YOUR QUERY HERE ──────────────────────────────────────────────────
QUERY = """
SELECT 
    P.Grade, 
    C.ClassName,
    L.SName, 
    L.FName, 
    P.LearnerAverage, 
    P.LearnerScore, 
    P.CodeSelected AS PromotionStatus,
    P.DataYear
FROM (Learner_Info L 
INNER JOIN LearnerPromotion P ON L.ID = P.LearnerId)
LEFT JOIN Classes C ON L.Class = C.ClassId
WHERE P.DataYear = '2026' 
  AND P.LearnerAverage IS NOT NULL
  AND P.ReportId = (
    SELECT MAX(P2.ReportId) 
    FROM LearnerPromotion P2 
    WHERE P2.LearnerId = P.LearnerId 
    AND P2.DataYear = '2026'
  )
ORDER BY P.Grade, C.ClassName, P.LearnerAverage DESC
"""
# ── END QUERY ──────────────────────────────────────────────────────────────

OUTPUT_NAME = ""


def connect():
    driver = "{Microsoft Access Driver (*.mdb, *.accdb)}"
    conn_str = f"DRIVER={driver};DBQ={os.path.abspath(DB_PATH)};PWD={DB_PASSWORD};"
    try:
        conn = pyodbc.connect(conn_str)
        print("[OK] Connected to database")
        return conn
    except (pyodbc.Error, UnicodeDecodeError) as e:
        print(f"[FAIL] Could not connect: {e}")
        sys.exit(1)


def run_query(conn, sql):
    cursor = conn.cursor()
    try:
        cursor.execute(sql)
        columns = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
        print(f"[OK] Query returned {len(rows)} rows, {len(columns)} columns")
        return columns, rows
    except (pyodbc.Error, UnicodeDecodeError) as e:
        print(f"[FAIL] Query error: {e}")
        sys.exit(1)


def save_csv(columns, rows, output_path):
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(columns)
        for row in rows:
            writer.writerow([str(v) if v is not None else "" for v in row])
    print(f"[OK] Saved to: {output_path}")


if __name__ == "__main__":
    sql = QUERY.strip()
    if not sql:
        print("[FAIL] No query found. Paste your SQL into the QUERY variable.")
        sys.exit(1)

    print(f"[>>] Running query...")
    print(f"     {sql[:120]}{'...' if len(sql) > 120 else ''}")
    print()

    conn = connect()
    columns, rows = run_query(conn, sql)
    conn.close()

    if OUTPUT_NAME:
        filename = OUTPUT_NAME if OUTPUT_NAME.endswith(".csv") else OUTPUT_NAME + ".csv"
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"SASAMS_result_{timestamp}.csv"

    output_path = os.path.join(OUTPUT_FOLDER, filename)
    save_csv(columns, rows, output_path)

    print()
    print(f"Done. Open the CSV: {output_path}")
