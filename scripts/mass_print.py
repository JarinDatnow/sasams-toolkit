"""
MASS PRINT SPREADSHEETS
========================
Prints the "Worksheet" sheet from every .xlsx in this folder to your default printer.

1. Put this script in the folder with your .xlsx files
2. Run:  python mass_print.py
3. Pages go brrr

Requires: pip install pywin32
(Excel must be installed on the machine)
"""

import sys
from pathlib import Path

try:
    import win32com.client
except ImportError:
    print("pywin32 not installed. Run: pip install pywin32")
    sys.exit(1)


def main():
    folder = Path('.').resolve()
    files = sorted([f for f in folder.glob('*.xlsx')])

    if not files:
        print("No .xlsx files found in this folder.")
        sys.exit(1)

    print(f"Found {len(files)} file(s) to print")
    print(f"Printing 'Worksheet' sheet from each to default printer...")
    print("-" * 50)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    printed = 0
    try:
        for f in files:
            try:
                wb = excel.Workbooks.Open(str(f))
                try:
                    ws = wb.Worksheets("Worksheet")
                    ws.PrintOut()
                    printed += 1
                    print(f"  Printed: {f.name}")
                except Exception:
                    print(f"  SKIPPED (no 'Worksheet' sheet): {f.name}")
                wb.Close(SaveChanges=False)
            except Exception as e:
                print(f"  ERROR with {f.name}: {e}")

        print("-" * 50)
        print(f"Done! Sent {printed}/{len(files)} to printer.")
    finally:
        excel.Quit()


if __name__ == '__main__':
    main()
