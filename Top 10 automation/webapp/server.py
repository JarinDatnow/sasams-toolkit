"""
Top 10 Club - local web UI for generate_top10.py AND generate_annexure_k.py

Runs a small local-only web server (no network access needed, nothing is
sent anywhere) that serves the themed page in webapp/static and calls
straight into generate_top10.generate() or generate_annexure_k.generate()
when you click GENERATE, depending on which mode is selected on the page.

Launch it by double-clicking "Launch Top 10 Club.vbs" in the project
folder (that starts this file with pythonw and opens your browser to it),
or run it directly with:  python webapp/server.py
"""

import contextlib
import io
import json
import os
import subprocess
import sys
import threading
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
ANNEXURE_ROOT = os.path.join(os.path.dirname(PROJECT_ROOT), "Annexure K automation")

sys.path.insert(0, PROJECT_ROOT)
import generate_top10  # noqa: E402  (must come after sys.path insert)

sys.path.insert(0, ANNEXURE_ROOT)
import generate_annexure_k  # noqa: E402  (must come after sys.path insert)

HOST = "127.0.0.1"
PORT = 5757

STATIC_FILES = {
    "/": ("index.html", "text/html; charset=utf-8"),
    "/index.html": ("index.html", "text/html; charset=utf-8"),
    "/style.css": ("style.css", "text/css; charset=utf-8"),
    "/app.js": ("app.js", "application/javascript; charset=utf-8"),
}

# Only one generation / browse dialog at a time - keeps things simple and
# avoids two Tk dialogs fighting each other.
_lock = threading.Lock()


def browse_for_file(kind):
    """Show a native "open file" dialog on top of everything else and
    return the chosen path, or "" if the user cancelled. `kind` picks the
    dialog's title/filetypes/starting folder."""
    import tkinter as tk
    from tkinter import filedialog

    if kind == "template":
        title = "Select the Annexure K template workbook"
        filetypes = [("Excel Workbook", "*.xlsx"), ("All files", "*.*")]
        initialdir = ANNEXURE_ROOT if os.path.isdir(ANNEXURE_ROOT) else PROJECT_ROOT
    else:
        title = "Select the Access database (.mdb / .accdb)"
        filetypes = [("Access Database", "*.mdb *.accdb"), ("All files", "*.*")]
        db_dir = os.path.dirname(generate_top10.DB_PATH)
        initialdir = db_dir if os.path.isdir(db_dir) else PROJECT_ROOT

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        path = filedialog.askopenfilename(
            title=title, filetypes=filetypes, initialdir=initialdir,
        )
    finally:
        root.destroy()
    return path or ""


def browse_for_folder(initial):
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        path = filedialog.askdirectory(
            title="Select the output folder",
            initialdir=initial if os.path.isdir(initial) else PROJECT_ROOT,
        )
    finally:
        root.destroy()
    return path or ""


class Handler(BaseHTTPRequestHandler):
    server_version = "TopTenClub/1.0"

    def log_message(self, fmt, *args):
        pass  # keep the console quiet

    def _send_json(self, obj, status=200):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        length = int(self.headers.get("Content-Length", 0))
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    # ---- routing -----------------------------------------------------
    def do_GET(self):
        if self.path.startswith("/api/defaults"):
            mode = "topten"
            if "?" in self.path:
                query = self.path.split("?", 1)[1]
                params = dict(p.split("=", 1) for p in query.split("&") if "=" in p)
                mode = params.get("mode", "topten")

            if mode == "annexurek":
                self._send_json({
                    "db_path": generate_annexure_k.DB_PATH,
                    "term": generate_annexure_k.TERM,
                    "year": generate_annexure_k.DATA_YEAR,
                    "groupings": generate_annexure_k.GROUPINGS,
                    "templates": generate_annexure_k.TEMPLATES,
                    "output_folder": generate_annexure_k.OUTPUT_FOLDER,
                })
            else:
                self._send_json({
                    "db_path": generate_top10.DB_PATH,
                    "term": generate_top10.TERM,
                    "year": generate_top10.YEAR,
                    "output_folder": generate_top10.OUTPUT_FOLDER,
                })
            return

        entry = STATIC_FILES.get(self.path.split("?")[0])
        if entry:
            filename, content_type = entry
            file_path = os.path.join(STATIC_DIR, filename)
            try:
                with open(file_path, "rb") as f:
                    body = f.read()
            except FileNotFoundError:
                self.send_error(404)
                return
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        self.send_error(404)

    def do_POST(self):
        if self.path == "/api/browse-file":
            data = self._read_json()
            with _lock:
                path = browse_for_file(data.get("kind", "db"))
            self._send_json({"path": path})
            return

        if self.path == "/api/browse-folder":
            data = self._read_json()
            with _lock:
                path = browse_for_folder(data.get("initial", PROJECT_ROOT))
            self._send_json({"path": path})
            return

        if self.path == "/api/open-folder":
            data = self._read_json()
            folder = data.get("path", "")
            if os.path.isdir(folder):
                subprocess.Popen(["explorer", folder])
                self._send_json({"ok": True})
            else:
                self._send_json({"ok": False, "error": "Folder not found."}, 400)
            return

        if self.path == "/api/generate":
            data = self._read_json()
            if data.get("mode") == "annexurek":
                self._handle_generate_annexurek(data)
            else:
                self._handle_generate_topten(data)
            return

        self.send_error(404)

    def _handle_generate_topten(self, data):
        db_path = (data.get("db_path") or "").strip()
        db_password = (data.get("db_password") or "").strip() or None
        term = data.get("term")
        year = data.get("year")
        output_folder = (data.get("output_folder") or "").strip() or None

        if not db_path:
            self._send_json({"ok": False, "log": "", "error": "No database selected."}, 400)
            return
        if not os.path.isfile(db_path):
            self._send_json(
                {"ok": False, "log": "", "error": f"Database file not found:\n{db_path}"},
                400,
            )
            return
        try:
            term = int(term)
            year = int(year)
        except (TypeError, ValueError):
            self._send_json({"ok": False, "log": "", "error": "Term and Year must be numbers."}, 400)
            return

        buf = io.StringIO()
        out_path = None
        ok = True
        error = None
        with _lock:
            try:
                with contextlib.redirect_stdout(buf):
                    out_path = generate_top10.generate(
                        db_path=db_path,
                        db_password=db_password,
                        term=term,
                        year=year,
                        output_folder=output_folder,
                    )
            except PermissionError as exc:
                ok = False
                locked_file = getattr(exc, "filename", None) or str(exc)
                error = (
                    f"Can't save the workbook - it looks like it's already open "
                    f"in Excel (or another program):\n{locked_file}\n\n"
                    "Close it and click GENERATE again."
                )
                buf.write("\n" + traceback.format_exc())
            except Exception as exc:  # noqa: BLE001 - surface any failure to the UI
                ok = False
                error = str(exc)
                buf.write("\n" + traceback.format_exc())

        self._send_json({
            "ok": ok,
            "log": buf.getvalue(),
            "error": error,
            "out_path": out_path,
        })

    def _handle_generate_annexurek(self, data):
        db_path = (data.get("db_path") or "").strip()
        db_password = (data.get("db_password") or "").strip() or None
        term = data.get("term")
        year = data.get("year")
        groupings = data.get("groupings") or []
        templates = data.get("templates") or {}
        output_folder = (data.get("output_folder") or "").strip() or None

        if not db_path:
            self._send_json({"ok": False, "log": "", "error": "No database selected."}, 400)
            return
        if not os.path.isfile(db_path):
            self._send_json(
                {"ok": False, "log": "", "error": f"Database file not found:\n{db_path}"},
                400,
            )
            return
        try:
            term = int(term)
            year = int(year)
        except (TypeError, ValueError):
            self._send_json({"ok": False, "log": "", "error": "Term and Year must be numbers."}, 400)
            return
        if not groupings:
            self._send_json(
                {"ok": False, "log": "", "error": "Select at least one grouping (8-12 or R-7)."}, 400
            )
            return

        selected_templates = {}
        for grouping in groupings:
            template_path = (templates.get(grouping) or "").strip()
            if not template_path:
                self._send_json(
                    {"ok": False, "log": "", "error": f"No template selected for {grouping}."}, 400
                )
                return
            if not os.path.isfile(template_path):
                self._send_json(
                    {"ok": False, "log": "",
                     "error": f"Template file not found for {grouping}:\n{template_path}"}, 400
                )
                return
            selected_templates[grouping] = template_path

        buf = io.StringIO()
        out_paths = None
        ok = True
        error = None
        with _lock:
            try:
                with contextlib.redirect_stdout(buf):
                    out_paths = generate_annexure_k.generate(
                        db_path=db_path,
                        db_password=db_password,
                        data_year=year,
                        term=term,
                        groupings=groupings,
                        templates=selected_templates,
                        output_folder=output_folder,
                    )
            except PermissionError as exc:
                ok = False
                locked_file = getattr(exc, "filename", None) or str(exc)
                error = (
                    f"Can't save the workbook - it looks like it's already open "
                    f"in Excel (or another program):\n{locked_file}\n\n"
                    "Close it and click GENERATE again."
                )
                buf.write("\n" + traceback.format_exc())
            except Exception as exc:  # noqa: BLE001 - surface any failure to the UI
                ok = False
                error = str(exc)
                buf.write("\n" + traceback.format_exc())

        self._send_json({
            "ok": ok,
            "log": buf.getvalue(),
            "error": error,
            "out_paths": out_paths,
        })


def main():
    server = HTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}/"
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    print(f"Top 10 Club running at {url}  (close this window to stop)")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
