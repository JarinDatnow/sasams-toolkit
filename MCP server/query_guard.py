"""
Read-only, privacy-preserving checks for SQL that Claude writes itself.

Every free-form query goes through check_sql() before it touches the
database, and every result goes through redact_rows() before it is handed
back to Claude. The goal is simple: marks, counts and averages can leave the
PC; names, ID numbers, contact details, medical notes and free-text comments
never do.
"""

import json
import re
from pathlib import Path

# Academic tables Claude may query directly. Everything else (staff, parents,
# finances, discipline, medical, etc.) is off-limits to free-form SQL.
ALLOWED_TABLES = {
    "Learner_Info",
    "Classes",
    "Subjects",
    "ReportCycles",
    "ReportMarks",
    "SubjectCriteria",
    "LearnerCass",
    "LearnerSubjects",
    "SchoolTerms",
    "Ana2012EvaluationLevels",
    "Absentees",
}

# Learner_Info holds almost every personal detail SASAMS keeps, so only these
# columns of it are usable. ID is SASAMS's internal row number: it means
# nothing outside the database, so it is the learner's pseudonym here.
LEARNER_INFO_ALLOWED = {
    "ID", "Grade", "Class", "Status", "Gender", "GradeYears", "GradeEntered",
    "PhaseYears", "ProgressedToGrade", "Repeat12",
}

# Column names that identify a person or hold sensitive free text, anywhere.
SENSITIVE_PATTERN = re.compile(
    r"(sname|fname|surname|firstname|nickname|secondname|thirdname|fullname|"
    r"learnername|initial|idno|id_?number|identity|foreignid|passport|permit|"
    r"birth|dob|address|street|suburb|postal|tel|phone|cell|fax|email|e_mail|"
    r"parent|guardian|kin|ice(name|surname|relation)|accession|lurits|examnr|"
    r"sams12lrnid|medical|clinic|^dr|doctor|disab|lsen|religion|race|"
    r"grant|social|pastel|photo|picture|signature|bank|account|salary|persal|"
    r"password|pwd|comment|reason|disciplin|guidance|transport|boarder|hostel|"
    r"deceased|family|immuniz|road2health|feeexempt|mentor|hseid|busroute)",
    re.IGNORECASE,
)

# Names that look sensitive but are harmless in the allowed tables
# (Subjects.Name, ReportCycles.Name, Classes.ClassName, etc.).
SAFE_OVERRIDES = {"name", "classname", "afrname", "id"}

FORBIDDEN_KEYWORDS = re.compile(
    r"\b(insert|update|delete|drop|alter|create|into|exec|execute|transform|"
    r"parameters|grant|revoke|merge|truncate|openrowset|opendatasource)\b",
    re.IGNORECASE,
)

MAX_ROWS_HARD_LIMIT = 1000


class QueryRejected(ValueError):
    """Raised with a plain-language reason Claude can act on."""


def load_schema(map_path):
    data = json.loads(Path(map_path).read_text(encoding="utf-8"))
    return {
        name: [c["name"] for c in info["columns"]]
        for name, info in data["tables"].items()
    }


def is_sensitive_column(name):
    if name.lower() in SAFE_OVERRIDES:
        return False
    return bool(SENSITIVE_PATTERN.search(name))


class QueryGuard:
    def __init__(self, schema):
        self.schema = schema
        self.tables_lower = {t.lower(): t for t in schema}
        self.allowed_tables_lower = {t.lower() for t in ALLOWED_TABLES}

        other_allowed_columns = set()
        for table in ALLOWED_TABLES:
            if table != "Learner_Info":
                other_allowed_columns |= {c.lower() for c in schema.get(table, [])}

        blocked = set()
        for table in ALLOWED_TABLES:
            for col in schema.get(table, []):
                if is_sensitive_column(col):
                    blocked.add(col.lower())
        for col in schema.get("Learner_Info", []):
            low = col.lower()
            if col not in LEARNER_INFO_ALLOWED and low not in other_allowed_columns:
                blocked.add(low)
        self.blocked_columns = blocked

    def table_is_queryable(self, table):
        return table.lower() in self.allowed_tables_lower

    def column_is_hidden(self, table, column):
        if table == "Learner_Info" and column not in LEARNER_INFO_ALLOWED:
            return True
        return column.lower() in self.blocked_columns or is_sensitive_column(column)

    def check_sql(self, sql):
        """Return a cleaned single SELECT statement, or raise QueryRejected."""
        if not sql or not sql.strip():
            raise QueryRejected("The query is empty.")

        text = re.sub(r"--[^\n]*", " ", sql)
        text = re.sub(r"/\*.*?\*/", " ", text, flags=re.DOTALL)
        text = text.strip().rstrip(";").strip()

        if ";" in text:
            raise QueryRejected("Only one statement at a time is allowed.")
        if not re.match(r"^select\b", text, re.IGNORECASE):
            raise QueryRejected("Only SELECT queries are allowed. This server never changes the database.")

        no_strings = re.sub(r"'(?:[^']|'')*'", "''", text)
        no_strings = re.sub(r'"(?:[^"]|"")*"', '""', no_strings)

        bad = FORBIDDEN_KEYWORDS.search(no_strings)
        if bad:
            raise QueryRejected(f"'{bad.group(1).upper()}' is not allowed. Queries are read-only.")
        indirect = re.search(r"\b(dlookup|dfirst|dlast|dmax|dmin|dsum|dcount|davg|dstdev|dvar|eval)\s*\(", no_strings, re.IGNORECASE)
        if indirect:
            raise QueryRejected(
                f"{indirect.group(1)}() is not allowed: it reads columns by name from text, "
                "which would get around the privacy checks."
            )
        if re.search(r"\bin\s*''", no_strings, re.IGNORECASE) or re.search(r"\[\s*[a-z]+\s*;", text, re.IGNORECASE):
            raise QueryRejected("Queries may not reach outside this database (external IN '...' or connection strings).")

        star_check = re.sub(r"count\s*\(\s*\*\s*\)", "count(x)", no_strings, flags=re.IGNORECASE)
        if re.search(
            r"(\bselect\s+(distinct\s+)?(top\s+\d+\s+)?\*)|(,\s*\*)|([\w\]]\.\*)",
            star_check,
            re.IGNORECASE,
        ):
            raise QueryRejected(
                "SELECT * is not allowed: name the columns you need. "
                "Use describe_table to see which columns are available."
            )

        identifiers = set()
        for bracketed in re.findall(r"\[([^\]]+)\]", no_strings):
            identifiers.add(bracketed.strip())
        for word in re.findall(r"[A-Za-z_][A-Za-z0-9_]*", re.sub(r"\[[^\]]+\]", " ", no_strings)):
            identifiers.add(word)

        for ident in identifiers:
            low = ident.lower()
            if low in self.tables_lower and low not in self.allowed_tables_lower:
                raise QueryRejected(
                    f"Table '{self.tables_lower[low]}' is not available. Only academic tables can be queried: "
                    + ", ".join(sorted(ALLOWED_TABLES))
                )
            if low in self.blocked_columns:
                raise QueryRejected(
                    f"Column '{ident}' holds personal information and can't be queried. "
                    "Work with the learner's ID instead, and use export_learner_list "
                    "to give the teacher a named list on this PC."
                )

        for target in re.findall(r"\b(?:from|join)\s+\[?([A-Za-z_][A-Za-z0-9_ ]*?)\]?(?=\s|$|\)|,)", no_strings, re.IGNORECASE):
            low = target.strip().lower()
            if low and low not in ("select",) and low not in self.allowed_tables_lower:
                raise QueryRejected(
                    f"Table '{target.strip()}' is not available. Only academic tables can be queried: "
                    + ", ".join(sorted(ALLOWED_TABLES))
                )
        return text

    def redact_rows(self, columns, rows):
        """Second line of defence: blank any result column that looks personal."""
        hidden = [is_sensitive_column(c) or c.lower() in self.blocked_columns for c in columns]
        if not any(hidden):
            return rows, []
        cleaned = []
        for row in rows:
            cleaned.append({
                col: ("[hidden]" if h else row[col]) for col, h in zip(columns, hidden)
            })
        return cleaned, [c for c, h in zip(columns, hidden) if h]
