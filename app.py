import math
from io import BytesIO
from flask import Flask, render_template, request, send_file
import pandas as pd
import openpyxl

app = Flask(__name__)
app.config["LAST_ERRORS"] = []
app.config["LAST_FILENAME"] = None


# ----------------- Helper: Excel column index -> letter -----------------
def col_index_to_letter(idx_zero_based: int) -> str:
    idx = idx_zero_based + 1  # convert to 1-based
    letters = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


# ----------------- Full-stop rule (Rule 8) -----------------
def has_full_stop_issue(text: str) -> bool:
    """
    Rule 8: A full stop (.) is an issue ONLY when it appears at the END of the value.
    """
    stripped = text.rstrip()
    return stripped.endswith(".")


# ----------------- Title-case rule (Rule 9 – non-Breadcrumb columns only) -----------------
def is_title_case_issue(text: str) -> bool:
    """
    Applies ONLY when:
      - cell has 2 or more words
      - column is NOT Breadcrumbs

    Rules:
      - First word must start uppercase
      - All other words must start lowercase
      - Hyphens remain part of a word
    """
    stripped = text.strip()
    words = stripped.split()

    if len(words) < 2:
        return False

    def first_alpha_char(word: str):
        for ch in word:
            if ch.isalpha():
                return ch
        return None

    # First word must start uppercase
    first = words[0]
    first_ch = first_alpha_char(first)
    if first_ch is not None and not first_ch.isupper():
        return True

    # Remaining words must start lowercase
    for w in words[1:]:
        ch = first_alpha_char(w)
        if ch is None:
            continue
        if ch.isupper():
            return True

    return False


# ----------------- Core validation logic (all 12 rules) -----------------
def validate_audit_sheet(df: pd.DataFrame):
    errors = []
    summary = {
        "Trailing spaces": 0,
        "Double spaces": 0,
        "Colons": 0,
        "Spaces after (": 0,
        "Spaces before )": 0,
        "Spaces around -": 0,
        "Unbalanced brackets": 0,
        "Full stop issues": 0,
        "Title case issues": 0,
        "Accents / non-ASCII": 0,
        "Breadcrumb ends with >": 0,
        "Duplicate breadcrumbs": 0,
    }

    # Detect Breadcrumbs column (case-insensitive)
    breadcrumb_col_name = None
    for col in df.columns:
        if str(col).strip().lower() == "breadcrumbs":
            breadcrumb_col_name = col
            break

    breadcrumb_seen = {}  # value -> first cell reference

    for row_idx, row in df.iterrows():
        for col_idx, col_name in enumerate(df.columns):
            value = row[col_name]
            text = "" if pd.isna(value) else str(value)

            if text.strip() == "":
                continue

            clean_text = text.strip()
            cell_address = f"{col_index_to_letter(col_idx)}{row_idx + 2}"
            issues = []

            # ------------------ Breadcrumb column (Only Rule 11 + Rule 12) ------------------
            if breadcrumb_col_name is not None and col_name == breadcrumb_col_name:
                norm = clean_text

                # Rule 12: Duplicate Breadcrumbs
                if norm:
                    if norm in breadcrumb_seen:
                        first_cell = breadcrumb_seen[norm]
                        issues.append(f"Duplicate breadcrumb (also in {first_cell})")
                        summary["Duplicate breadcrumbs"] += 1
                    else:
                        breadcrumb_seen[norm] = cell_address

                # Rule 11: Ends with '>'
                if norm.endswith(">"):
                    issues.append("Breadcrumb ends with '>'")
                    summary["Breadcrumb ends with >"] += 1

                if issues:
                    errors.append({"cell": cell_address, "value": text, "issues": issues})

                continue  # Skip remaining rules for Breadcrumbs

            # ------------------ Non-Breadcrumb columns (Rules 1–10) ------------------

            # Rule 1: Trailing spaces
            if text != text.rstrip(" "):
                issues.append("Trailing spaces found")
                summary["Trailing spaces"] += 1

            # Rule 2: Extra spaces between words
            if "  " in text:
                issues.append("Double spaces detected")
                summary["Double spaces"] += 1

            # Rule 3: Colon check
            if ":" in text:
                issues.append("Contains colon ':'")
                summary["Colons"] += 1

            # Rule 4: Space after '('
            if "( " in text:
                issues.append("Space after '('")
                summary["Spaces after ("] += 1

            # Rule 5: Space before ')'
            if " )" in text:
                issues.append("Space before ')'")
                summary["Spaces before )"] += 1

            # Rule 6: Spaces around hyphens
            if " -" in text or "- " in text:
                issues.append("Space around '-'")
                summary["Spaces around -"] += 1

            # Rule 7: Unbalanced brackets
            if text.count("(") != text.count(")"):
                issues.append("Unbalanced brackets")
                summary["Unbalanced brackets"] += 1

            # Rule 8: Full stop (only if ends with '.')
            if has_full_stop_issue(text):
                issues.append("Ends with full stop")
                summary["Full stop issues"] += 1

            # Rule 9: Title case
            if is_title_case_issue(text):
                issues.append("Title case detected")
                summary["Title case issues"] += 1

            # Rule 10: Accents / non-ASCII
            has_accent = any(ord(ch) > 127 for ch in text)
            if has_accent:
                issues.append("Contains accented or non-ASCII characters")
                summary["Accents / non-ASCII"] += 1

            # Store row issues
            if issues:
                errors.append({"cell": cell_address, "value": text, "issues": issues})

    return errors, summary


# ----------------- Flask Routes -----------------
@app.route("/", methods=["GET", "POST"])
def index():
    errors = []
    summary = None
    message = None
    filename = None
    had_run = False

    if request.method == "POST":
        had_run = True
        uploaded = request.files.get("file")

        if not uploaded or uploaded.filename.strip() == "":
            message = "Please upload an Excel file."
        else:
            filename = uploaded.filename
            try:
                file_bytes = uploaded.read()
                excel_io = BytesIO(file_bytes)

                xls = pd.ExcelFile(excel_io)
                sheet_names = [s.strip() for s in xls.sheet_names]

                if "Audit" not in sheet_names:
                    message = "Audit sheet not found in the file"
                else:
                    excel_io.seek(0)
                    df = pd.read_excel(excel_io, sheet_name="Audit", dtype=str)

                    errors, summary = validate_audit_sheet(df)
                    app.config["LAST_ERRORS"] = errors
                    app.config["LAST_FILENAME"] = filename

                    if not errors:
                        message = "No issues found."

            except Exception as e:
                message = f"Error reading file: {e}"

    return render_template(
        "index.html",
        errors=errors,
        summary=summary,
        message=message,
        filename=filename,
        had_run=had_run,
    )


@app.route("/download")
def download():
    errors = app.config.get("LAST_ERRORS") or []

    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Errors"

    ws.append(["Cell", "Cell Value", "Issues Found"])

    if errors:
        for e in errors:
            ws.append([e["cell"], e["value"], "; ".join(e["issues"])])
    else:
        ws.append(["N/A", "N/A", "No issues found"])

    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="audit-errors.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
