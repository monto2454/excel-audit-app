import unicodedata
from io import BytesIO

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    session,
    redirect,
    url_for,
)
import pandas as pd

app = Flask(__name__)
# Needed for using session (for error report storage)
app.secret_key = "change-this-to-a-random-secret-key"


# ---------- Helper: convert 0-based column index → Excel letters ----------
def col_index_to_letter(idx_zero_based: int) -> str:
    idx = idx_zero_based + 1  # convert to 1-based
    letters = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


# ---------- Core validation function ----------
def validate_dataframe(df: pd.DataFrame):
    errors = []

    # --------------------------------------------
    # Breadcrumb column is ALWAYS column R
    # R (1-based) -> index 17 (0-based: A=0, ..., R=17)
    # --------------------------------------------
    breadcrumb_col_index = 17  # Column R
    breadcrumb_col_name = df.columns[breadcrumb_col_index]

    # Track duplicate breadcrumbs using normalized values
    breadcrumb_seen = {}

    # Loop through every cell in the file
    for row_idx, (_, row) in enumerate(df.iterrows()):
        for col_idx, col_name in enumerate(df.columns):

            value = row[col_name]

            # Normalize empty/missing values
            if pd.isna(value):
                text = ""
            else:
                text = str(value)

            if text == "":
                continue

            stripped = text.strip()
            col_letter = col_index_to_letter(col_idx)
            row_number = row_idx + 2  # first data row = Excel row 2
            cell_address = f"{col_letter}{row_number}"

            # 1) Trailing spaces
            if text != text.rstrip(" "):
                errors.append((cell_address, "Trailing space at end of value"))

            # 2) Double spaces between words
            if "  " in text:
                errors.append((cell_address, "Multiple spaces between words"))

            # 4) Colon
            if ":" in text:
                errors.append((cell_address, "Contains colon ':'"))

            # 6) Space after (
            if "( " in text:
                errors.append((cell_address, "Space after '('"))

            # 7) Space before )
            if " )" in text:
                errors.append((cell_address, "Space before ')'"))

            # 8) Space around hyphen
            if " -" in text or "- " in text:
                errors.append((cell_address, "Space before or after '-'"))

            # 9) Unbalanced parentheses
            if text.count("(") != text.count(")"):
                errors.append((cell_address, "Unbalanced parentheses '(' and ')'"))

            # 10) Full-stop
            if "." in text:
                errors.append((cell_address, "Contains full-stop '.'"))

            # 11) Title Case detection (2+ words, each capitalized, not ALL CAPS)
            words = stripped.split()
            if len(words) >= 2:
                is_title_case = all(
                    w[0].isalpha()
                    and w[0].isupper()
                    and w[1:].islower()
                    for w in words
                    if w.isalpha()
                )
                if is_title_case and not stripped.isupper():
                    errors.append((cell_address, "Text appears to be in Title Case"))

            # 12) Accents / non-ASCII characters
            has_accent = False
            for ch in text:
                if ord(ch) > 127:
                    has_accent = True
                    break
                decomp = unicodedata.normalize("NFD", ch)
                if any(unicodedata.category(c) == "Mn" for c in decomp):
                    has_accent = True
                    break
            if has_accent:
                errors.append(
                    (cell_address, "Contains accented or non-ASCII characters")
                )

            # ---------- Breadcrumb-only rules ----------
            if col_idx == breadcrumb_col_index:  # ONLY column R
                display_value = stripped

                # Normalize for duplicate detection
                normalized_key = " ".join(stripped.split()).lower()

                # 3) Duplicate breadcrumb
                if normalized_key:
                    if normalized_key in breadcrumb_seen:
                        first_cell = breadcrumb_seen[normalized_key]
                        msg = (
                            f"Duplicate breadcrumb '{display_value}' "
                            f"(first seen at {first_cell})"
                        )
                        errors.append((cell_address, msg))
                    else:
                        breadcrumb_seen[normalized_key] = cell_address

                # 5) Breadcrumb ends with >
                if display_value.endswith(">"):
                    errors.append((cell_address, "Breadcrumb ends with '>'"))

    return errors, breadcrumb_col_name


# ---------- Routes ----------
@app.route("/", methods=["GET", "POST"])
def index():
    errors = None
    breadcrumb_col_name = None
    file_name = None

    # Clear previous error report by default
    session["error_report"] = None

    if request.method == "POST":
        uploaded_file = request.files.get("file")

        if uploaded_file and uploaded_file.filename:
            file_name = uploaded_file.filename

            try:
                df = pd.read_excel(uploaded_file, dtype=str)
                errors, breadcrumb_col_name = validate_dataframe(df)

                # Store error data in session for export as Excel
                if errors:
                    error_rows = [
                        {"Cell": cell, "Issue": msg} for (cell, msg) in errors
                    ]
                    session["error_report"] = error_rows
                else:
                    session["error_report"] = None

            except Exception as e:
                errors = [("N/A", f"Error reading file: {e}")]
                session["error_report"] = None

    return render_template(
        "index.html",
        errors=errors,
        breadcrumb_col_name=breadcrumb_col_name,
        file_name=file_name,
    )


@app.route("/download-report")
def download_report():
    data = session.get("error_report")

    if not data:
        # No report available – just go back to main page
        return redirect(url_for("index"))

    # Create DataFrame from stored errors
    df_errors = pd.DataFrame(data)

    # Write to in-memory Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_errors.to_excel(writer, index=False, sheet_name="Errors")
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="error_report.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
    )


if __name__ == "__main__":
    app.run(debug=True)
