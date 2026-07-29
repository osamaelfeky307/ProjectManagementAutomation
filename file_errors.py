"""
Structured error classes for file upload operations.

FileUploadError provides a title, user-friendly message, and actionable
suggestions so the UI layer can show rich error dialogs instead of raw
Python tracebacks.
"""


class FileUploadError(Exception):
    """Base error for file upload issues, with title and user-friendly detail."""

    def __init__(self, title: str, message: str, suggestions: list[str] | None = None):
        self.title = title
        self.message = message
        self.suggestions = suggestions or []
        # Build a full string so str(err) is still useful in logs/fallbacks
        full = message
        if self.suggestions:
            full += "\n\nSuggestions:\n" + "\n".join(f"  • {s}" for s in self.suggestions)
        super().__init__(full)


class FileReadError(FileUploadError):
    """The file could not be opened or decoded."""

    def __init__(self, filename: str, reason: str = ""):
        detail = reason or "The file could not be read."
        super().__init__(
            title="File Read Error",
            message=f"Could not read '{filename}'.\n\n{detail}",
            suggestions=[
                "Make sure the file is not open in another program (e.g. Excel).",
                "Verify the file is a valid .xlsx, .xlsm, or .csv file.",
                "If the file is password-protected, remove the password and try again.",
            ],
        )


class EmptyFileError(FileUploadError):
    """The file exists but contains no usable data."""

    def __init__(self, filename: str):
        super().__init__(
            title="Empty File",
            message=f"The file '{filename}' appears to be empty or contains no data rows.",
            suggestions=[
                "Make sure your data starts on row 2 (row 1 should be column headers).",
                "Check that the correct sheet is active if the workbook has multiple sheets.",
            ],
        )


class MissingColumnsError(FileUploadError):
    """Required columns were not found in the file headers."""

    def __init__(self, missing: list[str], detected: list[str] | None = None,
                 accepted_hints: dict[str, set[str]] | None = None):
        lines = [f"Missing required columns: {', '.join(missing)}"]

        if detected:
            lines.append(f"\nColumns detected in your file:\n  {', '.join(detected)}")

        if accepted_hints:
            lines.append("\nAccepted column name variants:")
            for col_name, variants in accepted_hints.items():
                if col_name in missing:
                    lines.append(f"  • {col_name}: {', '.join(sorted(variants))}")

        super().__init__(
            title="Missing Columns",
            message="\n".join(lines),
            suggestions=[
                "Rename your column headers to match one of the accepted variants above.",
                "Make sure headers are on row 1 of the sheet.",
            ],
        )


class InvalidPhoneWarning:
    """Non-exception container for rows with invalid phone numbers (used in summaries)."""

    def __init__(self, row_number: int, name: str, raw_phone: str, reason: str = ""):
        self.row_number = row_number
        self.name = name
        self.raw_phone = raw_phone
        self.reason = reason or "Phone number does not match the expected Egyptian format (+20XXXXXXXXXX)."

    def __str__(self):
        return f"Row {self.row_number}: {self.name} — phone '{self.raw_phone}' — {self.reason}"


class CoordinatorDirectoryError(FileUploadError):
    """The coordinator directory file is missing required name/phone columns."""

    def __init__(self, filename: str, detected_headers: list[str] | None = None):
        lines = [
            f"The coordinator data file '{filename}' must contain columns for "
            f"coordinator names and phone numbers."
        ]

        lines.append("\nAccepted column names for coordinator name:")
        lines.append("  coordinatorname, coordinator, name, fullname, اسم المنسق, الاسم, اسم")
        lines.append("\nAccepted column names for phone number:")
        lines.append("  phonenumber, phone, mobile, mobilenumber, whatsapp, رقم الهاتف, الموبايل")

        if detected_headers:
            lines.append(f"\nColumns found in your file:\n  {', '.join(detected_headers)}")

        super().__init__(
            title="Coordinator Data Error",
            message="\n".join(lines),
            suggestions=[
                "Rename your columns to match one of the accepted names above.",
                "Make sure the first row contains column headers.",
            ],
        )
