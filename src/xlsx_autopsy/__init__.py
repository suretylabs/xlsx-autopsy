"""xlsx-autopsy: decompose a huge Excel workbook without opening it.

Usage:
    uv run xlsx-autopsy --excel workbook.xlsx
    uv run python -m xlsx_autopsy --excel workbook.xlsx -o out --skip-formulas
"""

from xlsx_autopsy.secrets import redact_connection_fields, redact_connection_string

__all__ = ["redact_connection_fields", "redact_connection_string"]
__version__ = "0.1.2"
