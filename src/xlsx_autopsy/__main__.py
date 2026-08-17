"""Module entry for ``uv run python -m xlsx_autopsy``.

Usage:
    uv run python -m xlsx_autopsy --excel workbook.xlsx
"""

from xlsx_autopsy.reconstruct import main

if __name__ == "__main__":
    main()
