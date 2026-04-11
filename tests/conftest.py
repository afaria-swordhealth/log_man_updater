"""
tests/conftest.py
=================
Shared pytest fixtures for the update_logistics test-suite.

All fixtures use the helpers in tests/fixtures/make_workbook.py so that
workbook structure stays consistent across every test module.

Fixtures
--------
tmp_excel                  – 10-row clean workbook (tmp_path scope)
tmp_excel_with_appended    – 10 main + 3 appended rows (Bug 9 scenario)
tmp_excel_orphan           – 10 rows + orphan cell at row 500 (Bug 1 scenario)
sample_invoice_data        – synthetic (invoice_no, invoice_date, shipments)
"""

from datetime import datetime

import pytest

from tests.fixtures.make_workbook import (
    make_clean_workbook,
    make_workbook_with_appended,
    make_workbook_with_orphan_cells,
)

# ── Workbook fixtures ─────────────────────────────────────────────────────────

@pytest.fixture
def tmp_excel(tmp_path):
    """
    A clean temporary workbook with 10 normal data rows.

    Row layout (all rows):
      col A  : sequential integer (1..10)
      col R  : sequential 10-digit tracking string
               ("1234567890", "1234567891", …, "1234567899")
      col U  : empty
      col V  : empty
      col W  : empty

    Returns the Path to the saved .xlsx file.
    """
    return make_clean_workbook(tmp_path / "test_clean.xlsx", n_rows=10)


@pytest.fixture
def tmp_excel_with_appended(tmp_path):
    """
    A temporary workbook that simulates the Bug 9 / appended-row scenario.

    10 main rows (col A present) followed by 3 appended rows (col A absent,
    only col R filled).  The script's extended scan must find all 13 tracking
    numbers even though ``last_data_row`` only reaches row 11.

    Returns the Path to the saved .xlsx file.
    """
    return make_workbook_with_appended(
        tmp_path / "test_appended.xlsx",
        n_main=10,
        n_appended=3,
    )


@pytest.fixture
def tmp_excel_orphan(tmp_path):
    """
    A temporary workbook that simulates the Bug 1 / orphan-cell scenario.

    10 normal rows plus a stray value in row 500 of a column beyond col W.
    This inflates ``ws.max_row`` to 500; the col-A scan must still report
    row 11 (the last real data row) as ``last_data_row``.

    Returns the Path to the saved .xlsx file.
    """
    return make_workbook_with_orphan_cells(
        tmp_path / "test_orphan.xlsx",
        n_rows=10,
        orphan_row=500,
    )


# ── Invoice data fixture ──────────────────────────────────────────────────────

@pytest.fixture
def sample_invoice_data():
    """
    Synthetic invoice data for use in integration / unit tests.

    Returns
    -------
    tuple[str, datetime, dict]
        invoice_no   : "FT V/999999"
        invoice_date : datetime(2026, 3, 15)
        shipments    : dict mapping tracking string → float amount

    Shipment breakdown (designed against a 10-row clean workbook):
    ──────────────────────────────────────────────────────────────
    "1234567892"  → 126.09   exists in workbook row 3  (col U empty  → fill)
    "1234567896"  → 55.40    exists in workbook row 7  (col U empty  → fill)
    "9999999999"  → 200.00   NOT in workbook            → append at end
    3248136404.0  → 88.50    numeric float tracking     → tests Bug 2
                             normalise_tracking strips ".0" → "3248136404"
                             not in workbook             → append at end
    "1234567891"  → 310.75   exists in workbook row 2  (col U empty  → fill)

    Notes
    -----
    - Rows 3 and 7 map to tracking "1234567892" and "1234567896" because the
      clean workbook starts tracking at "1234567890" on data row 2 (row 1 is
      the header), so data row n contains tracking "123456789{n-2}".
    - The float key ``3248136404.0`` is stored as a Python float to reproduce
      the openpyxl behaviour where numeric-looking cells are read as floats
      (Bug 2).  ``normalize_tracking`` must convert it to "3248136404".
    """
    invoice_no   = "FT V/999999"
    invoice_date = datetime(2026, 3, 15)

    shipments = {
        "1234567892": 126.09,   # row 3  — fill in place
        "1234567896":  55.40,   # row 7  — fill in place
        "9999999999": 200.00,   # missing — append at end
        3248136404.0:  88.50,   # float key (Bug 2) — normalize → append
        "1234567891": 310.75,   # row 2  — fill in place
    }

    return invoice_no, invoice_date, shipments
