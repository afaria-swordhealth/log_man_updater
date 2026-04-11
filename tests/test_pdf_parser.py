"""
tests/test_pdf_parser.py
========================
Unit tests for ``parse_invoice_pdf`` in update_logistics.py.

All tests mock ``pdfplumber.open`` so no real PDF files are required.
Synthetic page text is crafted to exercise each regex branch, including
regression coverage for Bug 10 (multi-token reference field and Portuguese
thousands-separator amounts).
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import datetime
from unittest.mock import MagicMock, patch

import pytest

from update_logistics import parse_invoice_pdf


# ── Mock helpers ──────────────────────────────────────────────────────────────


def _make_pdf_mock(page_texts):
    """
    Build a MagicMock that mimics a pdfplumber PDF context-manager.

    Parameters
    ----------
    page_texts : list[str | None]
        One string per page.  None is treated as a page whose
        ``extract_text()`` returns None (edge-case test support).
    """
    mock_pdf = MagicMock()
    mock_pages = []
    for text in page_texts:
        p = MagicMock()
        p.extract_text.return_value = text
        mock_pages.append(p)
    mock_pdf.pages = mock_pages
    mock_pdf.__enter__ = lambda s: s
    mock_pdf.__exit__ = MagicMock(return_value=False)
    return mock_pdf


def _call(page_texts):
    """Convenience wrapper: patch pdfplumber and call parse_invoice_pdf."""
    with patch("pdfplumber.open", return_value=_make_pdf_mock(page_texts)):
        return parse_invoice_pdf("fake.pdf")


# ── Header parsing ────────────────────────────────────────────────────────────


class TestParseHeader:
    """Tests for page-1 invoice metadata extraction."""

    def test_parse_header_standard(self):
        """Standard header lines produce correct invoice_no and invoice_date."""
        page1 = (
            "Número da Fatura: FT V/424543\n"
            "Data Da Fatura: 15-03-2026\n"
        )
        invoice_no, invoice_date, shipments = _call([page1])

        assert invoice_no == "FT V/424543"
        assert invoice_date == datetime(2026, 3, 15)
        assert shipments == {}

    def test_parse_header_with_space_in_invoice_number(self):
        """
        Invoice number containing an internal space: 'FT V/424 543'.

        The regex ``FT\\s*\\S+`` captures 'FT' + optional whitespace +
        the next run of non-space characters.  When the text reads
        'FT V/424 543', the ``\\s*`` consumes the space and ``\\S+``
        matches 'V/424', stopping before the second space.  The result
        is therefore 'FT V/424'.
        """
        page1 = "Número da Fatura: FT V/424 543\n"
        invoice_no, invoice_date, _ = _call([page1])

        # Regex captures up to the first whitespace boundary after FT + token.
        assert invoice_no == "FT V/424"

    def test_parse_no_header_fields(self):
        """Page 1 with neither expected field → both metadata values are None."""
        page1 = "Algum texto sem campos relevantes\n"
        invoice_no, invoice_date, shipments = _call([page1])

        assert invoice_no is None
        assert invoice_date is None
        assert shipments == {}

    def test_parse_header_case_insensitive(self):
        """Header keywords are matched case-insensitively."""
        page1 = (
            "NÚMERO DA FATURA: FT V/100001\n"
            "DATA DA FATURA: 01-01-2025\n"
        )
        invoice_no, invoice_date, _ = _call([page1])

        assert invoice_no == "FT V/100001"
        assert invoice_date == datetime(2025, 1, 1)

    def test_parse_header_accented_variant(self):
        """'Numero' (without accent) is also accepted by the regex."""
        page1 = "Numero da Fatura: FT V/200002\n"
        invoice_no, _, _ = _call([page1])

        assert invoice_no == "FT V/200002"


# ── Amount regex ──────────────────────────────────────────────────────────────


class TestParseAmount:
    """Tests for the standalone-amount regex on shipment detail pages."""

    def _shipment_page(self, amount_line):
        """Build a minimal two-page mock: empty header + one shipment block."""
        page1 = ""
        page2 = f"1234567890 15-03-2026\n{amount_line}\n"
        return _call([page1, page2])

    def test_parse_amount_simple_comma(self):
        """Plain two-decimal comma format '26,61' → 26.61."""
        _, _, shipments = self._shipment_page("26,61")
        assert shipments == {"1234567890": 26.61}

    def test_parse_amount_pt_thousands(self):
        """Portuguese thousands separator '2.840,95' → 2840.95 (Bug 10 regression)."""
        _, _, shipments = self._shipment_page("2.840,95")
        assert shipments == {"1234567890": 2840.95}

    def test_parse_amount_large_thousands(self):
        """Larger PT-format amount '5.572,24' → 5572.24 (Bug 10 regression)."""
        _, _, shipments = self._shipment_page("5.572,24")
        assert shipments == {"1234567890": 5572.24}

    def test_parse_amount_not_matched_mid_line(self):
        """Amount embedded mid-line ('total 26,61 euros') must not be captured."""
        page1 = ""
        page2 = "1234567890 15-03-2026\ntotal 26,61 euros\n"
        _, _, shipments = _call([page1, page2])

        # No valid standalone amount was found; the shipment has no amount and
        # must not appear in the result dict.
        assert "1234567890" not in shipments

    def test_parse_amount_dot_decimal_format(self):
        """Dot-decimal format '126.09' (non-PT) is also parsed correctly."""
        _, _, shipments = self._shipment_page("126.09")
        assert shipments == {"1234567890": 126.09}


# ── Tracking / shipment block detection ──────────────────────────────────────


class TestParseShipment:
    """Tests for shipment block start detection and dict assembly."""

    def test_parse_shipment_simple(self):
        """Basic tracking + date on same line, amount on next line."""
        page1 = ""
        page2 = "1234567890 15-03-2026\n26,61\n"
        _, _, shipments = _call([page1, page2])

        assert shipments == {"1234567890": 26.61}

    def test_parse_shipment_multitoken_reference(self):
        """
        Multi-token reference field between tracking and date must not
        prevent the block from being recognised (Bug 10 regression).

        Input line: '1234567890 RTO 7981361555 15-03-2026'
        The tracking is the first 10-digit token; the date appears
        anywhere on the line.
        """
        page1 = ""
        page2 = "1234567890 RTO 7981361555 15-03-2026\n26,61\n"
        _, _, shipments = _call([page1, page2])

        assert shipments == {"1234567890": 26.61}

    def test_parse_tracking_without_date_ignored(self):
        """
        A line starting with 10 digits but containing no DD-MM-YYYY date
        does not open a new shipment block.
        """
        page1 = ""
        page2 = "1234567890 referencia\n26,61\n"
        _, _, shipments = _call([page1, page2])

        # No tracking block was opened, so the amount floats with no owner.
        assert shipments == {}

    def test_parse_multiple_shipments(self):
        """Three consecutive shipment blocks are all captured correctly."""
        page1 = ""
        page2 = (
            "1234567890 15-03-2026\n"
            "26,61\n"
            "1234567891 15-03-2026\n"
            "50,00\n"
            "1234567892 15-03-2026\n"
            "100,00\n"
        )
        _, _, shipments = _call([page1, page2])

        assert len(shipments) == 3
        assert shipments["1234567890"] == 26.61
        assert shipments["1234567891"] == 50.00
        assert shipments["1234567892"] == 100.00

    def test_parse_last_shipment_committed(self):
        """
        The final shipment in the file has no following tracking line to
        trigger the in-loop commit.  The post-loop commit must still add it.
        """
        page1 = ""
        page2 = (
            "1234567890 15-03-2026\n"
            "26,61\n"
            "9876543210 15-03-2026\n"
            "999,99\n"
            # ← no third tracking line follows; 9876543210 is the last
        )
        _, _, shipments = _call([page1, page2])

        assert "9876543210" in shipments
        assert shipments["9876543210"] == 999.99

    def test_parse_empty_page(self):
        """An empty detail page produces no shipments and no crash."""
        page1 = ""
        page2 = ""
        _, _, shipments = _call([page1, page2])

        assert shipments == {}

    def test_parse_empty_page_none_text(self):
        """extract_text() returning None is handled via the '… or \"\"' guard."""
        page1 = None
        page2 = None
        invoice_no, invoice_date, shipments = _call([page1, page2])

        assert invoice_no is None
        assert invoice_date is None
        assert shipments == {}


# ── Full-invoice integration simulation ──────────────────────────────────────


class TestParseFullInvoiceSimulation:
    """End-to-end simulation of a realistic multi-page invoice."""

    def test_parse_full_invoice_simulation(self):
        """
        Simulate a two-page invoice with 5 shipments, including:
        - a multi-token reference line (Bug 10)
        - a Portuguese thousands-separator amount (Bug 10)
        - a mix of plain comma amounts

        All 5 shipments must be present in the returned dict with correct
        float values, and header metadata must be extracted correctly.
        """
        page1 = (
            "Número da Fatura: FT V/424543\n"
            "Data Da Fatura: 15-03-2026\n"
            "Algum texto de rodapé irrelevante\n"
        )
        page2 = (
            # Shipment 1 — simple
            "1111111111 15-03-2026\n"
            "26,61\n"
            # Shipment 2 — multi-token reference (Bug 10)
            "2222222222 RTO 7981361555 15-03-2026\n"
            "126,09\n"
            # Shipment 3 — thousands separator (Bug 10)
            "3333333333 15-03-2026\n"
            "2.840,95\n"
            # Shipment 4 — another thousands value
            "4444444444 15-03-2026\n"
            "5.572,24\n"
            # Shipment 5 — plain comma, last entry (tests post-loop commit)
            "5555555555 15-03-2026\n"
            "88,50\n"
        )

        invoice_no, invoice_date, shipments = _call([page1, page2])

        assert invoice_no == "FT V/424543"
        assert invoice_date == datetime(2026, 3, 15)

        assert len(shipments) == 5
        assert shipments["1111111111"] == 26.61
        assert shipments["2222222222"] == 126.09
        assert shipments["3333333333"] == 2840.95
        assert shipments["4444444444"] == 5572.24
        assert shipments["5555555555"] == 88.50

    def test_parse_multipage_detail(self):
        """
        Shipments spread across three detail pages are all captured.
        The ``for page in pdf.pages[1:]`` loop must process every page.
        """
        page1 = "Número da Fatura: FT V/500000\nData Da Fatura: 01-04-2026\n"
        page2 = "6666666666 01-04-2026\n10,00\n7777777777 01-04-2026\n20,00\n"
        page3 = "8888888888 01-04-2026\n30,00\n"

        invoice_no, invoice_date, shipments = _call([page1, page2, page3])

        assert invoice_no == "FT V/500000"
        assert len(shipments) == 3
        assert shipments["6666666666"] == 10.00
        assert shipments["7777777777"] == 20.00
        assert shipments["8888888888"] == 30.00
