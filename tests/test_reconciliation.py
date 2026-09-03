"""
Comprehensive Unit & Integration Test Suite
============================================
Enterprise Finance Reconciliation Tool
"""

import os
import shutil
import tempfile
import unittest
from datetime import date, datetime
from decimal import Decimal

import pandas as pd

from reconciliation import (
    INVALID_VALUE_SENTINEL,
    STATUS_DUPLICATE_IN_ERP,
    STATUS_INVALID_FIELD_VALUE,
    STATUS_MATCH,
    STATUS_MISSING_REQUIRED_VALUE,
    OutputError,
    SchemaError,
    canonicalize_key,
    parse_date,
    parse_decimal,
    read_spreadsheet,
    reconcile_data,
    summarize_money_values,
    write_output,
)


class TestFourMandatoryGateTests(unittest.TestCase):
    """Four Mandatory Gate Tests for Fundamental Reconciliation Correctness."""

    def test_numeric_csv_vs_excel_key_reconciles(self):
        """Gate 1: CSV string key '1001' vs XLSX float key 1001.0 reconciles to MATCH."""
        df_erp = pd.DataFrame([{"invoice_id": "1001", "amount": "100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": 1001.0, "amount": "$100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["matches"]), 1)
        self.assertEqual(results["matches"]["reconciliation_status"].iloc[0], STATUS_MATCH)

    def test_default_schema_missing_currency_fails(self):
        """Gate 2: Missing required default schema column raises SchemaError without silent shrinking."""
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "customer_name": "Acme", "date": "2026-08-01"}])  # Missing currency!

        with self.assertRaises(SchemaError) as ctx:
            reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertIn("currency", str(ctx.exception))

    def test_required_amount_missing_both_never_matches(self):
        """Gate 3: Both missing required amount values become MISSING_REQUIRED_VALUE, never MATCH."""
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": None, "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["matches"]), 0)
        self.assertEqual(len(results["data_issues"]), 1)
        self.assertEqual(results["data_issues"]["reconciliation_status"].iloc[0], STATUS_MISSING_REQUIRED_VALUE)

    def test_invalid_missing_amount_not_counted_as_zero_exposure(self):
        """Gate 4: Invalid missing amounts are not silently converted to $0 in exposure KPIs."""
        df_erp = pd.DataFrame([{"invoice_id": "INV-500", "amount": "ERROR123", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])
        df_ext = pd.DataFrame(columns=["invoice_id", "amount", "currency", "customer_name", "date"])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        summary_dict = dict(zip(results["summary"]["Metric / Indicator"], results["summary"]["Value"]))

        self.assertEqual(summary_dict["Missing in External Financial Exposure (USD)"], "0.00")
        self.assertEqual(summary_dict["Missing in External Rows With Invalid Amount (USD)"], 1)


class TestCanonicalKeyProcessing(unittest.TestCase):
    """Tests for Key Canonicalization across CSV and XLSX Types."""

    def test_canonicalize_key(self):
        self.assertEqual(canonicalize_key(1001.0), "1001")
        self.assertEqual(canonicalize_key(" 1001 "), "1001")
        self.assertEqual(canonicalize_key("00123"), "00123")
        self.assertIsNone(canonicalize_key(None))
        self.assertIsNone(canonicalize_key("   "))


class TestDateParsing(unittest.TestCase):
    """Tests for Date Parsing and Comparison Rules."""

    def test_parse_date_valid_inputs(self):
        self.assertEqual(parse_date("2026-08-01"), pd.Timestamp("2026-08-01"))
        self.assertEqual(parse_date(pd.Timestamp("2026-08-01 14:30:00")), pd.Timestamp("2026-08-01"))
        self.assertEqual(parse_date(datetime(2026, 8, 1, 10, 0)), pd.Timestamp("2026-08-01"))
        self.assertEqual(parse_date(date(2026, 8, 1)), pd.Timestamp("2026-08-01"))

    def test_parse_date_missing_inputs(self):
        self.assertIsNone(parse_date(None))
        self.assertIsNone(parse_date(float("nan")))
        self.assertIsNone(parse_date(""))

    def test_parse_date_invalid_inputs(self):
        self.assertEqual(parse_date("invalid_date"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_date("2026-13-45"), INVALID_VALUE_SENTINEL)

    def test_invalid_date_never_matches(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "customer_name": "Acme", "date": "not-a-date"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "currency": "USD", "customer_name": "Acme", "date": "invalid-date"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["matches"]), 0)
        self.assertEqual(len(results["data_issues"]), 1)
        self.assertEqual(results["data_issues"]["reconciliation_status"].iloc[0], STATUS_INVALID_FIELD_VALUE)


class TestIdentifierPreservation(unittest.TestCase):
    """Tests for Preserving Raw Identifier Strings (e.g. Leading Zeros)."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_csv_preserves_leading_zeros(self):
        csv_path = os.path.join(self.temp_dir, "test_zeros.csv")
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write("invoice_id,amount,currency,date\n00123,100.00,USD,2026-08-01\n000001,200.00,USD,2026-08-01\n")

        df = read_spreadsheet(csv_path, source_name="ERP")
        self.assertEqual(df["invoice_id"].iloc[0], "00123")
        self.assertEqual(df["invoice_id"].iloc[1], "000001")


class TestMonetaryParsing(unittest.TestCase):
    """Comprehensive Monetary Parsing & Strict Syntax Tests."""

    def test_valid_monetary_formats(self):
        self.assertEqual(parse_decimal(1500), Decimal("1500"))
        self.assertEqual(parse_decimal(1500.00), Decimal("1500"))
        self.assertEqual(parse_decimal("$1,500.00"), Decimal("1500.00"))
        self.assertEqual(parse_decimal("₹1,500.00"), Decimal("1500.00"))
        self.assertEqual(parse_decimal("€1,500.00"), Decimal("1500.00"))
        self.assertEqual(parse_decimal("USD 1500.00"), Decimal("1500.00"))
        self.assertEqual(parse_decimal("-1500.00"), Decimal("-1500.00"))
        self.assertEqual(parse_decimal("(500.00)"), Decimal("-500.00"))
        self.assertEqual(parse_decimal("($1,500.00)"), Decimal("-1500.00"))
        self.assertEqual(parse_decimal("1,23,456.78"), Decimal("123456.78"))

    def test_invalid_monetary_formats(self):
        self.assertEqual(parse_decimal("ERROR"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("INVALID"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("ERROR123"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("ABC500XYZ"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("12.3.4"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("1,2,3"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("1,,000"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("USDABC100"), INVALID_VALUE_SENTINEL)

    def test_summarize_money_values_helper(self):
        series = pd.Series(["100.00", "$200.00", "ERROR123", "", None])
        total, inv, blk = summarize_money_values(series)
        self.assertEqual(total, Decimal("300.00"))
        self.assertEqual(inv, 1)
        self.assertEqual(blk, 2)


class TestReconciliationTaxonomyAndDuplicates(unittest.TestCase):
    """Tests for Taxonomy and Duplicate Quarantine."""

    def test_duplicate_isolation_both_sides(self):
        df_erp = pd.DataFrame(
            [
                {"invoice_id": "INV-1007", "amount": "800.00", "currency": "USD", "customer_name": "Zeta", "date": "2026-08-01"},
                {"invoice_id": "INV-1007", "amount": "800.00", "currency": "USD", "customer_name": "Zeta", "date": "2026-08-01"},
            ]
        )
        df_ext = pd.DataFrame(
            [
                {"invoice_id": "INV-1007", "amount": "$800.00", "currency": "USD", "customer_name": "Zeta", "date": "2026-08-01"},
            ]
        )

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["missing"]), 0)
        self.assertEqual(len(results["data_issues"]), 3)
        self.assertTrue((results["data_issues"]["reconciliation_status"] == STATUS_DUPLICATE_IN_ERP).all())


class TestOutputWorkbookAndSecurity(unittest.TestCase):
    """Tests for Output Excel Generation and Security."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_prevent_silent_overwrite(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])
        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")

        filename = "test_output.xlsx"
        write_output(results, self.temp_dir, output_filename=filename, overwrite=True)

        with self.assertRaises(OutputError):
            write_output(results, self.temp_dir, output_filename=filename, overwrite=False)


if __name__ == "__main__":
    unittest.main()
