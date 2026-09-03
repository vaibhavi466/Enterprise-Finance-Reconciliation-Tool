"""
Comprehensive Unit & Integration Test Suite
============================================
Enterprise Finance Reconciliation Tool
"""

import os
import shutil
import tempfile
import unittest
from decimal import Decimal

import pandas as pd
import openpyxl

from reconciliation import (
    INVALID_VALUE_SENTINEL,
    STATUS_DUPLICATE_IN_BOTH,
    STATUS_DUPLICATE_IN_ERP,
    STATUS_DUPLICATE_IN_EXTERNAL,
    STATUS_INVALID_FIELD_VALUE,
    STATUS_MATCH,
    STATUS_MISMATCH,
    STATUS_MISSING_IN_ERP,
    STATUS_MISSING_IN_EXTERNAL,
    STATUS_MISSING_PRIMARY_KEY,
    ConfigurationError,
    IngestionError,
    ReconciliationError,
    SchemaError,
    calculate_file_hash,
    parse_cli_args,
    parse_decimal,
    preprocess_data,
    read_spreadsheet,
    reconcile_data,
    sanitize_column_name,
    sanitize_excel_cell_value,
    write_output,
)


class TestFinancialParsing(unittest.TestCase):
    """1. Tests for Decimal Monetary Parsing & Invalid Value Handling."""

    def test_parse_decimal_valid_formats(self):
        self.assertEqual(parse_decimal(1500), Decimal("1500"))
        self.assertEqual(parse_decimal(1500.00), Decimal("1500"))
        self.assertEqual(parse_decimal("$1,500.00"), Decimal("1500.00"))
        self.assertEqual(parse_decimal("₹1,500.00"), Decimal("1500.00"))
        self.assertEqual(parse_decimal("€1,500.00"), Decimal("1500.00"))
        self.assertEqual(parse_decimal("USD 1500.00"), Decimal("1500.00"))
        self.assertEqual(parse_decimal("-1500.00"), Decimal("-1500.00"))
        self.assertEqual(parse_decimal("(500.00)"), Decimal("-500.00"))
        self.assertEqual(parse_decimal("($1,500.00)"), Decimal("-1500.00"))

    def test_parse_decimal_missing_values(self):
        self.assertIsNone(parse_decimal(None))
        self.assertIsNone(parse_decimal(float("nan")))
        self.assertIsNone(parse_decimal(""))
        self.assertIsNone(parse_decimal("   "))
        self.assertIsNone(parse_decimal("nan"))
        self.assertIsNone(parse_decimal("null"))

    def test_parse_decimal_invalid_values(self):
        self.assertEqual(parse_decimal("ERROR"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("INVALID"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("ERROR123"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("ABC500XYZ"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("12.3.4"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_decimal("non_financial_text"), INVALID_VALUE_SENTINEL)


class TestColumnSanitizationAndSchema(unittest.TestCase):
    """2. Tests for Header Sanitization and Collision Errors."""

    def test_sanitize_column_name(self):
        self.assertEqual(sanitize_column_name(" Invoice ID "), "invoice_id")
        self.assertEqual(sanitize_column_name("Invoice-ID!"), "invoice_id")
        self.assertEqual(sanitize_column_name("Total  Amount ($)"), "total_amount")
        self.assertEqual(sanitize_column_name("Customer...Name"), "customer_name")

    def test_header_collision_error(self):
        df_colliding = pd.DataFrame(columns=["Invoice ID", "invoice_id", "amount"])
        with self.assertRaises(SchemaError) as ctx:
            preprocess_data(df_colliding, source_name="TestERP")
        self.assertIn("Header normalization collision", str(ctx.exception))


class TestPrimaryKeysAndDuplicates(unittest.TestCase):
    """3. Tests for Null/Blank Key Quarantine & Duplicate Quarantine Edge Cases."""

    def test_null_and_blank_primary_keys(self):
        df_erp = pd.DataFrame([
            {"invoice_id": None, "amount": 100.00, "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "   ", "amount": 200.00, "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "nan", "amount": 300.00, "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "INV-100", "amount": 400.00, "currency": "USD", "date": "2026-08-01"},
        ])
        df_ext = pd.DataFrame([
            {"invoice_id": "INV-100", "amount": "$400.00", "currency": "USD", "date": "2026-08-01"},
        ])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        data_issues = results["data_issues"]
        self.assertEqual(len(data_issues), 3)
        self.assertTrue((data_issues["reconciliation_status"] == STATUS_MISSING_PRIMARY_KEY).all())

    def test_duplicate_isolation_edge_case(self):
        # ERP: INV-1007, INV-1007
        # External: INV-1007 (single row)
        # External row MUST NOT be marked as MISSING_IN_ERP!
        df_erp = pd.DataFrame([
            {"invoice_id": "INV-1007", "amount": 800.00, "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "INV-1007", "amount": 800.00, "currency": "USD", "date": "2026-08-01"},
        ])
        df_ext = pd.DataFrame([
            {"invoice_id": "INV-1007", "amount": "$800.00", "currency": "USD", "date": "2026-08-01"},
        ])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["missing"]), 0)
        self.assertEqual(len(results["data_issues"]), 3)
        self.assertTrue((results["data_issues"]["reconciliation_status"] == STATUS_DUPLICATE_IN_ERP).all())

    def test_duplicate_in_both_sides(self):
        df_erp = pd.DataFrame([
            {"invoice_id": "INV-200", "amount": 500.00, "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "INV-200", "amount": 500.00, "currency": "USD", "date": "2026-08-01"},
        ])
        df_ext = pd.DataFrame([
            {"invoice_id": "INV-200", "amount": "$500.00", "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "INV-200", "amount": "$500.00", "currency": "USD", "date": "2026-08-01"},
        ])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["data_issues"]), 4)
        self.assertTrue((results["data_issues"]["reconciliation_status"] == STATUS_DUPLICATE_IN_BOTH).all())


class TestCoreReconciliationLogic(unittest.TestCase):
    """4. Tests for Core Happy Path, Mismatches, Missing, and Invalid Values."""

    def setUp(self):
        self.df_erp = pd.DataFrame([
            {"invoice_id": "INV-101", "amount": 1000.00, "currency": "USD", "date": "2026-08-01"},  # Exact Match
            {"invoice_id": "INV-102", "amount": 2000.00, "currency": "USD", "date": "2026-08-01"},  # Amount Tolerance Match
            {"invoice_id": "INV-103", "amount": 3000.00, "currency": "USD", "date": "2026-08-01"},  # Amount Mismatch
            {"invoice_id": "INV-104", "amount": 4000.00, "currency": "USD", "date": "2026-08-01"},  # Currency Mismatch
            {"invoice_id": "INV-105", "amount": 5000.00, "currency": "USD", "date": "2026-08-01"},  # Missing in External
            {"invoice_id": "INV-107", "amount": 7000.00, "currency": "USD", "date": "2026-08-01"},  # Invalid Amount in External
        ])

        self.df_ext = pd.DataFrame([
            {"invoice_id": "INV-101", "amount": "$1,000.00", "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "INV-102", "amount": "$2,000.004", "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "INV-103", "amount": "$3,050.00", "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "INV-104", "amount": "$4,000.00", "currency": "EUR", "date": "2026-08-01"},
            {"invoice_id": "INV-106", "amount": "$6,000.00", "currency": "USD", "date": "2026-08-01"},  # Missing in ERP
            {"invoice_id": "INV-107", "amount": "ERROR123", "currency": "USD", "date": "2026-08-01"},
        ])

    def test_reconciliation_categories(self):
        results = reconcile_data(self.df_erp, self.df_ext, key_column="invoice_id")

        # Matches: INV-101, INV-102
        self.assertEqual(len(results["matches"]), 2)
        matched_keys = set(results["matches"]["invoice_id"])
        self.assertEqual(matched_keys, {"INV-101", "INV-102"})

        # Mismatches: INV-103, INV-104
        self.assertEqual(len(results["mismatches"]), 2)
        mismatched_keys = set(results["mismatches"]["invoice_id"])
        self.assertEqual(mismatched_keys, {"INV-103", "INV-104"})

        # Missing: INV-105 (in External), INV-106 (in ERP)
        self.assertEqual(len(results["missing"]), 2)

        # Data Issues: INV-107 (INVALID_FIELD_VALUE)
        self.assertEqual(len(results["data_issues"]), 1)
        self.assertEqual(results["data_issues"]["reconciliation_status"].iloc[0], STATUS_INVALID_FIELD_VALUE)

        # All Records consolidated check
        self.assertEqual(len(results["all_records"]), 7)

    def test_invalid_monetary_value_never_matches(self):
        # ERP: ERROR, External: INVALID -> must NOT become a match
        df_erp = pd.DataFrame([{"invoice_id": "INV-999", "amount": "ERROR", "currency": "USD", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-999", "amount": "INVALID", "currency": "USD", "date": "2026-08-01"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["matches"]), 0)
        self.assertEqual(len(results["data_issues"]), 1)
        self.assertEqual(results["data_issues"]["reconciliation_status"].iloc[0], STATUS_INVALID_FIELD_VALUE)

    def test_missing_required_field_schema_error(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": 100.00}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": 100.00}])
        # Config requires 'currency' which is missing from both
        custom_rules = {
            "amount": {"type": "money", "tolerance": Decimal("0.01")},
            "currency": {"type": "text"},
        }
        with self.assertRaises(SchemaError):
            reconcile_data(df_erp, df_ext, key_column="invoice_id", field_rules=custom_rules)


class TestOutputAndFormatting(unittest.TestCase):
    """5. Tests for Excel Output Generation, Styling, and Formula Injection Protection."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_excel_formula_injection_sanitization(self):
        self.assertEqual(sanitize_excel_cell_value("=SUM(A1:A10)"), "'=SUM(A1:A10)")
        self.assertEqual(sanitize_excel_cell_value("+12345"), "'+12345")
        self.assertEqual(sanitize_excel_cell_value("@command"), "'@command")
        self.assertEqual(sanitize_excel_cell_value("-500.00"), "-500.00")  # valid negative numeric kept
        self.assertEqual(sanitize_excel_cell_value("-cmd"), "'-cmd")

    def test_excel_sheet_order_and_styling(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": 100.00, "currency": "USD", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "currency": "USD", "date": "2026-08-01"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        output_path = write_output(results, self.temp_dir)

        self.assertTrue(os.path.exists(output_path))
        wb = openpyxl.load_workbook(output_path)

        expected_sheets = ["summary", "run_metadata", "matches", "mismatches", "missing", "data_issues", "all_records"]
        self.assertEqual(wb.sheetnames, expected_sheets)


class TestEmptyAndEdgeCases(unittest.TestCase):
    """6. Tests for Configuration Errors and File Hashing."""

    def test_negative_tolerance_error(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": 100.00, "currency": "USD", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "currency": "USD", "date": "2026-08-01"}])

        with self.assertRaises(ConfigurationError):
            reconcile_data(df_erp, df_ext, key_column="invoice_id", amount_tolerance=-0.05)

    def test_unsupported_file_extension(self):
        with self.assertRaises(IngestionError):
            read_spreadsheet("invalid_file.doc")


if __name__ == "__main__":
    unittest.main()
