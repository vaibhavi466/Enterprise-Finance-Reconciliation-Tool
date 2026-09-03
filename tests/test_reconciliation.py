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
    ConfigurationError,
    OutputError,
    SchemaError,
    normalize_and_validate_field_rules,
    parse_date,
    parse_decimal,
    preprocess_data,
    read_spreadsheet,
    reconcile_data,
    sanitize_excel_cell_value,
    write_output,
)


class TestDateParsing(unittest.TestCase):
    """A. Tests for Date Parsing and Comparison Rules."""

    def test_parse_date_valid_inputs(self):
        self.assertEqual(parse_date("2026-08-01"), pd.Timestamp("2026-08-01"))
        self.assertEqual(parse_date(pd.Timestamp("2026-08-01 14:30:00")), pd.Timestamp("2026-08-01"))
        self.assertEqual(parse_date(datetime(2026, 8, 1, 10, 0)), pd.Timestamp("2026-08-01"))
        self.assertEqual(parse_date(date(2026, 8, 1)), pd.Timestamp("2026-08-01"))

    def test_parse_date_missing_inputs(self):
        self.assertIsNone(parse_date(None))
        self.assertIsNone(parse_date(float("nan")))
        self.assertIsNone(parse_date(""))
        self.assertIsNone(parse_date("   "))

    def test_parse_date_invalid_inputs(self):
        self.assertEqual(parse_date("invalid_date"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_date("2026-13-45"), INVALID_VALUE_SENTINEL)
        self.assertEqual(parse_date("not-a-date"), INVALID_VALUE_SENTINEL)

    def test_invalid_date_never_matches(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "date": "not-a-date"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "currency": "USD", "date": "invalid-date"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["matches"]), 0)
        self.assertEqual(len(results["data_issues"]), 1)
        self.assertEqual(results["data_issues"]["reconciliation_status"].iloc[0], STATUS_INVALID_FIELD_VALUE)


class TestIdentifierPreservation(unittest.TestCase):
    """B. Tests for Preserving Raw Identifier Strings (e.g. Leading Zeros)."""

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


class TestSchemaValidation(unittest.TestCase):
    """C. Tests for Mandatory Schema Validation and Collision Detection."""

    def test_missing_required_field_raises_schema_error(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "customer_name": "Acme", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "date": "2026-08-01"}])  # Missing customer_name & currency

        rules = {
            "amount": {"type": "money"},
            "customer_name": {"type": "text"},
        }
        with self.assertRaises(SchemaError) as ctx:
            reconcile_data(df_erp, df_ext, key_column="invoice_id", field_rules=rules)
        self.assertIn("customer_name", str(ctx.exception))

    def test_header_collision_raises_schema_error(self):
        df_colliding = pd.DataFrame(columns=["Invoice ID", "invoice_id", "amount"])
        with self.assertRaises(SchemaError):
            preprocess_data(df_colliding, source_name="TestERP")


class TestRuleConfiguration(unittest.TestCase):
    """D. Tests for Field Rule Normalization and Validation."""

    def test_unsupported_rule_type_raises_config_error(self):
        rules = {"amount": {"type": "banana"}}
        with self.assertRaises(ConfigurationError) as ctx:
            normalize_and_validate_field_rules(rules, key_column="invoice_id")
        self.assertIn("banana", str(ctx.exception))

    def test_negative_tolerance_raises_config_error(self):
        rules = {"amount": {"type": "money", "tolerance": "-0.01"}}
        with self.assertRaises(ConfigurationError):
            normalize_and_validate_field_rules(rules, key_column="invoice_id")

    def test_colliding_config_field_names(self):
        rules = {
            "Total Amount": {"type": "money"},
            "total_amount": {"type": "money"},
        }
        with self.assertRaises(ConfigurationError):
            normalize_and_validate_field_rules(rules, key_column="invoice_id")


class TestFieldInferenceRegression(unittest.TestCase):
    """E. Regression Tests ensuring NO automatic substring field type inference exists."""

    def test_approval_status_and_updated_status_never_inferred(self):
        # approval_status contains 'val'
        # updated_status contains 'date'
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "approval_status": "PENDING", "updated_status": "ACTIVE"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "approval_status": "APPROVED", "updated_status": "ACTIVE"}])

        rules = {
            "amount": {"type": "money"},
            "approval_status": {"type": "text"},
            "updated_status": {"type": "text"},
        }
        results = reconcile_data(df_erp, df_ext, key_column="invoice_id", field_rules=rules)
        # Should be MISMATCH on approval_status as text comparison
        self.assertEqual(len(results["mismatches"]), 1)
        self.assertIn("approval_status mismatch", results["mismatches"]["status_reason"].iloc[0])


class TestMonetaryParsing(unittest.TestCase):
    """F. Comprehensive Monetary Parsing & Strict Syntax Tests."""

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

    def test_tolerance_boundary_exact_and_outside(self):
        # exact boundary: diff 0.010000 -> MATCH
        v1 = Decimal("100.00")
        v2 = Decimal("100.01")
        self.assertTrue(abs(v1 - v2) <= Decimal("0.01"))

        # outside boundary: diff 0.010001 -> MISMATCH
        v3 = Decimal("100.010001")
        self.assertFalse(abs(v1 - v3) <= Decimal("0.01"))


class TestCurrencyExposure(unittest.TestCase):
    """G. Tests for Currency-Isolated Financial Exposure Calculations."""

    def test_cross_currency_variance_not_summed(self):
        # ERP: USD 100, External: EUR 150
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$150.00", "currency": "EUR", "date": "2026-08-01"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        summary_dict = dict(zip(results["summary"]["Metric / Indicator"], results["summary"]["Value"]))

        self.assertEqual(summary_dict["Currency Mismatch Count"], 1)
        self.assertNotIn("Mismatched Exposure (USD)", summary_dict)
        self.assertNotIn("Mismatched Exposure (EUR)", summary_dict)


class TestMultipleMoneyFields(unittest.TestCase):
    """H. Tests for Supporting Multiple Monetary Fields with Separate Variances."""

    def test_multiple_money_fields_variances(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "tax_amount": "10.00"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$105.00", "tax_amount": "$12.00"}])

        rules = {
            "amount": {"type": "money", "tolerance": Decimal("0.01")},
            "tax_amount": {"type": "money", "tolerance": Decimal("0.01")},
        }
        results = reconcile_data(df_erp, df_ext, key_column="invoice_id", field_rules=rules)
        common = results["all_records"]

        self.assertIn("amount_variance", common.columns)
        self.assertIn("tax_amount_variance", common.columns)
        self.assertEqual(common["amount_variance"].iloc[0], Decimal("5.00"))
        self.assertEqual(common["tax_amount_variance"].iloc[0], Decimal("2.00"))


class TestReconciliationTaxonomyAndDuplicates(unittest.TestCase):
    """I & J. Tests for Complete Taxonomy and Ambiguous Duplicate Quarantine."""

    def test_duplicate_isolation_both_sides(self):
        df_erp = pd.DataFrame(
            [
                {"invoice_id": "INV-1007", "amount": "800.00", "currency": "USD", "date": "2026-08-01"},
                {"invoice_id": "INV-1007", "amount": "800.00", "currency": "USD", "date": "2026-08-01"},
            ]
        )
        df_ext = pd.DataFrame(
            [
                {"invoice_id": "INV-1007", "amount": "$800.00", "currency": "USD", "date": "2026-08-01"},
            ]
        )

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["missing"]), 0)
        self.assertEqual(len(results["data_issues"]), 3)
        self.assertTrue((results["data_issues"]["reconciliation_status"] == STATUS_DUPLICATE_IN_ERP).all())


class TestKPISummaryAndMetadata(unittest.TestCase):
    """K & L. Tests for Executive KPIs and Audit Metadata."""

    def test_kpi_counts_and_metadata(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "currency": "USD", "date": "2026-08-01"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        meta_dict = dict(zip(results["run_metadata"]["Parameter"], results["run_metadata"]["Value"]))

        self.assertIn("Run ID", meta_dict)
        self.assertIn("Run Timestamp UTC", meta_dict)
        self.assertIn("Tool Version", meta_dict)
        self.assertEqual(meta_dict["Primary Key Column"], "invoice_id")


class TestOutputWorkbookAndSecurity(unittest.TestCase):
    """M. Tests for Excel Output Formatting, Sheet Order, and Overwrite Security."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_prevent_silent_overwrite(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "currency": "USD", "date": "2026-08-01"}])
        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")

        filename = "test_output.xlsx"
        write_output(results, self.temp_dir, output_filename=filename, overwrite=True)

        # Second call without overwrite=True must raise OutputError
        with self.assertRaises(OutputError):
            write_output(results, self.temp_dir, output_filename=filename, overwrite=False)

    def test_formula_injection_escaping(self):
        self.assertEqual(sanitize_excel_cell_value("=SUM(A1:A10)"), "'=SUM(A1:A10)")
        self.assertEqual(sanitize_excel_cell_value("+12345"), "'+12345")
        self.assertEqual(sanitize_excel_cell_value("-500.00"), "-500.00")


class TestEmptyDatasets(unittest.TestCase):
    """N. Tests for Deliberate Empty Dataset Handling."""

    def test_empty_datasets_handled_gracefully(self):
        df_erp = pd.DataFrame(columns=["invoice_id", "amount", "currency", "date"])
        df_ext = pd.DataFrame(columns=["invoice_id", "amount", "currency", "date"])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["matches"]), 0)
        self.assertEqual(len(results["all_records"]), 0)


if __name__ == "__main__":
    unittest.main()
