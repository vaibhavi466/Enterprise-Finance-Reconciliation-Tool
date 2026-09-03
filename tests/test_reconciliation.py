"""
Comprehensive Unit & Integration Test Suite
============================================
Enterprise Finance Reconciliation Tool (C1-C8 Fixes)
"""

import os
import shutil
import tempfile
import unittest
from decimal import Decimal

import openpyxl
import pandas as pd

from reconciliation import (
    DQ_CURRENCY_AMOUNT_CONFLICT,
    DQ_DUPLICATE_IN_ERP,
    DQ_INVALID_FIELD_VALUE,
    DQ_MISSING_REQUIRED_VALUE,
    STATUS_MISSING_IN_EXTERNAL,
    normalize_text,
    parse_money,
    read_spreadsheet,
    reconcile_data,
)


class TestC1XlsxLeadingZeros(unittest.TestCase):
    """C1: Verify Excel leading-zero string IDs are preserved."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_xlsx_leading_zeros_preserved(self):
        xlsx_path = os.path.join(self.temp_dir, "test_zeros.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["invoice_id", "amount", "currency", "date"])
        ws.cell(row=2, column=1, value="00123")  # String text cell with leading zero
        ws.cell(row=2, column=2, value="100.00")
        ws.cell(row=2, column=3, value="USD")
        ws.cell(row=2, column=4, value="2026-08-01")
        wb.save(xlsx_path)

        df = read_spreadsheet(xlsx_path, source_name="External")
        self.assertEqual(df["invoice_id"].iloc[0], "00123")


class TestC2AndC6RequiredTextNormalization(unittest.TestCase):
    """C2 & C6: Verify text normalization and consistent required field semantics."""

    def test_normalize_text_helper(self):
        self.assertIsNone(normalize_text(None))
        self.assertIsNone(normalize_text(""))
        self.assertIsNone(normalize_text("   "))
        self.assertIsNone(normalize_text("null"))
        self.assertIsNone(normalize_text("none"))
        self.assertIsNone(normalize_text("<na>"))
        self.assertEqual(normalize_text(" USD "), "USD")

    def test_required_text_field_missing_either_side(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "currency": "", "customer_name": "Acme", "date": "2026-08-01"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["matches"]), 0)
        self.assertEqual(len(results["mismatches"]), 1)
        self.assertEqual(results["mismatches"]["data_quality_status"].iloc[0], DQ_MISSING_REQUIRED_VALUE)


class TestC3EmbeddedCurrencyConflict(unittest.TestCase):
    """C3: Verify money parsing detects embedded currency symbol/code conflicts."""

    def test_parse_money_structured_result(self):
        pm1 = parse_money("€100.00")
        self.assertEqual(pm1.amount, Decimal("100.00"))
        self.assertEqual(pm1.explicit_currency, "EUR")
        self.assertTrue(pm1.valid)

        pm2 = parse_money("USD 250.50")
        self.assertEqual(pm2.amount, Decimal("250.50"))
        self.assertEqual(pm2.explicit_currency, "USD")
        self.assertTrue(pm2.valid)

    def test_currency_amount_conflict_flagged(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "€100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        self.assertEqual(len(results["matches"]), 0)
        self.assertEqual(results["mismatches"]["data_quality_status"].iloc[0], DQ_CURRENCY_AMOUNT_CONFLICT)


class TestC4SeparatedOutputGrains(unittest.TestCase):
    """C4: Verify output workbook separates reconciliation_units and source_exceptions."""

    def test_separated_output_grains(self):
        df_erp = pd.DataFrame(
            [
                {"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"},
                {"invoice_id": "INV-2", "amount": "200.00", "currency": "USD", "customer_name": "Beta", "date": "2026-08-01"},
                {"invoice_id": "INV-2", "amount": "200.00", "currency": "USD", "customer_name": "Beta", "date": "2026-08-01"},
            ]
        )
        df_ext = pd.DataFrame(
            [
                {"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"},
            ]
        )

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")

        self.assertIn("reconciliation_units", results)
        self.assertIn("source_exceptions", results)

        # 1 unit for INV-1 (MATCH)
        self.assertEqual(len(results["reconciliation_units"]), 1)
        # 2 raw duplicate rows for INV-2 in source_exceptions
        self.assertEqual(len(results["source_exceptions"]), 2)
        self.assertEqual(results["source_exceptions"]["data_quality_status"].iloc[0], DQ_DUPLICATE_IN_ERP)


class TestC5DisambiguatedKPIs(unittest.TestCase):
    """C5: Verify exception-rate metrics are mathematically disambiguated."""

    def test_disambiguated_kpi_labels(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "$100.00", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        summary_dict = dict(zip(results["summary"]["Metric / Indicator"], results["summary"]["Value"]))

        self.assertIn("Source Row Exception Rate (%)", summary_dict)
        self.assertIn("Reconciliation Unit Exception Rate (%)", summary_dict)
        self.assertEqual(summary_dict["Source Row Exception Rate (%)"], "0.00%")
        self.assertEqual(summary_dict["Reconciliation Unit Exception Rate (%)"], "0.00%")


class TestC7DynamicExposure(unittest.TestCase):
    """C7: Verify dynamic per-money-field exposure calculation."""

    def test_dynamic_exposure_multiple_money_fields(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-1", "amount": "100.00", "fee": "10.00", "currency": "USD", "date": "2026-08-01"}])
        df_ext = pd.DataFrame([{"invoice_id": "INV-1", "amount": "150.00", "fee": "20.00", "currency": "USD", "date": "2026-08-01"}])

        rules = {
            "amount": {"type": "money", "tolerance": "0.01", "required": True, "currency_field": "currency"},
            "fee": {"type": "money", "tolerance": "0.01", "required": False, "currency_field": "currency"},
            "currency": {"type": "text", "case_sensitive": False, "required": True},
            "date": {"type": "date", "format": "%Y-%m-%d", "required": True},
        }

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id", field_rules=rules)
        summary_dict = dict(zip(results["summary"]["Metric / Indicator"], results["summary"]["Value"]))

        self.assertIn("Mismatched Amount Exposure (USD)", summary_dict)
        self.assertIn("Mismatched Fee Exposure (USD)", summary_dict)
        self.assertEqual(summary_dict["Mismatched Amount Exposure (USD)"], "50.00")
        self.assertEqual(summary_dict["Mismatched Fee Exposure (USD)"], "10.00")


class TestC8SourceLevelValidation(unittest.TestCase):
    """C8: Verify source-level data quality validation on missing records."""

    def test_source_level_validation_on_missing_record(self):
        df_erp = pd.DataFrame([{"invoice_id": "INV-500", "amount": "ERROR123", "currency": "USD", "customer_name": "Acme", "date": "2026-08-01"}])
        df_ext = pd.DataFrame(columns=["invoice_id", "amount", "currency", "customer_name", "date"])

        results = reconcile_data(df_erp, df_ext, key_column="invoice_id")
        missing_df = results["missing"]

        self.assertEqual(missing_df["reconciliation_status"].iloc[0], STATUS_MISSING_IN_EXTERNAL)
        self.assertEqual(missing_df["data_quality_status"].iloc[0], DQ_INVALID_FIELD_VALUE)


if __name__ == "__main__":
    unittest.main()
