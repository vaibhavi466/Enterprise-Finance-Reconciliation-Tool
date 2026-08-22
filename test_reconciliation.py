"""
Unit & Integration Test Suite for Enterprise Finance Reconciliation Tool
========================================================================
"""

import os
import unittest
import pandas as pd
from reconciliation import (
    preprocess_data,
    parse_numeric,
    reconcile_data,
    read_spreadsheet,
    write_output,
)


class TestReconciliationTool(unittest.TestCase):

    def setUp(self):
        self.output_dir = "test_output"
        os.makedirs(self.output_dir, exist_ok=True)

        self.df_erp = pd.DataFrame([
            {"invoice_id": "INV-101", "amount": 1000.00, "currency": "USD", "date": "2026-08-01"},  # Match
            {"invoice_id": "INV-102", "amount": 2000.00, "currency": "USD", "date": "2026-08-01"},  # Mismatch (Amount)
            {"invoice_id": "INV-103", "amount": 3000.00, "currency": "USD", "date": "2026-08-01"},  # Mismatch (Currency)
            {"invoice_id": "INV-104", "amount": 4000.00, "currency": "USD", "date": "2026-08-01"},  # Missing in Bank
            {"invoice_id": "INV-106", "amount": 6000.00, "currency": "USD", "date": "2026-08-01"},  # Duplicate ERP
            {"invoice_id": "INV-106", "amount": 6000.00, "currency": "USD", "date": "2026-08-01"},  # Duplicate ERP
        ])

        self.df_bank = pd.DataFrame([
            {"invoice_id": "INV-101", "amount": "$1,000.00", "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "INV-102", "amount": "$2,050.00", "currency": "USD", "date": "2026-08-01"},
            {"invoice_id": "INV-103", "amount": "$3,000.00", "currency": "EUR", "date": "2026-08-01"},
            {"invoice_id": "INV-105", "amount": "$5,000.00", "currency": "USD", "date": "2026-08-01"},  # Missing in ERP
        ])

    def test_parse_numeric(self):
        self.assertEqual(parse_numeric("$1,250.50"), 1250.50)
        self.assertEqual(parse_numeric(500), 500.0)
        self.assertEqual(parse_numeric("invalid"), None)

    def test_preprocessing(self):
        raw_df = pd.DataFrame({" Invoice ID ": [" INV-001 "], " Total Amount ": [" 500 "]})
        cleaned = preprocess_data(raw_df)
        self.assertIn("invoice_id", cleaned.columns)
        self.assertIn("total_amount", cleaned.columns)
        self.assertEqual(cleaned["invoice_id"].iloc[0], "INV-001")

    def test_reconciliation_logic(self):
        erp_clean = preprocess_data(self.df_erp)
        bank_clean = preprocess_data(self.df_bank)

        results = reconcile_data(erp_clean, bank_clean, key_column="invoice_id")

        # 1. Matches: INV-101
        self.assertEqual(len(results["matches"]), 1)
        self.assertEqual(results["matches"]["invoice_id"].iloc[0], "INV-101")

        # 2. Mismatches: INV-102 (Amount diff), INV-103 (Currency diff)
        self.assertEqual(len(results["mismatches"]), 2)

        # 3. Missing: INV-104 (Bank), INV-105 (ERP)
        self.assertEqual(len(results["missing"]), 2)

        # 4. Data Issues: INV-106 (Duplicates)
        self.assertEqual(len(results["data_issues"]), 2)

    def test_excel_export(self):
        erp_clean = preprocess_data(self.df_erp)
        bank_clean = preprocess_data(self.df_bank)
        results = reconcile_data(erp_clean, bank_clean)

        output_path = write_output(results, self.output_dir)
        self.assertTrue(os.path.exists(output_path))


if __name__ == "__main__":
    unittest.main()
