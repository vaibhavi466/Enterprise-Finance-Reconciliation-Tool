"""
Sample Data Generator for Enterprise Finance Reconciliation Tool
==================================================================
Produces deterministic sample financial datasets (CSV and XLSX) in sample_data/
demonstrating all reconciliation taxonomy outcomes:
- Exact Match (MATCH)
- Amount Tolerance Match (MATCH)
- Amount Discrepancy (MISMATCH)
- Currency Discrepancy (MISMATCH)
- Missing in External (MISSING_IN_EXTERNAL)
- Missing in ERP (MISSING_IN_ERP)
- Duplicate ERP Primary Key (DUPLICATE_KEY_IN_ERP)
- Duplicate External Primary Key (DUPLICATE_KEY_IN_EXTERNAL)
- Missing Primary Key (MISSING_PRIMARY_KEY)
- Invalid Monetary Format (INVALID_FIELD_VALUE)
"""

import os

import pandas as pd


def main():
    sample_dir = "sample_data"
    os.makedirs(sample_dir, exist_ok=True)

    erp_rows = [
        # 1. Exact Match
        {
            "invoice_id": "INV-1001",
            "amount": 1500.00,
            "currency": "USD",
            "customer_name": "Acme Corp",
            "date": "2026-08-01",
        },
        # 2. Tolerance Match (within ±0.01 threshold)
        {
            "invoice_id": "INV-1002",
            "amount": 2500.00,
            "currency": "USD",
            "customer_name": "Beta LLC",
            "date": "2026-08-02",
        },
        # 3. Amount Mismatch
        {
            "invoice_id": "INV-1003",
            "amount": 3200.00,
            "currency": "USD",
            "customer_name": "Gamma Inc",
            "date": "2026-08-03",
        },
        # 4. Currency Mismatch
        {
            "invoice_id": "INV-1004",
            "amount": 4500.00,
            "currency": "USD",
            "customer_name": "Delta Co",
            "date": "2026-08-04",
        },
        # 5. Missing in External
        {
            "invoice_id": "INV-1005",
            "amount": 1200.00,
            "currency": "EUR",
            "customer_name": "Epsilon Ltd",
            "date": "2026-08-05",
        },
        # 7. Duplicate Key in ERP (Row 1 & Row 2)
        {
            "invoice_id": "INV-1007",
            "amount": 800.00,
            "currency": "USD",
            "customer_name": "Zeta Enterprises",
            "date": "2026-08-07",
        },
        {
            "invoice_id": "INV-1007",
            "amount": 800.00,
            "currency": "USD",
            "customer_name": "Zeta Enterprises",
            "date": "2026-08-07",
        },
        # 8. Single ERP record corresponding to Duplicate Key in External
        {
            "invoice_id": "INV-1008",
            "amount": 950.00,
            "currency": "USD",
            "customer_name": "Theta Co",
            "date": "2026-08-08",
        },
        # 9. Missing Primary Key
        {"invoice_id": "   ", "amount": 100.00, "currency": "USD", "customer_name": "Null Corp", "date": "2026-08-09"},
        # 10. Invalid Monetary Format in External
        {
            "invoice_id": "INV-1010",
            "amount": 600.00,
            "currency": "USD",
            "customer_name": "Iota LLC",
            "date": "2026-08-10",
        },
    ]

    bank_rows = [
        # 1. Exact Match
        {
            "invoice_id": "INV-1001",
            "amount": "$1,500.00",
            "currency": "USD",
            "customer_name": "Acme Corp",
            "date": "2026-08-01",
        },
        # 2. Tolerance Match ($2,500.004 vs 2500.00 -> diff 0.004)
        {
            "invoice_id": "INV-1002",
            "amount": "$2,500.004",
            "currency": "USD",
            "customer_name": "Beta LLC",
            "date": "2026-08-02",
        },
        # 3. Amount Mismatch ($3,250.00 vs 3200.00)
        {
            "invoice_id": "INV-1003",
            "amount": "$3,250.00",
            "currency": "USD",
            "customer_name": "Gamma Inc",
            "date": "2026-08-03",
        },
        # 4. Currency Mismatch (EUR vs USD)
        {
            "invoice_id": "INV-1004",
            "amount": "$4,500.00",
            "currency": "EUR",
            "customer_name": "Delta Co",
            "date": "2026-08-04",
        },
        # 6. Missing in ERP
        {
            "invoice_id": "INV-1006",
            "amount": "$5,000.00",
            "currency": "USD",
            "customer_name": "Eta Corp",
            "date": "2026-08-06",
        },
        # 7. Single External record corresponding to Duplicate Key in ERP
        {
            "invoice_id": "INV-1007",
            "amount": "$800.00",
            "currency": "USD",
            "customer_name": "Zeta Enterprises",
            "date": "2026-08-07",
        },
        # 8. Duplicate Key in External (Row 1 & Row 2)
        {
            "invoice_id": "INV-1008",
            "amount": "$950.00",
            "currency": "USD",
            "customer_name": "Theta Co",
            "date": "2026-08-08",
        },
        {
            "invoice_id": "INV-1008",
            "amount": "$950.00",
            "currency": "USD",
            "customer_name": "Theta Co",
            "date": "2026-08-08",
        },
        # 10. Invalid Monetary Format in External
        {
            "invoice_id": "INV-1010",
            "amount": "ERROR123",
            "currency": "USD",
            "customer_name": "Iota LLC",
            "date": "2026-08-10",
        },
    ]

    df_erp = pd.DataFrame(erp_rows)
    df_bank = pd.DataFrame(bank_rows)

    erp_csv_path = os.path.join(sample_dir, "erp_data.csv")
    bank_xlsx_path = os.path.join(sample_dir, "bank_data.xlsx")

    df_erp.to_csv(erp_csv_path, index=False)
    df_bank.to_excel(bank_xlsx_path, index=False, engine="openpyxl")

    print("[SUCCESS] Deterministic sample datasets created successfully:")
    print(f"   - ERP Data: {erp_csv_path} ({len(df_erp)} rows)")
    print(f"   - External Data: {bank_xlsx_path} ({len(df_bank)} rows)")


if __name__ == "__main__":
    main()
