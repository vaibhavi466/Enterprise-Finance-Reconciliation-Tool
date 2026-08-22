"""
Enterprise Finance Reconciliation Script
======================================

Purpose:
--------
This script performs a production-style financial reconciliation between:
1. A System of Record file (ERP / Enterprise data)
2. An External file (Bank / Third-party output)

It classifies records into:
✅ Match
⚠️ Mismatch (Amount / Currency / Date)
❌ Missing
🔁 Data Issues (Duplicates)

The output is a business-friendly Excel file with multiple sheets.

Author:
-------
Finance Automation – Intern Learning Exercise
"""

# ==============================
# Imports
# ==============================

import pandas as pd
import os
from datetime import datetime

# ==============================
# Configuration
# ==============================

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls"}

# Allow small rounding differences in amount comparison
AMOUNT_TOLERANCE = 0.01

# ==============================
# Utility Functions
# ==============================

def read_spreadsheet(file_path):
    """
    Reads a CSV or Excel file into a pandas DataFrame.
    Handles file paths pasted with quotes (Windows CMD safe).
    """

    # Remove quotes if user pasted path with them
    file_path = file_path.strip().strip('"').strip("'")

    _, ext = os.path.splitext(file_path)
    ext = ext.lower()

    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Unsupported file type: {ext}")

    if ext == ".csv":
        return pd.read_csv(file_path)
    else:
        return pd.read_excel(file_path, engine="openpyxl")


def preprocess_data(df):
    """
    Cleans and standardises the dataframe so both systems
    can be compared fairly.
    """

    # Standardise column names
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_", regex=False)
    )

    # Drop completely empty rows
    df = df.dropna(how="all")

    # Trim whitespace from all string cells
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    return df

# ==============================
# Core Reconciliation Logic
# ==============================

KEY_COLUMN = "invoice_id"


def reconcile_data(df_erp, df_bank):
    results = {}

    erp_cols = list(df_erp.columns)
    bank_cols = list(df_bank.columns)

    # Columns to compare (shared between both files, excluding the key)
    compare_cols = [c for c in erp_cols if c in bank_cols and c != KEY_COLUMN]

    # ------------------------------
    # Detect duplicates (Data Issues)
    # ------------------------------
    duplicates_erp = df_erp[df_erp.duplicated(subset=[KEY_COLUMN], keep=False)]
    duplicates_bank = df_bank[df_bank.duplicated(subset=[KEY_COLUMN], keep=False)]

    results["data_issues"] = pd.concat(
        [
            duplicates_erp.assign(source="ERP"),
            duplicates_bank.assign(source="BANK"),
        ],
        ignore_index=True,
    )

    # Remove duplicates before reconciliation
    df_erp = df_erp.drop_duplicates(subset=[KEY_COLUMN])
    df_bank = df_bank.drop_duplicates(subset=[KEY_COLUMN])

    # ------------------------------
    # Merge datasets on key column
    # ------------------------------
    merged = df_erp.merge(
        df_bank,
        on=KEY_COLUMN,
        how="outer",
        suffixes=("_erp", "_bank"),
        indicator=True,
    )

    # ------------------------------
    # Missing records
    # ------------------------------
    results["missing"] = merged[merged["_merge"] != "both"].copy()

    # ------------------------------
    # Records present in both systems
    # ------------------------------
    common = merged[merged["_merge"] == "both"].copy()

    # ------------------------------
    # Dynamic field-level comparisons for all shared columns
    # ------------------------------
    match_flags = {}
    for col in compare_cols:
        erp_col = f"{col}_erp"
        bank_col = f"{col}_bank"
        if erp_col not in common.columns or bank_col not in common.columns:
            continue
        try:
            match_flags[col] = (
                (common[erp_col].astype(float) - common[bank_col].astype(float)).abs()
                <= AMOUNT_TOLERANCE
            )
        except (ValueError, TypeError):
            match_flags[col] = (
                common[erp_col].astype(str) == common[bank_col].astype(str)
            )

    if match_flags:
        all_match = pd.concat(match_flags.values(), axis=1).all(axis=1)
        common["reconciliation_status"] = all_match.map({True: "Match", False: "Mismatch"})
        common["mismatch_reason"] = pd.concat(match_flags, axis=1).apply(
            lambda row: ", ".join(col for col in match_flags if not row[col]), axis=1
        )
    else:
        common["reconciliation_status"] = "Match"
        common["mismatch_reason"] = ""

    results["matches"] = common[common["reconciliation_status"] == "Match"]
    results["mismatches"] = common[common["reconciliation_status"] == "Mismatch"]
    results["detailed"] = common

    # ------------------------------
    # Summary
    # ------------------------------
    results["summary"] = pd.DataFrame({
        "Category": ["Match", "Mismatch", "Missing", "Data Issues"],
        "Count": [
            len(results["matches"]),
            len(results["mismatches"]),
            len(results["missing"]),
            len(results["data_issues"]),
        ],
    })

    return results

# ==============================
# Output Writer
# ==============================

def write_output(results, output_folder):
    """
    Writes reconciliation results into a single Excel file
    with multiple sheets for business users.
    """

    # Remove quotes if pasted
    output_folder = output_folder.strip().strip('"').strip("'")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(
        output_folder,
        f"reconciliation_output_{timestamp}.xlsx"
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    return output_path

# ==============================
# Main Program
# ==============================

def main():
    print("=" * 60)
    print(" Enterprise Finance Reconciliation Tool ")
    print("=" * 60)

    erp_path = input("Enter ERP / System of Record file path: ")
    bank_path = input("Enter External / Bank file path: ")
    output_folder = input("Enter output folder path: ")

    df_erp = preprocess_data(read_spreadsheet(erp_path))
    df_bank = preprocess_data(read_spreadsheet(bank_path))

    results = reconcile_data(df_erp, df_bank)
    output_file = write_output(results, output_folder)

    print("\nReconciliation completed successfully ✅")
    print(f"Output saved to: {output_file}")

# ==============================
# Entry Point
# ==============================

if __name__ == "__main__":
    main()