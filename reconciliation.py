"""
Enterprise Finance Reconciliation Tool
======================================

Purpose:
--------
Performs automated end-to-end reconciliation of financial records between an Enterprise 
System of Record (ERP) and an External Source (Bank / Payment Gateway / Partner).

Key Features:
-------------
1. Multi-format Support: Ingests CSV, XLSX, and XLS files smoothly.
2. Data Sanitization: Standardizes headers, strips cell whitespace, and cleans formatted currency strings.
3. Strict Duplicate Isolation: Identifies and isolates duplicate primary key records prior to matching
   to prevent false positives.
4. Smart Field-Level Comparison:
   - Amount/Numeric fields: Evaluated using a configurable tolerance threshold (default: ±0.01).
   - Text/Categorical fields: Evaluated via standardized exact match.
5. Comprehensive Reporting: Outputs a multi-sheet Excel report with high-level summary metrics,
   exact matches, field-level mismatch explanations, missing records, and duplicate data issues.
"""

import os
import sys
import re
from datetime import datetime
import pandas as pd

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls"}
DEFAULT_TOLERANCE = 0.01
DEFAULT_KEY_COLUMN = "invoice_id"


def read_spreadsheet(file_path: str) -> pd.DataFrame:
    """Reads a CSV or Excel file into a pandas DataFrame."""
    file_path = file_path.strip().strip('"').strip("'")
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: '{file_path}'")

    _, ext = os.path.splitext(file_path)
    ext = ext.lower()

    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported file type '{ext}'. Supported types: {', '.join(SUPPORTED_EXTENSIONS)}"
        )

    if ext == ".csv":
        return pd.read_csv(file_path)
    else:
        return pd.read_excel(file_path, engine="openpyxl")


def sanitize_column_name(col: str) -> str:
    """Normalizes column headers to lowercase snake_case."""
    return str(col).strip().lower().replace(" ", "_")


def preprocess_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans and standardizes the DataFrame:
    - Normalizes column names.
    - Drops completely empty rows.
    - Trims leading/trailing whitespace from string fields.
    """
    df = df.copy()
    df.columns = [sanitize_column_name(c) for c in df.columns]
    df = df.dropna(how="all")

    # Trim string cells
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    return df


def parse_numeric(val):
    """
    Safely converts values (including formatted currency strings like '$1,250.00') to float.
    Returns None if parsing fails.
    """
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        cleaned = re.sub(r"[^\d.-]", "", val)
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def reconcile_data(
    df_erp: pd.DataFrame,
    df_bank: pd.DataFrame,
    key_column: str = DEFAULT_KEY_COLUMN,
    amount_tolerance: float = DEFAULT_TOLERANCE,
) -> dict[str, pd.DataFrame]:
    """
    Reconciles ERP and Bank datasets on a primary key column.

    Returns a dictionary of DataFrames:
    - 'summary': High-level reconciliation metrics
    - 'matches': Fully reconciled records across all shared columns
    - 'mismatches': Records present in both systems but with field-level discrepancies
    - 'missing': Records present in only one system
    - 'data_issues': Isolated duplicate key records
    - 'detailed': Merged dataset containing all common records with status annotations
    """
    results = {}

    key_col = sanitize_column_name(key_column)
    if key_col not in df_erp.columns:
        raise KeyError(
            f"Key column '{key_col}' not found in ERP file. Available columns: {list(df_erp.columns)}"
        )
    if key_col not in df_bank.columns:
        raise KeyError(
            f"Key column '{key_col}' not found in Bank file. Available columns: {list(df_bank.columns)}"
        )

    # -------------------------------------------------------------
    # 1. Duplicate Detection & Isolation
    # -------------------------------------------------------------
    dupes_erp = df_erp[df_erp.duplicated(subset=[key_col], keep=False)].copy()
    dupes_bank = df_bank[df_bank.duplicated(subset=[key_col], keep=False)].copy()

    non_empty_dupes = []
    if not dupes_erp.empty:
        non_empty_dupes.append(dupes_erp.assign(source_file="ERP"))
    if not dupes_bank.empty:
        non_empty_dupes.append(dupes_bank.assign(source_file="BANK"))

    if non_empty_dupes:
        results["data_issues"] = pd.concat(non_empty_dupes, ignore_index=True)
    else:
        results["data_issues"] = pd.DataFrame(columns=list(df_erp.columns) + ["source_file"])

    # Isolate duplicates from main dataset prior to matching to avoid false positives
    df_erp_clean = df_erp[~df_erp[key_col].isin(dupes_erp[key_col])].copy()
    df_bank_clean = df_bank[~df_bank[key_col].isin(dupes_bank[key_col])].copy()

    # -------------------------------------------------------------
    # 2. Outer Merge on Primary Key
    # -------------------------------------------------------------
    merged = df_erp_clean.merge(
        df_bank_clean,
        on=key_col,
        how="outer",
        suffixes=("_erp", "_bank"),
        indicator=True,
    )

    # -------------------------------------------------------------
    # 3. Process Unmatched / Missing Records
    # -------------------------------------------------------------
    missing = merged[merged["_merge"] != "both"].copy()
    missing["missing_status"] = missing["_merge"].map(
        {
            "left_only": "Present in ERP only (Missing in Bank)",
            "right_only": "Present in Bank only (Missing in ERP)",
        }
    )
    missing = missing.drop(columns=["_merge"])
    results["missing"] = missing

    # -------------------------------------------------------------
    # 4. Field-Level Comparison for Common Records
    # -------------------------------------------------------------
    common = merged[merged["_merge"] == "both"].copy().drop(columns=["_merge"])
    shared_cols = [
        c for c in df_erp_clean.columns if c in df_bank_clean.columns and c != key_col
    ]

    match_flags = {}
    mismatch_details = {}

    for col in shared_cols:
        erp_col = f"{col}_erp"
        bank_col = f"{col}_bank"

        # Differentiate between numeric amount fields vs categorical/text fields
        is_amount_col = any(
            kw in col for kw in ["amount", "price", "fee", "total", "val", "cost", "sum"]
        )

        if is_amount_col:
            erp_nums = common[erp_col].apply(parse_numeric)
            bank_nums = common[bank_col].apply(parse_numeric)

            both_null = erp_nums.isna() & bank_nums.isna()
            diff = (erp_nums - bank_nums).abs()
            is_match = both_null | (diff <= amount_tolerance)

            match_flags[col] = is_match
            mismatch_details[col] = common.apply(
                lambda r, ec=erp_col, bc=bank_col, m=is_match: ""
                if m[r.name]
                else f"{col} diff (ERP: {r[ec]}, Bank: {r[bc]})",
                axis=1,
            )
        else:
            erp_str = common[erp_col].astype(str).str.strip().str.lower()
            bank_str = common[bank_col].astype(str).str.strip().str.lower()

            both_null = common[erp_col].isna() & common[bank_col].isna()
            is_match = both_null | (erp_str == bank_str)

            match_flags[col] = is_match
            mismatch_details[col] = common.apply(
                lambda r, ec=erp_col, bc=bank_col, m=is_match: ""
                if m[r.name]
                else f"{col} mismatch (ERP: '{r[ec]}', Bank: '{r[bc]}')",
                axis=1,
            )

    if match_flags:
        all_match = pd.concat(match_flags.values(), axis=1).all(axis=1)
        common["reconciliation_status"] = all_match.map(
            {True: "Match", False: "Mismatch"}
        )

        mismatch_df = pd.concat(mismatch_details, axis=1)
        common["mismatch_reason"] = mismatch_df.apply(
            lambda row: "; ".join([val for val in row if val != ""]), axis=1
        )
    else:
        common["reconciliation_status"] = "Match"
        common["mismatch_reason"] = ""

    results["matches"] = common[common["reconciliation_status"] == "Match"].copy()
    results["mismatches"] = common[common["reconciliation_status"] == "Mismatch"].copy()
    results["detailed"] = common

    # -------------------------------------------------------------
    # 5. Summary Dashboard Metrics
    # -------------------------------------------------------------
    results["summary"] = pd.DataFrame(
        {
            "Metric / Category": [
                "Total ERP Records Ingested",
                "Total Bank Records Ingested",
                "Fully Reconciled Matches",
                "Discrepancies / Mismatches",
                "Unmatched / Missing Records",
                "Data Quality Issues (Duplicates)",
            ],
            "Count": [
                len(df_erp),
                len(df_bank),
                len(results["matches"]),
                len(results["mismatches"]),
                len(results["missing"]),
                len(results["data_issues"]),
            ],
        }
    )

    return results


def write_output(results: dict[str, pd.DataFrame], output_folder: str) -> str:
    """Writes reconciliation results into a multi-tab Excel workbook."""
    output_folder = output_folder.strip().strip('"').strip("'")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"reconciliation_output_{timestamp}.xlsx"
    output_path = os.path.join(output_folder, output_filename)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    return output_path


def main():
    print("=" * 65)
    print(" 🏦 Enterprise Finance Reconciliation Tool ")
    print("=" * 65)

    erp_path = input("Enter ERP / System of Record file path: ")
    bank_path = input("Enter External / Bank file path: ")
    output_folder = input("Enter output directory path: ")

    key_input = input(f"Enter Key Column name [Default: '{DEFAULT_KEY_COLUMN}']: ").strip()
    key_column = key_input if key_input else DEFAULT_KEY_COLUMN

    try:
        print("\n⏳ Ingesting and preprocessing files...")
        df_erp = preprocess_data(read_spreadsheet(erp_path))
        df_bank = preprocess_data(read_spreadsheet(bank_path))

        print("⚡ Reconciling datasets...")
        results = reconcile_data(df_erp, df_bank, key_column=key_column)

        print("💾 Saving Excel report...")
        output_file = write_output(results, output_folder)

        print("\n✅ Reconciliation Completed Successfully!")
        print(f"📊 Summary Results:")
        for _, row in results["summary"].iterrows():
            print(f"   - {row['Metric / Category']}: {row['Count']}")
        print(f"\n📂 Output File Saved To:\n   {output_file}\n")

    except Exception as e:
        print(f"\n❌ Error during reconciliation: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
