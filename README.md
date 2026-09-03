# 🏦 Enterprise Finance Reconciliation Tool

A rule-driven Python financial reconciliation engine designed with production-oriented validation, auditability, dual-status architecture, and automated testing. It reconciles financial records between an **Enterprise System of Record (ERP)** and an **External Source (Bank / Payment Gateway / Partner)** — producing a deterministic, auditable multi-sheet Excel workbook with executive KPI dashboards and separated output grains.

---

## 📌 Business Problem & Importance

Manual financial reconciliation between ERP ledgers and external bank statements is a risk-prone operation. Misaligned currency formats, floating-point rounding errors, ambiguous duplicate keys, and silent merge fan-outs frequently corrupt audit trails.

This engine automates end-to-end financial reconciliation with **strict data-quality quarantine rules**, **exact `Decimal` monetary precision**, **explicit typed field rules**, **source row traceability**, and **protection against Excel formula injection**.

---

## ✨ Key Features

- **Raw Identifier Preservation (C1)**: Ingests CSV (`dtype=str`) and XLSX (`dtype=object, engine="openpyxl"`) preserving raw textual representations (e.g. leading zero keys `"00123"`). *Note: Excel cells containing leading zeros should be formatted as Text cells.*
- **Strict Text Normalization (C2)**: Uses `normalize_text()` converting `None`, `NaN`, `""`, whitespace, `"null"`, `"none"`, and `"<na>"` cleanly into `None`.
- **Structured Money Parsing & Currency Conflict Check (C3)**: Parses amount, symbol (`$`, `€`, `₹`, `£`), and explicit ISO currency code (`USD`, `EUR`, `INR`). Validates embedded currency against the record's currency field, flagging conflicts as `CURRENCY_AMOUNT_CONFLICT`.
- **Separated Output Grains (C4)**: Outputs distinct sheets for `reconciliation_units` (1 row per logical reconciliation unit) and `source_exceptions` (1 row per problematic physical source record).
- **Disambiguated Exception-Rate KPIs (C5)**: Reports separate `Source Row Exception Rate (%)` (denominator: total ingested source rows) and `Reconciliation Unit Exception Rate (%)` (denominator: total reconciliation units).
- **Strict Required Field Semantics (C6)**: Enforces `"required": true/false`. If a required field is missing on either or both sides, it is flagged under data quality status `MISSING_REQUIRED_VALUE`.
- **Dynamic Per-Money-Field Exposure (C7)**: Dynamically loops over all configured money rules (e.g., `amount`, `tax_amount`), resolving their associated `currency_field` to calculate dynamic currency-isolated monetary exposure.
- **Dual-Status Architecture & Source Validation (C8)**: Separates `reconciliation_status` (`MATCH`, `MISMATCH`, `MISSING_IN_EXTERNAL`, `MISSING_IN_ERP`) from `data_quality_status` (`VALID`, `INVALID_FIELD_VALUE`, `MISSING_REQUIRED_VALUE`, `CURRENCY_AMOUNT_CONFLICT`, `DUPLICATE_KEY_*`).

---

## 🏷️ Machine-Readable Dual-Status Taxonomy

### 1. Reconciliation Status (`reconciliation_status`)
| Status | Description |
|---|---|
| `MATCH` | Record present in both sources and fully aligned across all configured comparison fields |
| `MISMATCH` | Record present in both sources but diverges on monetary, text, date, or required fields |
| `MISSING_IN_EXTERNAL` | Unique key present in ERP, but absent in External source |
| `MISSING_IN_ERP` | Unique key present in External source, but absent in ERP |

### 2. Data Quality Status (`data_quality_status`)
| Status | Description |
|---|---|
| `VALID` | All fields parse cleanly and satisfy required field rules |
| `MISSING_PRIMARY_KEY` | Primary key is missing, blank, `NaN`, or whitespace-only |
| `DUPLICATE_KEY_IN_ERP` | Primary key appears multiple times in ERP (quarantined prior to merge) |
| `DUPLICATE_KEY_IN_EXTERNAL` | Primary key appears multiple times in External source (quarantined prior to merge) |
| `DUPLICATE_KEY_IN_BOTH` | Primary key appears multiple times in both ERP and External sources |
| `INVALID_FIELD_VALUE` | Field value failed numeric/date parsing (e.g. `ERROR123` or invalid date string) |
| `MISSING_REQUIRED_VALUE` | Field configured as `required: true` is blank/missing on either or both datasets |
| `CURRENCY_AMOUNT_CONFLICT` | Embedded currency symbol/code in amount cell (e.g. `€100`) conflicts with currency column (`USD`) |

---

## ⚙️ Reconciliation Pipeline Architecture

```
   ┌──────────────────────┐         ┌──────────────────────────┐
   │    ERP Input File    │         │  External Source File    │
   └──────────┬───────────┘         └────────────┬─────────────┘
              │                                  │
              ▼                                  ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 1. Ingestion & Raw String Identifier Preservation (C1)    │
   └──────────┬────────────────────────────────────────────────┘
              │
              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 2. Header Normalization & Collision Validation (snake_case)│
   └──────────┬────────────────────────────────────────────────┘
              │
              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 3. Explicit Field Rule Normalization & Mandatory Schema Check│
   └──────────┬────────────────────────────────────────────────┘
              │
              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 4. Missing Key Quarantine (data_quality: MISSING_PRIMARY_KEY)│
   └──────────┬────────────────────────────────────────────────┘
              │
              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 5. Duplicate Key Quarantine (data_quality: DUPLICATE_KEY_*)│
   └──────────┬────────────────────────────────────────────────┘
              │
              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 6. One-to-One Outer Merge on Canonical Key (_recon_key)   │
   └──────────┬────────────────────────────────────────────────┘
              │
              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 7. Typed Comparison, Currency Conflict & Required Validation│
   └──────────┬────────────────────────────────────────────────┘
              │
              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 8. Executive KPI Summary & Auditable Excel Export         │
   └───────────────────────────────────────────────────────────┘
```

---

## 📄 Explicit Field Rule Configuration

```json
{
  "key_column": "invoice_id",
  "field_rules": {
    "amount": {
      "type": "money",
      "tolerance": "0.01",
      "required": true,
      "currency_field": "currency"
    },
    "currency": {
      "type": "text",
      "case_sensitive": false,
      "required": true
    },
    "customer_name": {
      "type": "text",
      "case_sensitive": false,
      "required": false
    },
    "date": {
      "type": "date",
      "format": "%Y-%m-%d",
      "required": true
    }
  }
}
```

---

## 📊 Output Excel Structure

The tool generates a timestamped Excel workbook (`reconciliation_YYYYMMDDTHHMMSS_<short_id>.xlsx`) with 7 explicitly ordered worksheets:

| Sheet Name | Contents & Business Purpose |
|---|---|
| `summary` | Executive KPI dashboard (counts, match rate %, exception rates, currency-grouped exposures) |
| `run_metadata` | Audit trail parameters (Run ID, UTC timestamp, file paths, SHA-256 hashes, tolerance, serialized rules) |
| `matches` | All fully reconciled records (`MATCH`) |
| `mismatches` | Records with field discrepancies or data quality flags (`MISMATCH`) |
| `missing` | Unmatched records annotated with origin (`MISSING_IN_EXTERNAL`, `MISSING_IN_ERP`) |
| `source_exceptions` | Physical problematic source records (quarantined duplicate keys, blank primary keys, source data issues) |
| `reconciliation_units` | Master consolidated view containing 1 row per logical reconciliation unit |

---

## 🚀 Installation & Usage

### 1. Installation

```bash
git clone https://github.com/vaibhavi466/Enterprise-Finance-Reconciliation-Tool.git
cd Enterprise-Finance-Reconciliation-Tool
pip install -e ".[dev]"
```

### 2. Generate Sample Data

```bash
python generate_sample_data.py
```

### 3. Run via CLI

```bash
python reconciliation.py \
  --erp sample_data/erp_data.csv \
  --external sample_data/bank_data.xlsx \
  --config reconciliation_config.json \
  --output output \
  --key invoice_id \
  --amount-tolerance 0.01 \
  --overwrite
```

---

## 🧪 Automated Testing & CI

Run unit and integration test suite:

```bash
python -m unittest discover tests
```

Run code formatting and lint checks:

```bash
ruff check .
ruff format --check .
```

### GitHub Actions CI
Automated testing is configured via [.github/workflows/tests.yml](.github/workflows/tests.yml), executing Ruff lint checks, Ruff format checks, unit tests, and CLI end-to-end runs across Python 3.10, 3.11, 3.12, and 3.13 on every `push` and `pull_request`.

---

## 📐 Assumptions & Limitations

1. **Excel String Formatting for Leading Zeros**: Identifier columns containing meaningful leading zeros in Excel files should be formatted as Text cells to preserve exact string values.
2. **Key-Based Reconciliation**: Matching is driven by an explicit primary key column (e.g. `invoice_id`).
3. **No Dynamic FX Conversion**: Amounts in differing currencies (e.g. USD vs EUR) are flagged as currency mismatches rather than dynamically converted using live FX APIs.
4. **Duplicate Quarantine**: Duplicate primary keys are quarantined for operational human review rather than arbitrarily resolved.
