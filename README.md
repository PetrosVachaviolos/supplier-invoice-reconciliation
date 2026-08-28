# Supplier Invoice Reconciliation Engine

Automated reconciliation engine that matches unposted supplier invoices against internal SAP records, identifying discrepancies and categorizing them for review.

## Business Problem

In large manufacturing environments, supplier invoices submitted to tax authorities (myDATA) must be matched against internal SAP postings. This process was previously done manually — cross-referencing thousands of entries across multiple Excel files, which was time-consuming and error-prone.

This engine automates the entire matching process and produces a structured output that accountants can review immediately.

## How It Works

### Input
| File | Source | Description |
|---|---|---|
| `ΕΛΒΑΛ.xlsx` | SAP export (Power Automate Desktop) | Internal posted transactions |
| `ΠΡΟΜΗΘΕΥΤΕΣ.xlsx` | SAP export (Power Automate Desktop) | Supplier master data (VAT lookup) |
| `ΑΚΑΤΑΧΩΡΗΤΑ.xlsx` | Finance team (shared folder) | Combined ELVAL + HALCOR unposted entries |

> In production, `ΕΛΒΑΛ.xlsx` and `ΠΡΟΜΗΘΕΥΤΕΣ.xlsx` are automatically exported from SAP GUI using a Power Automate Desktop + pywinauto automation, eliminating manual data collection entirely.

### Matching Logic

Each unposted entry is matched against SAP records on **4 keys**: VAT ID, Date, Amount, Reference Number.

| Result | Condition |
|---|---|
| ΒΡΕΘΗΚΑΝ | All 4 keys match exactly |
| ΛΑΘΟΣ ΑΞΙΑ | VAT + Date + REF match, Amount differs |
| ΛΑΘΟΣ ΗΜΕΡΟΜΗΝΙΑ | VAT + Amount + REF match, Date differs |
| ΛΑΘΟΣ ΑΝΑΦΟΡΑ | VAT match only, REF not found |
| ΔΕΝ ΒΡΕΘΗΚΑΝ | No SAP record found (likely HALCOR entries) |

### Reference Number Matching

Suppliers often submit reference numbers in different formats than internal SAP records (e.g. supplier sends `2024001`, SAP stores `TIM-2024-2024001`). The engine strips non-numeric characters and uses a **contains** check to handle this automatically.

### Output

`MATCHING_RESULT.xlsx` with 5 sheets — one per category — formatted as Excel tables with auto-fitted columns.

## Tech Stack

- Python 3
- pandas
- openpyxl

## Usage

```bash
pip install pandas openpyxl
python MATCHING_REPORT_PORTFOLIO_FULL.py
```

The script runs in two phases:
1. **Phase 1** — generates dummy data (simulates SAP exports and supplier file)
2. **Phase 2** — runs the reconciliation engine and produces `MATCHING_RESULT.xlsx`

## Sample Output

| NAME | VAT | REF | AMOUNT | DATE | SAP_DOC_ID | SAP_REF | SAP_AMOUNT | SAP_DATE |
|---|---|---|---|---|---|---|---|---|
| ALPHA INDUSTRIES SA | 123456789 | 5915 | 2397.74 | 26/11/2024 | SAP0056 | TIM-2024-5915 | 2397.74 | 26/11/2024 |
| BETA TRADING SA | 987654321 | 4295 | 13969.25 | 21/08/2024 | SAP0031 | TIM-2024-4295 | 13969.25 | 19/12/2024 |

## Author

Petros Vachaviolos — Business Intelligence & Process Automation Specialist

[LinkedIn](https://www.linkedin.com/in/petros-vachaviolos)
