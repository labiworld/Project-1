# LandCulture Investment Ltd — Contract of Sale Generator

An AI agent that auto-fills a "Contract of Sale" Word document template for LandCulture Investment Ltd.

## Prerequisites

Install dependencies:

```bash
pip install python-docx openpyxl
```

## Project Files

| File | Description |
|------|-------------|
| `contract_template.docx` | Word template with `{{PLACEHOLDER}}` markers |
| `generate_contract.py` | Main agent script (chat + batch modes) |
| `customers_template.xlsx` | Excel template for batch mode |
| `test_generate.py` | Non-interactive test with sample data |

---

## Usage

### Chat Mode (Interactive — one customer at a time)

```bash
python generate_contract.py
```

The script will prompt you for each field one at a time, show a summary, ask for confirmation, then save the contract.

Optionally specify an output directory:

```bash
python generate_contract.py --output-dir ./contracts
```

### Batch Mode (multiple customers from Excel)

1. Fill in `customers_template.xlsx` — one row per customer.
2. Run:

```bash
python generate_contract.py --batch customers_template.xlsx
```

Each row produces one `.docx` file named `CONTRACT_{CUSTOMER_NAME}_{DATE}.docx` in the same directory as the Excel file.

You can also specify a different output directory:

```bash
python generate_contract.py --batch customers_template.xlsx --output-dir ./contracts
```

---

## Placeholders Reference

| Placeholder | Example Value |
|---|---|
| `{{CONTRACT_DAY}}` | `5th` |
| `{{CONTRACT_MONTH}}` | `June` |
| `{{CONTRACT_YEAR}}` | `2026` |
| `{{CUSTOMER_NAME}}` | `MR. OJOLOWO BOLUWATIFE` |
| `{{CUSTOMER_ADDRESS}}` | `9, OYEBANKE OSUNSANYA CRESCENT, LAGOS` |
| `{{NUM_PLOTS_WORDS}}` | `Two (2)` |
| `{{NUM_PLOTS_DIGITS}}` | `2` |
| `{{PLOT_SIZE_SQM}}` | `900` |
| `{{TOTAL_PRICE_DIGITS}}` | `3,000,000` |
| `{{TOTAL_PRICE_WORDS}}` | `THREE MILLION NAIRA ONLY` |
| `{{DEPOSIT_DIGITS}}` | `800,000.00` |
| `{{DEPOSIT_WORDS}}` | `EIGHT HUNDRED THOUSAND NAIRA ONLY` |
| `{{BALANCE_DIGITS}}` | `2,200,000.00` |
| `{{BALANCE_WORDS}}` | `TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY` |
| `{{PAYMENT_START_DATE}}` | `26th March, 2026` |
| `{{PAYMENT_DURATION_MONTHS}}` | `12` |
| `{{PAYMENT_END_DATE}}` | `26th Day of March, 2027` |
| `{{PAYMENT_START_DAY_FULL}}` | `26TH MARCH, 2027` |
| `{{NUM_PLOTS_WORDS_UPPER}}` | `TWO (2)` |

---

## Running the Test

To verify the setup works without interactive input:

```bash
python test_generate.py
```

Expected output:
```
Running test generation for OJOLOWO BOLUWATIFE...
SUCCESS: Contract generated at:
  /home/user/Project-1/contracts/CONTRACT_MR_OJOLOWO_BOLUWATIFE_2026-06-13.docx
All placeholders successfully replaced.

Spot checks:
  PASS: contract date
  PASS: customer name
  PASS: customer address
  PASS: payment start date
  PASS: payment deadline
```
