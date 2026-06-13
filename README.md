# LandCulture Investment Ltd - AI Contract Generator

Automatically generates land purchase contracts from the `contract_template.docx` template.

## Requirements

```bash
pip install python-docx openpyxl
```

## Usage

### Chat Mode (default)

Run without arguments to be prompted for each field one at a time:

```bash
python3 generate_contract.py
```

You will be asked to enter:
- Contract date (day, month, year)
- Customer name and address
- Number of plots and plot size
- Pricing details (total, deposit, balance)
- Payment schedule (start date, duration, end date)

After entering all fields, a summary is shown and you confirm before the contract is generated.

**Output:** `contracts/CONTRACT_{CUSTOMER_NAME}_{DATE}.docx`

### Batch Mode

Process multiple customers from an Excel file:

```bash
python3 generate_contract.py --batch customers_template.xlsx
```

The Excel file must have these exact column headers (see `customers_template.xlsx` for a template):

| Column | Example |
|--------|---------|
| CONTRACT_DAY | 5th |
| CONTRACT_MONTH | June |
| CONTRACT_YEAR | 2026 |
| CUSTOMER_NAME | MR. JOHN DOE |
| CUSTOMER_ADDRESS | 9, OYEBANKE OSUNSANYA CRESCENT, LAGOS |
| NUM_PLOTS_WORDS | Two (2) |
| NUM_PLOTS_DIGITS | 2 |
| PLOT_SIZE_SQM | 900 |
| TOTAL_PRICE_DIGITS | 3,000,000 |
| TOTAL_PRICE_WORDS | THREE MILLION NAIRA ONLY |
| DEPOSIT_DIGITS | 800,000.00 |
| DEPOSIT_WORDS | EIGHT HUNDRED THOUSAND NAIRA ONLY |
| BALANCE_DIGITS | 2,200,000.00 |
| BALANCE_WORDS | TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY |
| PAYMENT_START_DATE | 26th March, 2026 |
| PAYMENT_DURATION_MONTHS | 12 |
| PAYMENT_END_DATE | 26th Day of March, 2027 |
| PAYMENT_START_DAY_FULL | 26TH MARCH, 2027 |

One `.docx` contract is generated per row. All outputs are saved to the `contracts/` folder.

### Custom Template

To use a different template file:

```bash
python3 generate_contract.py --template /path/to/your_template.docx
python3 generate_contract.py --batch customers.xlsx --template /path/to/your_template.docx
```

## Files

| File | Purpose |
|------|---------|
| `contract_template.docx` | Template with `{{PLACEHOLDER}}` markers |
| `generate_contract.py` | Main contract generation script |
| `customers_template.xlsx` | Sample Excel file for batch mode |
| `contracts/` | Output folder for generated contracts |

## Example: Programmatic Use

```python
from generate_contract import generate_contract

fields = {
    "CONTRACT_DAY": "5th",
    "CONTRACT_MONTH": "June",
    "CONTRACT_YEAR": "2026",
    "CUSTOMER_NAME": "MR. JOHN DOE",
    "CUSTOMER_ADDRESS": "12, EXAMPLE STREET, ABUJA",
    "NUM_PLOTS_WORDS": "One (1)",
    "NUM_PLOTS_DIGITS": "1",
    "PLOT_SIZE_SQM": "450",
    "TOTAL_PRICE_DIGITS": "1,500,000",
    "TOTAL_PRICE_WORDS": "ONE MILLION FIVE HUNDRED THOUSAND NAIRA ONLY",
    "DEPOSIT_DIGITS": "500,000.00",
    "DEPOSIT_WORDS": "FIVE HUNDRED THOUSAND NAIRA ONLY",
    "BALANCE_DIGITS": "1,000,000.00",
    "BALANCE_WORDS": "ONE MILLION NAIRA ONLY",
    "PAYMENT_START_DATE": "1st July, 2026",
    "PAYMENT_DURATION_MONTHS": "12",
    "PAYMENT_END_DATE": "1st Day of July, 2027",
    "PAYMENT_START_DAY_FULL": "1ST JULY, 2027",
}

output_path = generate_contract(fields, 'contract_template.docx')
print(f"Contract saved to: {output_path}")
```
