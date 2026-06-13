"""Create the customers_template.xlsx with all placeholder field headers."""
import openpyxl

FIELD_KEYS = [
    "CONTRACT_DAY",
    "CONTRACT_MONTH",
    "CONTRACT_YEAR",
    "CUSTOMER_NAME",
    "CUSTOMER_ADDRESS",
    "NUM_PLOTS_WORDS",
    "NUM_PLOTS_DIGITS",
    "PLOT_SIZE_SQM",
    "TOTAL_PRICE_DIGITS",
    "TOTAL_PRICE_WORDS",
    "DEPOSIT_DIGITS",
    "DEPOSIT_WORDS",
    "BALANCE_DIGITS",
    "BALANCE_WORDS",
    "PAYMENT_START_DATE",
    "PAYMENT_DURATION_MONTHS",
    "PAYMENT_END_DATE",
    "PAYMENT_START_DAY_FULL",
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Customers"

# Write headers
for col, key in enumerate(FIELD_KEYS, start=1):
    cell = ws.cell(row=1, column=col, value=key)
    cell.font = openpyxl.styles.Font(bold=True)

# Write one sample row
sample = [
    "5th",
    "June",
    "2026",
    "MR. OJOLOWO BOLUWATIFE",
    "9, OYEBANKE OSUNSANYA CRESCENT, LAGOS",
    "Two (2)",
    "2",
    "900",
    "3,000,000",
    "THREE MILLION NAIRA ONLY",
    "800,000.00",
    "EIGHT HUNDRED THOUSAND NAIRA ONLY",
    "2,200,000.00",
    "TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY",
    "26th March, 2026",
    "12",
    "26th Day of March, 2027",
    "26TH MARCH, 2027",
]
for col, val in enumerate(sample, start=1):
    ws.cell(row=2, column=col, value=val)

# Auto-width columns
for col in ws.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    ws.column_dimensions[col[0].column_letter].width = max_len + 4

out = "/home/user/Project-1/customers_template.xlsx"
wb.save(out)
print(f"Created: {out}")
