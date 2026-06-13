import openpyxl

headers = [
    "CONTRACT_DAY", "CONTRACT_MONTH", "CONTRACT_YEAR",
    "CUSTOMER_NAME", "CUSTOMER_ADDRESS",
    "NUM_PLOTS_WORDS", "NUM_PLOTS_DIGITS", "PLOT_SIZE_SQM",
    "TOTAL_PRICE_DIGITS", "TOTAL_PRICE_WORDS",
    "DEPOSIT_DIGITS", "DEPOSIT_WORDS",
    "BALANCE_DIGITS", "BALANCE_WORDS",
    "PAYMENT_START_DATE", "PAYMENT_DURATION_MONTHS", "PAYMENT_END_DATE", "PAYMENT_START_DAY_FULL",
]

sample_row = [
    "5th", "June", "2026",
    "MR. OJOLOWO BOLUWATIFE", "9 OYEBANKE OSUNSANYA CRESCENT LAGOS",
    "Two (2)", "2", "900",
    "3000000", "THREE MILLION NAIRA ONLY",
    "800000.00", "EIGHT HUNDRED THOUSAND NAIRA ONLY",
    "2200000.00", "TWO MILLION TWO HUNDRED THOUSAND NAIRA ONLY",
    "26th March 2026", "12", "26th Day of March 2027", "26TH MARCH 2027",
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Customers"
ws.append(headers)
ws.append(sample_row)

# Auto-fit column widths
for col in ws.columns:
    max_len = 0
    col_letter = col[0].column_letter
    for cell in col:
        try:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        except Exception:
            pass
    ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

wb.save('/home/user/Project-1/customers_template.xlsx')
print("customers_template.xlsx created")
