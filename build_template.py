from docx import Document

def replace_in_paragraph(paragraph, old, new):
    if old in paragraph.text:
        for run in paragraph.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)
        if old in paragraph.text:
            full_text = ''.join(run.text for run in paragraph.runs)
            if old in full_text:
                new_text = full_text.replace(old, new)
                for i, run in enumerate(paragraph.runs):
                    run.text = new_text if i == 0 else ''

doc = Document('/root/.claude/uploads/0a666ae4-7959-5a8f-b94b-8116a61fa32c/c8be110a-OJOLOWO_BOLUWATIFS.docx')

# Order matters: do longer/more specific strings first
replacements = [
    # Date line
    ("……...Day of…………………. 2026", "{{CONTRACT_DAY}} Day of {{CONTRACT_MONTH}} {{CONTRACT_YEAR}}"),
    # Customer name (also appears at signature line)
    ("MR. OJOLOWO BOLUWATIFE", "{{CUSTOMER_NAME}}"),
    # Customer address
    ("9, OYEBANKE OSUNSANYA CRESCENT, LAGOS", "{{CUSTOMER_ADDRESS}}"),
    # Plot descriptions - handle all variants
    ("Two (2) plots of land measuring 900 Square Meters", "{{NUM_PLOTS_WORDS}} plots of land measuring {{PLOT_SIZE_SQM}} Square Meters"),
    ("TWO (2) plot of land measuring 900 Square Meters", "{{NUM_PLOTS_WORDS}} plot of land measuring {{PLOT_SIZE_SQM}} Square Meters"),
    # Monetary - balance (has extra spaces in docx)
    ("₦2,200,000.00  (TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY)", "₦{{BALANCE_DIGITS}} ({{BALANCE_WORDS}})"),
    # Deposit
    ("₦800,000.00 (EIGHT HUNDRED THOUSAND NAIRA ONLY)", "₦{{DEPOSIT_DIGITS}} ({{DEPOSIT_WORDS}})"),
    # Total price - two variants (para 37 split, para 115)
    ("₦3,000,000 (THREE MILLION NAIRA ONLY", "₦{{TOTAL_PRICE_DIGITS}} ({{TOTAL_PRICE_WORDS}}"),
    ("₦3,000,000,000 (THREE MILLION NAIRA ONLY)", "₦{{TOTAL_PRICE_DIGITS}} ({{TOTAL_PRICE_WORDS}})"),
    # Payment dates - longer strings first
    ("26th Day of March, 2027", "{{PAYMENT_END_DATE}}"),
    ("26th Day of March 2026", "{{PAYMENT_START_DATE}}"),
    ("26th March, 2026", "{{PAYMENT_START_DATE}}"),
    ("26TH MARCH, 2027", "{{PAYMENT_START_DAY_FULL}}"),
    # Duration
    ("12months", "{{PAYMENT_DURATION_MONTHS}}months"),
]

def apply_replacements(paragraph):
    for old, new in replacements:
        replace_in_paragraph(paragraph, old, new)

for para in doc.paragraphs:
    apply_replacements(para)

for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                apply_replacements(para)

doc.save('/home/user/Project-1/contract_template.docx')
print("Template saved to /home/user/Project-1/contract_template.docx")

# Verify - show all non-empty paragraphs
doc2 = Document('/home/user/Project-1/contract_template.docx')
print("\nVerifying - paragraphs with placeholders or customer data:")
for i, para in enumerate(doc2.paragraphs):
    t = para.text.strip()
    if t and ('{{' in t or 'OJOLOWO' in t or '3,000,000' in t or '800,000' in t or '2,200,000' in t or '26th' in t or '26TH' in t or '12months' in t or 'Day of' in t):
        print(f"[{i}] {repr(para.text)}")
