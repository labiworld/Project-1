"""
Build contract_template.docx from the original OJOLOWO contract.
Replaces customer-specific values with placeholders.
"""
from docx import Document

SRC = '/root/.claude/uploads/0a666ae4-7959-5a8f-b94b-8116a61fa32c/c8be110a-OJOLOWO_BOLUWATIFS.docx'
DST = '/home/user/Project-1/contract_template.docx'

doc = Document(SRC)


def replace_in_paragraph(paragraph, old, new):
    """Replace old with new inside a paragraph, handling run-splits."""
    if old not in paragraph.text:
        return False
    # First pass: simple per-run replacement
    replaced = False
    for run in paragraph.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            replaced = True
    # Second pass: rebuild if still present (split across runs)
    if old in paragraph.text:
        full_text = ''.join(run.text for run in paragraph.runs)
        if old in full_text:
            new_text = full_text.replace(old, new)
            for i, run in enumerate(paragraph.runs):
                run.text = new_text if i == 0 else ''
            replaced = True
    return replaced


def replace_all(doc, old, new):
    count = 0
    for para in doc.paragraphs:
        if replace_in_paragraph(para, old, new):
            count += 1
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if replace_in_paragraph(para, old, new):
                        count += 1
    return count


# --- Date header (paragraph 1) ---
replace_all(doc, '……...Day of…………………. 2026', '{{CONTRACT_DAY}} Day of {{CONTRACT_MONTH}} {{CONTRACT_YEAR}}')

# --- Customer name & address ---
replace_all(doc, 'MR. OJOLOWO BOLUWATIFE of 9, OYEBANKE OSUNSANYA CRESCENT, LAGOS',
            '{{CUSTOMER_NAME}} of {{CUSTOMER_ADDRESS}}')
# Standalone name (signature block, paragraph 228)
replace_all(doc, 'MR. OJOLOWO BOLUWATIFE', '{{CUSTOMER_NAME}}')

# --- Plot descriptions ---
# Para 28: "Two (2)" and "900 Square Meters"
replace_all(doc, 'Two (2)', '{{NUM_PLOTS_WORDS}}')
# TWO (2) in bold sections
replace_all(doc, 'TWO (2)', '{{NUM_PLOTS_WORDS_UPPER}}')
replace_all(doc, '900 Square Meters', '{{PLOT_SIZE_SQM}} Square Meters')

# --- Para 37: total price split across runs ---
# The runs contain: '₦', '3,000', ',000 (', 'THREE MILLION NAIRA ONLY'
# Full text reconstructed = '₦3,000,000 (THREE MILLION NAIRA ONLY'
replace_all(doc, '₦3,000,000 (THREE MILLION NAIRA ONLY',
            '₦{{TOTAL_PRICE_DIGITS}} ({{TOTAL_PRICE_WORDS}}')

# --- Para 38: deposit split across runs ---
# runs: '₦', '8', '00,000.00 (', 'EIGHT ', 'HUNDRED THOUSAND NAIRA ONLY)'
replace_all(doc, '₦800,000.00 (EIGHT HUNDRED THOUSAND NAIRA ONLY)',
            '₦{{DEPOSIT_DIGITS}} ({{DEPOSIT_WORDS}})')

# --- Para 42: balance and dates ---
replace_all(doc, '₦2,200,000.00  (TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY)',
            '₦{{BALANCE_DIGITS}} ({{BALANCE_WORDS}})')
# dates in para 42 - split runs: '2','6','th',' March, 2026'
replace_all(doc, '26th March, 2026', '{{PAYMENT_START_DATE}}')
replace_all(doc, '12months', '{{PAYMENT_DURATION_MONTHS}}months')
# "26th Day of March 2026" (without comma) then "26th Day of March, 2027"
replace_all(doc, '26th Day of March 2026', '{{PAYMENT_START_DATE}}')
replace_all(doc, '26th Day of March, 2027', '{{PAYMENT_END_DATE}}')

# --- Para 115: agreed price - split: '₦','3,000,0','00,000 (T','HREE MILLION NAIRA ONLY',')'
replace_all(doc, '₦3,000,000,000 (THREE MILLION NAIRA ONLY)',
            '₦{{TOTAL_PRICE_DIGITS}} ({{TOTAL_PRICE_WORDS}})')

# --- Para 117-118: default date (26TH MARCH, 2027) ---
replace_all(doc, '26TH MARCH, 2027', '{{PAYMENT_START_DAY_FULL}}')

doc.save(DST)
print(f"Saved template to {DST}")

# Verify key placeholders present
doc2 = Document(DST)
full = '\n'.join(p.text for p in doc2.paragraphs)
placeholders = [
    '{{CONTRACT_DAY}}', '{{CONTRACT_MONTH}}', '{{CONTRACT_YEAR}}',
    '{{CUSTOMER_NAME}}', '{{CUSTOMER_ADDRESS}}',
    '{{NUM_PLOTS_WORDS}}', '{{PLOT_SIZE_SQM}}',
    '{{TOTAL_PRICE_DIGITS}}', '{{TOTAL_PRICE_WORDS}}',
    '{{DEPOSIT_DIGITS}}', '{{DEPOSIT_WORDS}}',
    '{{BALANCE_DIGITS}}', '{{BALANCE_WORDS}}',
    '{{PAYMENT_START_DATE}}', '{{PAYMENT_DURATION_MONTHS}}',
    '{{PAYMENT_END_DATE}}', '{{PAYMENT_START_DAY_FULL}}',
]
print("\nPlaceholder check:")
for p in placeholders:
    found = p in full
    print(f"  {'OK' if found else 'MISSING'}: {p}")
