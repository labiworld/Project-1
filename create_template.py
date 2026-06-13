"""
Creates contract_template.docx from the original OJOLOWO_BOLUWATIFS.docx
by replacing customer-specific values with {{PLACEHOLDER}} tokens.
"""
from docx import Document


SOURCE = '/root/.claude/uploads/0a666ae4-7959-5a8f-b94b-8116a61fa32c/c8be110a-OJOLOWO_BOLUWATIFS.docx'
OUTPUT = '/home/user/Project-1/contract_template.docx'


def replace_in_paragraph(paragraph, old, new):
    """Replace old->new in a paragraph, handling text split across runs."""
    if old not in paragraph.text:
        return False

    # Try simple per-run replacement first
    replaced = False
    for run in paragraph.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            replaced = True

    # If still not replaced (split across runs), rebuild into first run
    if old in paragraph.text:
        full_text = ''.join(run.text for run in paragraph.runs)
        if old in full_text:
            new_text = full_text.replace(old, new)
            if paragraph.runs:
                paragraph.runs[0].text = new_text
                for run in paragraph.runs[1:]:
                    run.text = ''
            replaced = True

    return replaced


def replace_all(doc, old, new):
    """Replace text throughout all paragraphs and table cells."""
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


# Ordered list — longer / more specific strings first to avoid partial matches
REPLACEMENTS = [
    # ---- Date line ----
    ('……...Day of…………………. 2026', '{{CONTRACT_DAY}} Day of {{CONTRACT_MONTH}} {{CONTRACT_YEAR}}'),

    # ---- Customer ----
    ('MR. OJOLOWO BOLUWATIFE', '{{CUSTOMER_NAME}}'),
    ('9, OYEBANKE OSUNSANYA CRESCENT, LAGOS', '{{CUSTOMER_ADDRESS}}'),

    # ---- Balance (longest money phrase first) ----
    ('₦2,200,000.00  (TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY)', '₦{{BALANCE_DIGITS}}  ({{BALANCE_WORDS}})'),
    ('TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY', '{{BALANCE_WORDS}}'),
    ('2,200,000.00', '{{BALANCE_DIGITS}}'),

    # ---- Deposit ----
    # In original P38, run4='EIGHT ' run5='HUNDRED THOUSAND NAIRA ONLY)' — handle split
    ('EIGHT HUNDRED THOUSAND NAIRA ONLY', '{{DEPOSIT_WORDS}}'),
    ('800,000.00', '{{DEPOSIT_DIGITS}}'),

    # ---- Total price (word form) ----
    # P37: run5='THREE MILLION NAIRA ONLY'  — fine
    # P115: runs split as '3,000,0' '00,000 (T' 'HREE MILLION NAIRA ONLY' ')'
    #   We handle the word fragment separately after merging
    ('THREE MILLION NAIRA ONLY', '{{TOTAL_PRICE_WORDS}}'),
    # Catch fragment left over from P115 split (T already consumed by merge)
    ('HREE MILLION NAIRA ONLY', '{{TOTAL_PRICE_WORDS}}'),
    # ---- Total price (digit form) ----
    # P37 has '3,000' ',' '000 (' spread across runs — merge first via paragraph scan
    ('3,000,000', '{{TOTAL_PRICE_DIGITS}}'),

    # ---- Payment dates ----
    ('26th Day of March, 2027', '{{PAYMENT_END_DATE}}'),
    # P42 has '26th' split as run8='2' run9='6' run10='th' run11=' March, 2026'
    # and run15='2' run16='6' run17='th' run18=' Day of' run20='March 2026 to...'
    # The paragraph-merge approach will combine them
    ('26th March, 2026', '{{PAYMENT_START_DATE}}'),
    # Second occurrence in P42: "26th Day of March 2026" (split across runs 15-20)
    ('26th Day of March 2026', '{{PAYMENT_START_DATE}}'),
    # Duration
    ('12months', '{{PAYMENT_DURATION_MONTHS}} months'),
    ('12 months', '{{PAYMENT_DURATION_MONTHS}} months'),
    # P118 default date
    ('26TH MARCH, 2027', '{{PAYMENT_START_DAY_FULL}}'),

    # ---- Plot counts ----
    # "Two (2)" lowercase-first in P28
    ('Two (2)', '{{NUM_PLOTS_WORDS}}'),
    # "TWO (2)" uppercase in P38, P83, P115
    ('TWO (2)', '{{NUM_PLOTS_WORDS}}'),

    # ---- Plot size ----
    ('900 Square Meters', '{{PLOT_SIZE_SQM}} Square Meters'),
]


def main():
    doc = Document(SOURCE)

    for old, new in REPLACEMENTS:
        n = replace_all(doc, old, new)
        status = f'({n} place(s))' if n else 'NOT FOUND'
        print(f'  {old[:55]!r:58} -> {new[:40]!r}  {status}')

    # Special-case: P115 has a typo "₦3,000,000,000" with runs deeply split.
    # After the per-run pass the paragraph text may still contain pieces.
    # Do a targeted full-text merge for any paragraph that still has an
    # unreplaced total-price digit fragment.
    for para in doc.paragraphs:
        ft = para.text
        if '3,000,0' in ft and 'TOTAL_PRICE_DIGITS' not in ft:
            full = ''.join(r.text for r in para.runs)
            # Could be 3,000,000 or 3,000,000,000 (typo in original)
            fixed = full.replace('3,000,000,000', '{{TOTAL_PRICE_DIGITS}}')
            fixed = fixed.replace('3,000,000', '{{TOTAL_PRICE_DIGITS}}')
            if fixed != full and para.runs:
                para.runs[0].text = fixed
                for r in para.runs[1:]:
                    r.text = ''
                print(f'  Fixed split price digits in paragraph: {fixed[:80]!r}')

    doc.save(OUTPUT)
    print(f'\nTemplate saved: {OUTPUT}')

    # ---- Verification ----
    doc2 = Document(OUTPUT)
    full_text = '\n'.join(p.text for p in doc2.paragraphs)

    required = [
        'CONTRACT_DAY', 'CONTRACT_MONTH', 'CONTRACT_YEAR',
        'CUSTOMER_NAME', 'CUSTOMER_ADDRESS',
        'NUM_PLOTS_WORDS',
        'PLOT_SIZE_SQM',
        'TOTAL_PRICE_DIGITS', 'TOTAL_PRICE_WORDS',
        'DEPOSIT_DIGITS', 'DEPOSIT_WORDS',
        'BALANCE_DIGITS', 'BALANCE_WORDS',
        'PAYMENT_START_DATE', 'PAYMENT_DURATION_MONTHS',
        'PAYMENT_END_DATE', 'PAYMENT_START_DAY_FULL',
    ]
    print('\nPlaceholder check:')
    all_ok = True
    for p in required:
        found = '{{' + p + '}}' in full_text
        print(f'  {{{{ {p} }}}}: {"OK" if found else "MISSING ***"}')
        if not found:
            all_ok = False
    print('\nTemplate creation:', 'PASS' if all_ok else 'FAIL — some placeholders missing')


if __name__ == '__main__':
    main()
