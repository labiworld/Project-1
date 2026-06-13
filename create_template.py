"""
Script to create contract_template.docx from the original document.
Replaces customer-specific text with placeholders.
"""
from docx import Document
import copy


def replace_in_paragraph(paragraph, old, new):
    """Replace text in a paragraph, handling text split across runs."""
    if old not in paragraph.text:
        return False

    # First try simple per-run replacement
    replaced = False
    for run in paragraph.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            replaced = True

    # If still not replaced (split across runs), rebuild full text
    if old in paragraph.text:
        full_text = ''.join(run.text for run in paragraph.runs)
        if old in full_text:
            new_text = full_text.replace(old, new)
            for i, run in enumerate(paragraph.runs):
                run.text = new_text if i == 0 else ''
            replaced = True

    return replaced


def replace_in_doc(doc, old, new):
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


def create_template():
    src = '/root/.claude/uploads/0a666ae4-7959-5a8f-b94b-8116a61fa32c/c8be110a-OJOLOWO_BOLUWATIFS.docx'
    doc = Document(src)

    # Order matters: replace longer/more specific strings first to avoid partial matches

    replacements = [
        # Date line
        ('……...Day of…………………. 2026', '{{CONTRACT_DAY}} Day of {{CONTRACT_MONTH}} {{CONTRACT_YEAR}}'),

        # Customer name (appears in multiple places)
        ('MR. OJOLOWO BOLUWATIFE', '{{CUSTOMER_NAME}}'),

        # Customer address
        ('9, OYEBANKE OSUNSANYA CRESCENT, LAGOS', '{{CUSTOMER_ADDRESS}}'),

        # Prices - longer/more specific first
        ('THREE MILLION NAIRA ONLY', '{{TOTAL_PRICE_WORDS}}'),
        ('HREE MILLION NAIRA ONLY', '{{TOTAL_PRICE_WORDS_FRAGMENT}}'),  # handle split run in P115
        ('EIGHT HUNDRED THOUSAND NAIRA ONLY', '{{DEPOSIT_WORDS}}'),
        ('EIGHT \nHUNDRED THOUSAND NAIRA ONLY', '{{DEPOSIT_WORDS}}'),
        ('TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY', '{{BALANCE_WORDS}}'),
        ('TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY)', '{{BALANCE_WORDS}})'),

        # Numeric amounts - more specific first
        ('₦3,000,000 (', '₦{{TOTAL_PRICE_DIGITS}} ('),
        ('₦800,000.00 (', '₦{{DEPOSIT_DIGITS}} ('),
        ('₦2,200,000.00  (', '₦{{BALANCE_DIGITS}}  ('),

        # Plot counts - more specific first
        ('Two (2)', '{{NUM_PLOTS_WORDS}}'),
        ('TWO (2)', '{{NUM_PLOTS_WORDS_UPPER}}'),

        # Plot size
        ('900 Square Meters', '{{PLOT_SIZE_SQM}} Square Meters'),

        # Payment dates and durations
        ('26th March, 2026', '{{PAYMENT_START_DATE}}'),
        ('12months', '{{PAYMENT_DURATION_MONTHS}}months'),
        ('26th Day of March 2026', '{{PAYMENT_START_DATE}} Day'),   # partial - handle carefully
        ('26th Day of March, 2027', '{{PAYMENT_END_DATE}}'),
        ('26TH MARCH, 2027', '{{PAYMENT_START_DAY_FULL}}'),
    ]

    for old, new in replacements:
        n = replace_in_doc(doc, old, new)
        if n:
            print(f'  Replaced "{old[:50]}" -> "{new[:50]}" ({n} para(s))')
        # else:
        #     print(f'  NOT FOUND: "{old[:60]}"')

    # Now fix up the fragment from P115 split run
    # P115: '₦3,000,000,000 (THREE MILLION...' - this is a typo in original doc (3 billion vs 3 million)
    # The run split was: run6='3,000,0' run7='00,000 (T' run8='HREE MILLION NAIRA ONLY'
    # After the replacements above these runs will have been merged or may still be split
    # Let's do targeted fix for that paragraph
    for para in doc.paragraphs:
        if '{{TOTAL_PRICE_WORDS_FRAGMENT}}' in para.text:
            replace_in_paragraph(para, '{{TOTAL_PRICE_WORDS_FRAGMENT}}', '{{TOTAL_PRICE_WORDS}}')
            print('  Fixed TOTAL_PRICE_WORDS_FRAGMENT -> TOTAL_PRICE_WORDS')

    # Fix '{{NUM_PLOTS_WORDS_UPPER}}' back to proper placeholder
    # In context: "TWO (2)" appears in several places, we want consistent placeholder
    for para in doc.paragraphs:
        if '{{NUM_PLOTS_WORDS_UPPER}}' in para.text:
            replace_in_paragraph(para, '{{NUM_PLOTS_WORDS_UPPER}}', '{{NUM_PLOTS_WORDS}}')

    # Handle the P42 split runs for dates: '26th' is split as run8='2' run9='6' run10='th'
    # After merging via replace_in_paragraph approach they should be combined
    # But payment start date text '26th March, 2026' may be split
    # Let's check and fix P42 specifically by looking at full text
    for para in doc.paragraphs:
        if '{{PAYMENT_START_DATE}} Day' in para.text:
            # This was incorrectly replaced, revert that part
            replace_in_paragraph(para, '{{PAYMENT_START_DATE}} Day', '{{PAYMENT_START_DATE}}')
            print('  Fixed PAYMENT_START_DATE Day fragment')

    # Fix 12months -> {{PAYMENT_DURATION_MONTHS}} months
    for para in doc.paragraphs:
        if '{{PAYMENT_DURATION_MONTHS}}months' in para.text:
            replace_in_paragraph(para, '{{PAYMENT_DURATION_MONTHS}}months', '{{PAYMENT_DURATION_MONTHS}} months')
            print('  Fixed PAYMENT_DURATION_MONTHS months')

    # Handle ₦3,000,0 split in P115 (the actual price paragraph)
    # After replacing '₦3,000,000 (' we also need to handle P115 which shows ₦3,000,000,000
    # but the runs are split: run6='3,000,0' run7='00,000 (T'
    # The full_text approach should handle this when we replace '3,000,0\n00,000 (T'
    # Let's scan for any remaining issues
    for para in doc.paragraphs:
        ft = para.text
        if '3,000,0' in ft and '00,000' in ft:
            # This is the split price in P115 - fix it
            full = ''.join(r.text for r in para.runs)
            if '3,000,0' in full:
                new_full = full.replace('3,000,000,000', '{{TOTAL_PRICE_DIGITS}}')
                if new_full != full:
                    for i, run in enumerate(para.runs):
                        run.text = new_full if i == 0 else ''
                    print('  Fixed split 3,000,000,000 in P115')

    # Save template
    out = '/home/user/Project-1/contract_template.docx'
    doc.save(out)
    print(f'\nTemplate saved to: {out}')

    # Verify
    doc2 = Document(out)
    print('\nVerification - paragraphs containing placeholders:')
    for i, para in enumerate(doc2.paragraphs):
        if '{{' in para.text:
            print(f'  P{i}: {para.text[:120]}')


if __name__ == '__main__':
    create_template()
