#!/usr/bin/env python3
"""
LandCulture Investment Ltd - Contract Generator
Supports chat mode (interactive) and batch mode (--batch customers.xlsx)
"""

import argparse
import copy
import os
import sys
from datetime import datetime

from docx import Document

DEFAULT_TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'contract_template.docx')
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'contracts')

PLACEHOLDERS = [
    ("CONTRACT_DAY", "Contract day (e.g. 5th)"),
    ("CONTRACT_MONTH", "Contract month (e.g. June)"),
    ("CONTRACT_YEAR", "Contract year (e.g. 2026)"),
    ("CUSTOMER_NAME", "Customer full name (e.g. MR. JOHN DOE)"),
    ("CUSTOMER_ADDRESS", "Customer address"),
    ("NUM_PLOTS_WORDS", "Number of plots in words (e.g. Two (2))"),
    ("NUM_PLOTS_DIGITS", "Number of plots in digits (e.g. 2)"),
    ("PLOT_SIZE_SQM", "Plot size in square meters (e.g. 900)"),
    ("TOTAL_PRICE_DIGITS", "Total price in digits (e.g. 3,000,000)"),
    ("TOTAL_PRICE_WORDS", "Total price in words (e.g. THREE MILLION NAIRA ONLY)"),
    ("DEPOSIT_DIGITS", "Deposit amount in digits (e.g. 800,000.00)"),
    ("DEPOSIT_WORDS", "Deposit amount in words (e.g. EIGHT HUNDRED THOUSAND NAIRA ONLY)"),
    ("BALANCE_DIGITS", "Balance amount in digits (e.g. 2,200,000.00)"),
    ("BALANCE_WORDS", "Balance amount in words (e.g. TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY)"),
    ("PAYMENT_START_DATE", "Payment start date (e.g. 26th March, 2026)"),
    ("PAYMENT_DURATION_MONTHS", "Payment duration in months (e.g. 12)"),
    ("PAYMENT_END_DATE", "Payment end date (e.g. 26th Day of March, 2027)"),
    ("PAYMENT_START_DAY_FULL", "Payment deadline full date (e.g. 26TH MARCH, 2027)"),
]


def replace_in_paragraph(paragraph, old, new):
    """Replace text in a paragraph, handling run-split text."""
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


def replace_in_xml(element, old, new):
    """Replace text directly in XML — handles text boxes and drawing canvases."""
    from lxml import etree
    xml_str = etree.tostring(element, encoding='unicode')
    if old in xml_str:
        xml_str = xml_str.replace(f'>{old}<', f'>{new}<')
        new_element = etree.fromstring(xml_str)
        element.getparent().replace(element, new_element)


def generate_contract(fields_dict, template_path=DEFAULT_TEMPLATE):
    """
    Generate a contract docx from a fields dictionary and template.
    Returns the output file path.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    doc = Document(template_path)

    for key, value in fields_dict.items():
        placeholder = f"{{{{{key}}}}}"
        val = str(value)
        # Replace in normal paragraphs
        for para in doc.paragraphs:
            replace_in_paragraph(para, placeholder, val)
            # Also replace in text boxes / drawings inside this paragraph
            if placeholder in para._element.xml:
                replace_in_xml(para._element, placeholder, val)
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        replace_in_paragraph(para, placeholder, val)

    customer_name = fields_dict.get("CUSTOMER_NAME", "CUSTOMER").replace(" ", "_").replace(".", "")
    date_str = datetime.now().strftime("%Y%m%d")
    filename = f"CONTRACT_{customer_name}_{date_str}.docx"
    output_path = os.path.join(OUTPUT_DIR, filename)

    doc.save(output_path)
    return output_path


def chat_mode(template_path=DEFAULT_TEMPLATE):
    """Interactive chat mode: ask one question at a time."""
    print("\n=== LandCulture Investment Ltd - Contract Generator ===")
    print("Enter customer details (press Ctrl+C to cancel)\n")

    fields = {}
    for key, prompt in PLACEHOLDERS:
        while True:
            value = input(f"{prompt}: ").strip()
            if value:
                fields[key] = value
                break
            print("  (This field cannot be empty, please enter a value)")

    print("\n=== Summary ===")
    for key, _ in PLACEHOLDERS:
        print(f"  {key}: {fields[key]}")

    confirm = input("\nGenerate contract with these details? (yes/no): ").strip().lower()
    if confirm not in ("yes", "y"):
        print("Cancelled.")
        return

    output_path = generate_contract(fields, template_path)
    print(f"\nContract generated: {output_path}")


def batch_mode(excel_path, template_path=DEFAULT_TEMPLATE):
    """Batch mode: read Excel file and generate one contract per row."""
    try:
        import openpyxl
    except ImportError:
        print("openpyxl is required for batch mode. Install with: pip install openpyxl")
        sys.exit(1)

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    print(f"Found headers: {headers}")

    generated = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(v is None for v in row):
            continue
        fields = {}
        for col_idx, header in enumerate(headers):
            if header and col_idx < len(row):
                val = row[col_idx]
                fields[header] = str(val) if val is not None else ""

        print(f"\nRow {row_idx}: Generating contract for {fields.get('CUSTOMER_NAME', 'UNKNOWN')}...")
        try:
            output_path = generate_contract(fields, template_path)
            generated.append(output_path)
            print(f"  -> Saved: {output_path}")
        except Exception as e:
            print(f"  -> ERROR: {e}")

    print(f"\nBatch complete. {len(generated)} contract(s) generated.")
    return generated


def main():
    parser = argparse.ArgumentParser(
        description="LandCulture Investment Ltd - Contract Generator"
    )
    parser.add_argument(
        "--batch",
        metavar="CUSTOMERS_XLSX",
        help="Batch mode: path to Excel file with customer data",
    )
    parser.add_argument(
        "--template",
        default=DEFAULT_TEMPLATE,
        help=f"Path to contract template docx (default: {DEFAULT_TEMPLATE})",
    )
    args = parser.parse_args()

    if args.batch:
        batch_mode(args.batch, args.template)
    else:
        chat_mode(args.template)


if __name__ == "__main__":
    main()
