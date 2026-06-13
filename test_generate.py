#!/usr/bin/env python3
"""
Test script: generates a contract using sample OJOLOWO BOLUWATIFE data.
Run with: python test_generate.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from generate_contract import generate_contract

SAMPLE_DATA = {
    "CONTRACT_DAY":            "5th",
    "CONTRACT_MONTH":          "June",
    "CONTRACT_YEAR":           "2026",
    "CUSTOMER_NAME":           "MR. OJOLOWO BOLUWATIFE",
    "CUSTOMER_ADDRESS":        "9, OYEBANKE OSUNSANYA CRESCENT, LAGOS",
    "NUM_PLOTS_WORDS":         "Two (2)",
    "NUM_PLOTS_DIGITS":        "2",
    "PLOT_SIZE_SQM":           "900",
    "TOTAL_PRICE_DIGITS":      "3,000,000",
    "TOTAL_PRICE_WORDS":       "THREE MILLION NAIRA ONLY",
    "DEPOSIT_DIGITS":          "800,000.00",
    "DEPOSIT_WORDS":           "EIGHT HUNDRED THOUSAND NAIRA ONLY",
    "BALANCE_DIGITS":          "2,200,000.00",
    "BALANCE_WORDS":           "TWO MILLION, TWO HUNDRED THOUSAND NAIRA ONLY",
    "PAYMENT_START_DATE":      "26th March, 2026",
    "PAYMENT_DURATION_MONTHS": "12",
    "PAYMENT_END_DATE":        "26th Day of March, 2027",
    "PAYMENT_START_DAY_FULL":  "26TH MARCH, 2027",
}

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "contracts")

def main():
    print("Running test generation for OJOLOWO BOLUWATIFE...")
    out_path = generate_contract(SAMPLE_DATA, output_dir=OUTPUT_DIR)
    print(f"SUCCESS: Contract generated at:\n  {out_path}")

    # Quick verification: open and check some placeholders are gone
    from docx import Document
    doc = Document(out_path)
    full_text = "\n".join(p.text for p in doc.paragraphs)

    remaining = [f"{{{{{k}}}}}" for k in SAMPLE_DATA if f"{{{{{k}}}}}" in full_text]
    if remaining:
        print(f"\nWARNING: These placeholders were NOT replaced: {remaining}")
    else:
        print("All placeholders successfully replaced.")

    # Spot-check a few values
    checks = [
        ("5th Day of June 2026", "contract date"),
        ("MR. OJOLOWO BOLUWATIFE", "customer name"),
        ("9, OYEBANKE OSUNSANYA CRESCENT, LAGOS", "customer address"),
        ("26th March, 2026", "payment start date"),
        ("26TH MARCH, 2027", "payment deadline"),
    ]
    print("\nSpot checks:")
    for value, label in checks:
        found = value in full_text
        print(f"  {'PASS' if found else 'FAIL'}: {label} ({value!r})")

if __name__ == "__main__":
    main()
