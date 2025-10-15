#!/usr/bin/env python3
import pdfplumber
import pandas as pd
from pathlib import Path
import argparse
import tempfile
from pdf2image import convert_from_path
import pytesseract

def ocr_pdf_to_text(pdf_path):
    images = convert_from_path(pdf_path)
    text_pages = []
    for img in images:
        text = pytesseract.image_to_string(img)
        text_pages.append(text)
    return text_pages

def pdf_to_excel(pdf_path: str, output_path: str = None):
    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        raise FileNotFoundError(f"File not found: {pdf_file}")

    all_tables = []

    try:
        with pdfplumber.open(pdf_file) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                for table in tables:
                    df = pd.DataFrame(table)
                    df.insert(0, "Page", i)
                    all_tables.append(df)
    except Exception:
        pass

    # OCR fallback if no tables found
    if not all_tables:
        print("No tables detected, performing OCR...")
        temp_dir = tempfile.mkdtemp()
        images = convert_from_path(pdf_file, output_folder=temp_dir)
        for i, img in enumerate(images, start=1):
            text = pytesseract.image_to_string(img)
            lines = [line.split() for line in text.splitlines() if line.strip()]
            if lines:
                df = pd.DataFrame(lines)
                df.insert(0, "Page", i)
                all_tables.append(df)

    if not all_tables:
        raise ValueError("No readable tables or text found in PDF.")

    combined_df = pd.concat(all_tables, ignore_index=True)
    output_file = Path(output_path) if output_path else pdf_file.with_suffix(".xlsx")
    combined_df.to_excel(output_file, index=False)
    print(f"Saved to {output_file}")

def main():
    parser = argparse.ArgumentParser(
        description="Extract tables from a PDF (scanned or digital) and export to Excel."
    )
    parser.add_argument("pdf_path", help="Path to input PDF file")
    parser.add_argument("-o", "--output", help="Path to output Excel file (.xlsx)")
    args = parser.parse_args()
    pdf_to_excel(args.pdf_path, args.output)

if __name__ == "__main__":
    main()

