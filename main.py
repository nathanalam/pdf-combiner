import os
from pathlib import Path

from pypdf import PdfReader, PdfWriter


def main():
    current_dir = Path(__file__).parent
    converted_pdfs = []

    # Get all PDF files (excluding the output file if it exists)
    all_pdf_files = sorted(current_dir.glob("*.pdf"))
    pdf_files = [pdf for pdf in all_pdf_files if pdf.name != "combined_output.pdf"]

    if not pdf_files:
        print("No PDF files found to merge.")
        return

    print(f"Found {len(pdf_files)} PDF file(s) to merge:")
    for pdf in pdf_files:
        marker = "(converted)" if pdf in converted_pdfs else ""
        print(f"  - {pdf.name} {marker}")

    # Create a PdfWriter object
    merger = PdfWriter()

    # Append each PDF
    print("\nMerging PDFs...")
    for pdf_file in pdf_files:
        try:
            reader = PdfReader(str(pdf_file))
            for page in reader.pages:
                merger.add_page(page)
            print(f"  Added: {pdf_file.name}")
        except Exception as e:
            print(f"  Error adding {pdf_file.name}: {e}")

    # Write the merged PDF
    output_file = current_dir / "combined_output.pdf"
    with open(output_file, "wb") as f:
        merger.write(f)

    print(f"\nSuccessfully combined {len(pdf_files)} PDFs into: {output_file.name}")


if __name__ == "__main__":
    main()
