import os
from pathlib import Path
from pypdf import PdfWriter, PdfReader
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet


def convert_xlsx_to_pdf(xlsx_path: Path, output_path: Path) -> bool:
    """Convert an XLSX file to PDF format."""
    try:
        # Load the workbook
        wb = load_workbook(xlsx_path, data_only=True)
        
        # Create PDF
        doc = SimpleDocTemplate(str(output_path), pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()
        
        # Process each sheet
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # Add sheet name as header
            if len(wb.sheetnames) > 1:
                elements.append(Paragraph(f"<b>Sheet: {sheet_name}</b>", styles['Heading1']))
                elements.append(Spacer(1, 0.2*inch))
            
            # Get all data from the sheet
            data = []
            for row in ws.iter_rows(values_only=True):
                # Convert None to empty string and all values to strings
                row_data = [str(cell) if cell is not None else "" for cell in row]
                # Skip completely empty rows
                if any(cell for cell in row_data):
                    data.append(row_data)
            
            if data:
                # Create table
                table = Table(data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                ]))
                elements.append(table)
                elements.append(Spacer(1, 0.5*inch))
        
        # Build PDF
        doc.build(elements)
        return True
    except Exception as e:
        print(f"  Error converting {xlsx_path.name}: {e}")
        return False


def main():
    current_dir = Path(__file__).parent
    
    # Convert XLSX files to PDF first
    xlsx_files = sorted(current_dir.glob("*.xlsx"))
    converted_pdfs = []
    
    if xlsx_files:
        print(f"Found {len(xlsx_files)} XLSX file(s) to convert:")
        for xlsx_file in xlsx_files:
            print(f"  - {xlsx_file.name}")
        
        print("\nConverting XLSX files to PDF...")
        for xlsx_file in xlsx_files:
            output_pdf = current_dir / f"{xlsx_file.stem}_converted.pdf"
            if convert_xlsx_to_pdf(xlsx_file, output_pdf):
                print(f"  Converted: {xlsx_file.name} -> {output_pdf.name}")
                converted_pdfs.append(output_pdf)
        print()
    
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
