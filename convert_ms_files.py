import os
import sys

import comtypes.client


def convert_pptx_to_pdf(input_file_path, output_file_path=None):
    """
    Converts a PowerPoint .pptx file to a .pdf file.
    ...
    """

    # --- PowerPoint constants ---
    # From https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
    ppFormatPDF = 32

    # ... (Path handling logic remains the same) ...
    if not os.path.isabs(input_file_path):
        input_file_path = os.path.abspath(input_file_path)

    if output_file_path is None:
        file_name, _ = os.path.splitext(input_file_path)
        output_file_path = file_name + ".pdf"
    elif not os.path.isabs(output_file_path):
        output_file_path = os.path.abspath(output_file_path)

    output_dir = os.path.dirname(output_file_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    powerpoint = None
    deck = None

    try:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = True

        print(f"Opening presentation: {input_file_path}")
        deck = powerpoint.Presentations.Open(input_file_path, WithWindow=False)

        print(f"Saving PDF to: {output_file_path}")
        deck.SaveAs(output_file_path, ppFormatPDF)

        print("Conversion successful.")

    except Exception as e:
        print(f"An error occurred during PowerPoint conversion:")
        print(str(e))

    finally:
        if deck:
            deck.Close()
            print("Closed presentation.")
        if powerpoint:
            powerpoint.Quit()
            print("Quit PowerPoint application.")

        deck = None
        powerpoint = None
        # comtypes.CoUninitialize() # <-- REMOVED (Called ONCE at the end)


def convert_xlsx_to_pdf(input_file_path, output_file_path=None):
    """
    Converts an Excel .xlsx file to a .pdf file.

    Args:
        input_file_path (str): The absolute path to the input .xlsx file.
        output_file_path (str, optional): The absolute path for the output .pdf.
                                         If None, it saves the PDF in the same
                                         directory as the input file with the
                                         same name.
    """

    # --- Excel constants ---
    # From https://learn.microsoft.com/en-us/office/vba/api/excel.xlfixedformattype
    xlTypePDF = 0

    # --- Path handling ---
    if not os.path.isabs(input_file_path):
        input_file_path = os.path.abspath(input_file_path)

    if output_file_path is None:
        file_name, _ = os.path.splitext(input_file_path)
        output_file_path = file_name + ".pdf"
    elif not os.path.isabs(output_file_path):
        output_file_path = os.path.abspath(output_file_path)

    output_dir = os.path.dirname(output_file_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    excel = None
    workbook = None

    try:
        # Start the Excel application
        excel = comtypes.client.CreateObject("Excel.Application")
        # Set visibility to True, similar to the PowerPoint fix
        excel.Visible = True

        print(f"Opening workbook: {input_file_path}")

        # Open the workbook
        workbook = excel.Workbooks.Open(input_file_path)

        print(f"Saving PDF to: {output_file_path}")

        # Export as PDF
        # Signature: ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas)
        workbook.ExportAsFixedFormat(xlTypePDF, output_file_path)

        print("Conversion successful.")

    except Exception as e:
        print(f"An error occurred during Excel conversion:")
        print(str(e))

    finally:
        # Always close the workbook and quit Excel
        if workbook:
            workbook.Close(
                SaveChanges=False
            )  # Do not save changes to the original file
            print("Closed workbook.")
        if excel:
            excel.Quit()
            print("Quit Excel application.")

        # Clean up COM references
        workbook = None
        excel = None
        # comtypes.CoUninitialize() # <-- REMOVED (Called ONCE at the end)


def main():
    """
    Main function to find and convert all .pptx/.ppt and .xlsx/.xls files
    in the current directory. Skips files that have already been converted.
    """

    print(f"Scanning for presentations and spreadsheets in: {os.getcwd()}")

    converted_count = 0
    skipped_count = 0

    for file in os.listdir():
        # Determine file name and extension
        file_name, file_ext = os.path.splitext(file)

        # Check if this is a file we care about
        is_powerpoint = file_ext.lower() in (".pptx", ".ppt")
        is_excel = file_ext.lower() in (".xlsx", ".xls")

        if not (is_powerpoint or is_excel):
            continue  # Not an office file, skip it

        # --- Check if output PDF already exists ---
        pdf_output_name = file_name + ".pdf"

        if os.path.exists(pdf_output_name):
            print(f"Skipping '{file}'; PDF '{pdf_output_name}' already exists.")
            skipped_count += 1
            continue  # Skip to the next file
        # -----------------------------------------

        # Get absolute path for the file (only if we're converting)
        abs_input_path = os.path.abspath(file)

        # Check for PowerPoint files
        if is_powerpoint:
            print("---")  # Separator for clarity
            # Pass None so the function creates the default output name
            convert_pptx_to_pdf(abs_input_path, None)
            converted_count += 1

        # Check for Excel files
        elif is_excel:
            print("---")  # Separator for clarity
            # Pass None so the function creates the default output name
            convert_xlsx_to_pdf(abs_input_path, None)
            converted_count += 1

    print("---")
    if converted_count == 0 and skipped_count == 0:
        print("No .pptx, .ppt, .xlsx, or .xls files found to convert.")
    else:
        print(
            f"Finished. Converted {converted_count} file(s), skipped {skipped_count} existing file(s)."
        )


if __name__ == "__main__":
    # Initialize the COM library for this thread
    is_com_initialized = False
    try:
        comtypes.CoInitialize()
        is_com_initialized = True
        main()
    except Exception as e:
        print(f"An error occurred: {e}")
        # Check if it's the main init that failed
        if not is_com_initialized:
            print(f"Failed to initialize COM library: {e}")
        sys.exit(1)
    finally:
        # Uninitialize COM *only if* it was successfully initialized
        if is_com_initialized:
            comtypes.CoUninitialize()
            print("Uninitialized COM library.")
