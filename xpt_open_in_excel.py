import sys
import os
import pyreadstat
import pandas as pd
import tempfile
import openpyxl
import subprocess
from pathlib import Path

def convert_xpt_to_excel(xpt_file):
    # Get the base name of the XPT file (without path and extension)
    base_name = Path(xpt_file).stem

    # Read the XPT file
    df, meta = pyreadstat.read_xport(xpt_file)

    # Replace missing values with empty string
    df = df.fillna('')

    # Create a named temporary file name
    with tempfile.NamedTemporaryFile(prefix=f"{base_name}_", suffix='.xlsx', delete=False) as tmp_file:
        temp_excel_file = tmp_file.name

    # Save the DataFrame to the temporary Excel file with the sheet name as the base name
    with pd.ExcelWriter(temp_excel_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=base_name, index=False)

        # Get the workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets[base_name]

        # Calculate and set column widths
        for column in df.columns:
            column_index = df.columns.get_loc(column) + 1  # openpyxl is 1-indexed
            column_letter = openpyxl.utils.get_column_letter(column_index)

            # Calculate max length for this column
            # Include column header length
            max_length = len(str(column))

            # Check all values in the column
            for value in df[column]:
                if value is not None:
                    max_length = max(max_length, len(str(value)))

            # Cap the width at 40 and add a little padding
            adjusted_width = min(max_length + 2, 40)

            # Set the column width
            worksheet.column_dimensions[column_letter].width = adjusted_width

    print(f"Temporary Excel file created at: {temp_excel_file}")
    print("Opening Excel file...")

    try:
        # Try Windows first (os.startfile only exists on Windows)
        os.startfile(temp_excel_file)
        print("Excel file opened successfully!")
    except AttributeError:

        try:
            # Required when running as a macOS GUI bundle
            from AppKit import NSApplication, NSApplicationActivationPolicyRegular
            NSApplication.sharedApplication().setActivationPolicy_(NSApplicationActivationPolicyRegular)
        except Exception:
            pass


        # Not Windows, try macOS
        try:
            # subprocess.run(["/Applications/LibreOffice.app/Contents/MacOS/soffice",
            #                 temp_excel_file])
            subprocess.run(["/usr/bin/open", temp_excel_file], check=True)


            print("Excel file opened successfully!")
        except (subprocess.CalledProcessError, FileNotFoundError):
            # macOS failed or not macOS, fallback
            print(f"Could not open file automatically.")
            print(f"Excel file saved at: {temp_excel_file}")
            print("Please open it manually.")
    except OSError as e:
        # Windows but os.startfile failed
        print(f"Could not open file automatically: {e}")
        print(f"Excel file saved at: {temp_excel_file}")
        print("Please open it manually.")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        xpt_file = sys.argv[1]
        convert_xpt_to_excel(xpt_file)
    else:
        print("No file specified. Please associate this program with XPT files.")
