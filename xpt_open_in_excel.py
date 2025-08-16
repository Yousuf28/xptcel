import sys
import os
import pyreadstat
import pandas as pd
import tempfile
import openpyxl
from pathlib import Path

def convert_xpt_to_excel(xpt_file):
    # Get the base name of the XPT file (without path and extension)
    base_name = Path(xpt_file).stem

    # Read the XPT file
    df, meta = pyreadstat.read_xport(xpt_file)

    # Replace missing values with empty string
    df = df.fillna('')

    # Create a named temporary file  name
    with tempfile.NamedTemporaryFile(prefix=f"{base_name}_", suffix='.xlsx', delete=False) as tmp_file:
        temp_excel_file = tmp_file.name

    # Save the DataFrame to the temporary Excel file with the sheet name as the base name
    with pd.ExcelWriter(temp_excel_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=base_name, index=False)

    print(f"Temporary Excel file created at: {temp_excel_file}")
    print("Opening Excel file...")

    try:
        if platform.system() == "Darwin":  # macOS
            os.system(f"open '{temp_excel_file}'")
        elif platform.system() == "Windows":
            os.startfile(temp_excel_file)
        else:  # Linux
            os.system(f"xdg-open '{temp_excel_file}'")
    except Exception as e:
        print(f"Could not open file automatically: {e}")
        print(f"Excel file saved at: {temp_excel_file}")
        print("Please open it manually.")


    # # Open the Excel file
    # try:
    #     os.startfile(temp_excel_file)
    # except AttributeError:
    #     print(f"Excel file saved at: {temp_excel_file}")
    #     print("Please open it manually.")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        xpt_file = sys.argv[1]
        convert_xpt_to_excel(xpt_file)
    else:
        print("No file specified. Please associate this program with XPT files.")

