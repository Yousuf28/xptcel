import sys
import os
import traceback
import subprocess
from pathlib import Path
import uuid

import pyreadstat
import pandas as pd
import openpyxl
import tempfile


########################################################
# ENABLE OR DISABLE LOGGING
########################################################
ENABLE_LOGGING = False   # True = create log file, False = no logging


########################################################
# FIX HOME DIRECTORY FOR macOS APP BUNDLE
########################################################

home = str(Path.home())
if home in ("/var/empty", "/", "", None) or not os.path.isdir(home):
    try:
        home = str(Path(os.path.abspath(__file__)).parent)
    except Exception:
        home = "/tmp"

os.environ["HOME"] = home


########################################################
# LOGGING HELPER
########################################################

LOGFILE = Path(os.environ["HOME"]) / "xpt_open_in_excel.log"

def log(msg):
    if not ENABLE_LOGGING:
        return
    try:
        with open(LOGFILE, "a", encoding="utf-8") as f:
            f.write(str(msg) + "\n")
    except Exception:
        pass



########################################################
# CONVERSION LOGIC
########################################################

def convert_xpt_to_excel(xpt_file):
    log("\n" + "=" * 70)
    log("START CONVERSION")
    log(f"xpt_file = {xpt_file}")

    base_name = Path(xpt_file).stem

    # Load XPT
    try:
        df, meta = pyreadstat.read_xport(xpt_file)
        df = df.fillna("")
        log("Loaded XPT successfully.")
    except Exception:
        log("Error loading XPT:")
        log(traceback.format_exc())
        return

    # Save Excel into HOME
    # excel_path = Path(os.environ["HOME"]) / f"{base_name}_{uuid.uuid4().hex}.xlsx"

    temp_dir = Path(tempfile.gettempdir())
    excel_path = temp_dir / f"{base_name}_{uuid.uuid4().hex}.xlsx"
    excel_path_str = str(excel_path)

    # excel_path_str = str(excel_path)

    try:
        with pd.ExcelWriter(excel_path_str, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=base_name, index=False)
            ws = writer.sheets[base_name]

            # Column widths
            for col in df.columns:
                idx = df.columns.get_loc(col) + 1
                letter = openpyxl.utils.get_column_letter(idx)
                max_len = max(len(str(col)), *(len(str(v)) for v in df[col]))
                ws.column_dimensions[letter].width = min(max_len + 2, 40)

        log(f"Excel saved: {excel_path_str}")
    except Exception:
        log("Error saving Excel:")
        log(traceback.format_exc())
        return

    print(f"Created: {excel_path_str}")
    print("Opening…")

    # Windows open
    try:
        os.startfile(excel_path_str)
        log("Opened via os.startfile")
        return
    except Exception:
        pass

    # macOS open
    if sys.platform == "darwin":
        try:
            result = subprocess.run(
                ["/usr/bin/open", excel_path_str],
                capture_output=True,
                text=True
            )
            log(f"macOS open returncode: {result.returncode}")

            if result.returncode == 0:
                log("File opened successfully.")
                return

            subprocess.run(["/usr/bin/open", "-R", excel_path_str])
            log("Revealed in Finder.")

        except Exception:
            log("macOS open() failed:")
            log(traceback.format_exc())



########################################################
# macOS FILE-OPEN EVENT HANDLER
########################################################

if sys.platform == "darwin":
    from Cocoa import NSApplication, NSObject

    class AppDelegate(NSObject):
        def application_openFile_(self, app, filename):
            log(f"macOS open-file event: {filename}")

            try:
                convert_xpt_to_excel(filename)
            except Exception:
                log("Conversion error:")
                log(traceback.format_exc())

            # Quit app after handling
            try:
                app.terminate_(None)
            except Exception:
                log("Terminate error:")
                log(traceback.format_exc())

            return True

    delegate = AppDelegate.alloc().init()
    NSApplication.sharedApplication().setDelegate_(delegate)



########################################################
# ENTRY POINT
########################################################

if __name__ == "__main__":
    try:
        log("Application launched")
        log(f"Executable: {sys.executable}")

        # macOS app bundle — must run NSApplication loop
        if sys.platform == "darwin":
            app = NSApplication.sharedApplication()
            app.run()   # REQUIRED for Finder double-click support
            sys.exit(0)

        # Windows/Linux command-line usage
        if len(sys.argv) > 1:
            convert_xpt_to_excel(sys.argv[1])
        else:
            print("No file provided.")
            log("No file provided.")

    except Exception as e:
        log("Fatal error:")
        log(traceback.format_exc())
        raise
