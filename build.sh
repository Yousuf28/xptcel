#!/usr/bin/env sh
# pyinstaller --windowed --onefile \
#   --hidden-import=pyreadstat._readstat_writer \
#   --hidden-import=pyreadstat._readstat_parser \
#   --collect-all pyreadstat \
#   --exclude-module matplotlib --exclude-module scipy \
#   --exclude-module sklearn --exclude-module pandas.tests \
#   --exclude-module torch \
# --name xpt-to-excel  xpt_open_in_excel.py


# pyinstaller --onefile --windowed --hidden-import=pyreadstat \
#   --exclude-module matplotlib --exclude-module scipy \
#   --exclude-module sklearn --exclude-module pandas.tests \
#   xpt_viewer.py
# pyinstaller --onefdir --hidden-import "pandas.io.excel._xlsxwriter" --hidden-import "pandas.io.excel._openpyxl" --exclude-module "pandas.io.formats.style" --exclude-module "pandas.plotting" --exclude-module "pandas.tseries"  xpt_to_excel.py

pyinstaller --windowed --onedir \
  --hidden-import=pyreadstat._readstat_writer \
  --hidden-import=pyreadstat._readstat_parser \
  --collect-all pyreadstat \
  --exclude-module matplotlib --exclude-module scipy \
  --exclude-module sklearn --exclude-module pandas.tests \
  --exclude-module torch --exclude-module numpy.tests \
  --exclude-module IPython --exclude-module jupyter \
  --exclude-module notebook --exclude-module tkinter \
  --name xpt-to-excel-fast xpt_open_in_excel.py
