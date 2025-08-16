#!/usr/bin/env sh

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
