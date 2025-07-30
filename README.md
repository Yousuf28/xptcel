# XPT to Excel Converter

A standalone Python application that converts SAS XPT (XPORT) files to Excel format and automatically opens them for viewing.

## Features

- Converts SAS XPT files to Excel (.xlsx) format
- Automatically opens the converted file in Excel
- Handles missing values by converting them to empty strings
- Creates temporary files with descriptive names
- Can be associated with XPT files for double-click conversion

## Installation Options

### Option 1: Download Pre-built Executable (Recommended for most users)

1. Go to the [Releases](../../releases) page
2. Download the latest `xpt-to-excel.exe` file
3. Place it in a folder of your choice
4. Optional: Associate XPT files with this executable for double-click conversion

### Option 2: Build from Source

#### Prerequisites

- Python 3.8 or higher
- Git (optional, for cloning)

#### Step 1: Clone or Download the Repository

```bash
# Using Git
git clone https://github.com/yourusername/xpt-to-excel-converter.git
cd xpt-to-excel-converter

# Or download and extract the ZIP file from GitHub
```

#### Step 2: Set Up Virtual Environment

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate

# On macOS/Linux:
source venv/bin/activate
```

#### Step 3: Install Dependencies

```bash
pip install -r requirements.txt
```

#### Step 4: Build the Executable

```bash
# Make the build script executable (macOS/Linux only)
chmod +x build.sh

# Run the build script
# On Windows (using Git Bash or WSL):
./build.sh

# On Windows (using Command Prompt):
pyinstaller --windowed --onefile --hidden-import=pyreadstat._readstat_writer --hidden-import=pyreadstat._readstat_parser --collect-all pyreadstat --exclude-module matplotlib --exclude-module scipy --exclude-module sklearn --exclude-module pandas.tests --exclude-module torch --name xpt-to-excel xpt_open_in_excel.py
```

#### Step 5: Find Your Executable

After building, the executable will be located in the `dist/` folder:
- `dist/xpt-to-excel.exe` (Windows)
- `dist/xpt-to-excel` (macOS/Linux)

## Usage

### Command Line
```bash
xpt-to-excel.exe path/to/your/file.xpt
```

### File Association (Windows)
1. Right-click on any XPT file
2. Select "Open with" → "Choose another app"
3. Browse and select `xpt-to-excel.exe`
4. Check "Always use this app to open .xpt files"

## Development

### Running from Source
```bash
python xpt_open_in_excel.py path/to/your/file.xpt
```

### Project Structure
```
xpt-to-excel-converter/
├── xpt_open_in_excel.py      # Main application
├── xpt-to-excel.spec         # PyInstaller spec file
├── build.sh                  # Build script
├── requirements.txt          # Python dependencies
├── README.md                 # This file
└── dist/                     # Generated executables (after building)
```

## Dependencies

See `requirements.txt` for the complete list of Python packages required.

## Troubleshooting

### Antivirus Warnings
If your antivirus software flags the executable as suspicious:
1. This is a common false positive for PyInstaller-generated executables
2. You can build from source to verify the code is safe
3. Add the executable to your antivirus whitelist
4. The executable gains trust over time as more users download it safely

### Build Issues
- Ensure you're using Python 3.8 or higher
- Make sure all dependencies are installed in your virtual environment
- On Windows, you may need to install Microsoft Visual C++ Build Tools

### Runtime Issues
- Ensure the XPT file path doesn't contain special characters
- Make sure you have write permissions in your temp directory
- Verify that Excel or a compatible application is installed to open .xlsx files

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

---

**Note**: This application creates temporary Excel files in your system's temp directory. These files are not automatically deleted to allow you to save them if needed.
```

**requirements.txt**
```txt
pyreadstat==1.2.7
pandas==2.1.4
openpyxl==3.1.2
pyinstaller==6.3.0
```

**build_requirements.txt** (optional, for development)
```txt
pyreadstat==1.2.7
pandas==2.1.4
openpyxl==3.1.2
pyinstaller==6.3.0
setuptools>=65.0.0
wheel>=0.37.0
```
