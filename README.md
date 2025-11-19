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

1. Go to the [action](https://github.com/Yousuf28/xptcel/actions/runs/17003257797) page
2. Download the zip file 
3. Place it in a folder of your choice and unzip

### Option 2: Build from Source

#### Prerequisites

- Python 3.8 or higher
- Git (optional, for cloning)

#### Step 1: Clone or Download the Repository

```bash
# Using Git
git clone git@github.com:Yousuf28/xptcel.git
cd xptcel 

# Or download and extract the ZIP file from GitHub
# for macos, use macos branch
```

#### Step 2: Set Up Virtual Environment

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# windows git Bash
source venv/Script/activate

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

```

#### Step 5: Find Your Executable

After building, the executable will be located in the `dist/` folder:
- `dist/xpt-to-excel-fast/xpt-to-excel-fast.exe`

## Usage


### File Association (Windows)
1. Right-click on any XPT file
2. Select "Open with" → "Choose another app"
3. Browse and select `xpt-to-excel-fast.exe`
4. Check "Always use this app to open .xpt files"

## Development

### Running from Source
```bash
python xpt_open_in_excel.py path/to/your/file.xpt
```

### Project Structure
```
xptcel/
├── xpt_open_in_excel.py      # Main application
├── xpt-to-excel-fast.spec    # PyInstaller spec file
├── build.sh                  # Build script
├── requirements.txt          # Python dependencies
├── README.md                 # This file
└── dist/                     # Generated executables (after building)
```

## Dependencies

See `requirements.txt` for the complete list of Python packages required.



## License

This project is licensed under the MIT License.

---

**Note**: This application creates temporary Excel files in your system's temp directory. These files are not automatically deleted to allow you to save them if needed.
