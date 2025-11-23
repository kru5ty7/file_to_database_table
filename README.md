# File to Database Table Converter

A modern Windows GUI application to convert CSV and Excel files to SQL Server database tables with advanced data preview and column customization. Built with CustomTkinter for a sleek, contemporary interface.

## Features

- **Batch File Processing** - Add multiple CSV/Excel files to a queue with dynamic display
- **Modern Data Preview** - Interactive preview of first 20 rows with column-by-column editing
- **Column Customization** - Rename columns and override data types with visual controls
- **Real-Time Statistics** - View row counts, column counts, and NULL value statistics per column
- **Multi-Sheet Support** - Handle Excel files with multiple sheets seamlessly
- **Connection Management** - Store encrypted database credentials securely
- **Progress Tracking** - Visual progress bar with percentage display and detailed logging
- **Auto Type Detection** - Intelligent SQL type inference (BIGINT, FLOAT, NVARCHAR, DATE, DATETIME)
- **Modern UI** - Clean, professional interface built with CustomTkinter

## 🚀 Quick Start

### Option 1: Download Executable (Recommended)

1. Go to [Releases](../../releases)
2. Download the latest `FileToDBConverter-vX.X.X-windows.zip`
3. Extract and run `FileToDBConverter.exe`
4. No Python installation required!

### Option 2: Run from Source

**Using uv (Recommended):**
```bash
# Clone the repository
git clone https://github.com/kru5ty7/file_to_database_table.git
cd file_to_database_table

# Install uv
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# Install dependencies
uv sync

# Run the application
uv run python -m src.app
```

**Using pip:**
```bash
# Clone and create virtual environment
git clone https://github.com/kru5ty7/file_to_database_table.git
cd file_to_database_table
python -m venv venv
.\venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python -m src.app
```

## 📖 Usage Guide

### 1. Add Files to Queue
- Click **"Add Files"** button
- Select one or more CSV or Excel files
- Files appear in the queue list

### 2. Preview and Customize (Optional but Recommended!)
- Select a file from the queue
- Click **"Preview File"** button
- In the modern preview dialog:
  - View **overall statistics** (total rows, columns, NULL counts)
  - Browse **sheet selector** for multi-sheet Excel files
  - See **column-by-column layout** with:
    - Editable column name field
    - Detected type display (auto-inferred)
    - SQL type dropdown selector
    - Per-column NULL count and percentage
    - First 20 data rows displayed vertically
  - **Horizontal scroll** for files with many columns
  - Click **"Apply Changes"** to save customizations
  - Click **"Reset to Defaults"** to revert changes

### 3. Configure Database Connection
- Click **"Manage..."** button
- Add a new connection with:
  - Server name
  - Database name
  - Username and password (encrypted)
  - Driver (default: ODBC Driver 17 for SQL Server)
- Click **"Test"** to verify connection
- Click **"Save"**

### 4. Convert to Database
- Select a connection from the dropdown
- Click **"Convert to Database"**
- Monitor progress in the log window
- Tables are created with sanitized names

## 🎯 Column Type Detection

The application intelligently detects SQL types:

| Data Pattern | Detected Type |
|--------------|---------------|
| Integers (< 2 billion) | `BIGINT` |
| Large integers (> 2 billion) | `NVARCHAR(MAX)` |
| Decimal numbers | `FLOAT` |
| Text/Mixed content | `NVARCHAR(MAX)` |

You can override these in the preview dialog!

## Requirements

- **Windows 10/11**
- **SQL Server** with ODBC Driver 17 or later
- For source installation: Python 3.9+

### Installing ODBC Driver

Download from: [Microsoft ODBC Driver for SQL Server](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server)

## 🏗️ Building from Source

### Local Build

```bash
# Install dependencies (including PyInstaller)
uv sync
# or
pip install -r requirements.txt

# Build executable
build_exe.bat

# Or manually:
pyinstaller FileToDBConverter.spec

# Executable created at: dist\FileToDBConverter.exe
```

### Automated GitHub Release

This project uses GitHub Actions for automatic releases:

1. **Create a version tag**:
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```

2. **GitHub Actions automatically**:
   - Builds the Windows executable
   - Creates a GitHub Release
   - Uploads the `.exe` file

See [RELEASE.md](RELEASE.md) for detailed instructions.

## Project Structure

```
file_to_database_table/
├── src/
│   ├── app.py                  # Application entry point
│   ├── gui_main.py             # Main GUI window
│   ├── database.py             # Database operations
│   ├── file_processor.py       # File reading and type inference
│   ├── utils.py                # Encryption and utilities
│   └── dialogs/
│       ├── preview_dialog.py   # Data preview dialog
│       └── connection_dialog.py # Connection management dialog
├── pyproject.toml              # Project configuration (uv/pip)
├── requirements.txt            # Python dependencies
├── FileToDBConverter.spec      # PyInstaller build configuration
├── build_exe.bat               # Local build script
├── .github/
│   └── workflows/
│       └── build-release.yml   # GitHub Actions workflow
├── logs/                       # Application logs (auto-created)
└── config.json                 # Encrypted connection configs (auto-created)
```

## 🔒 Security

- **Passwords are encrypted** using Fernet (symmetric encryption)
- Connection details stored in `config.json`
- Encryption key stored in `.encryption_key`
- Both files are `.gitignore`d for safety

## Troubleshooting

### Connection Fails
- Verify SQL Server is running
- Check Windows Firewall settings
- Ensure ODBC Driver is installed
- Test connection using "Test Connection" button

### Files Not Showing in Queue
- Check file permissions
- Ensure files are valid CSV/Excel format
- Look for errors in the log window

### Preview Dialog Issues
- Empty sheets (0 rows) will show "No data rows in this sheet"
- Large files may take time to load
- Files with many columns require horizontal scrolling
- Corrupted files will show error messages in the log

### Build Issues
- Ensure all dependencies are installed: `uv sync` or `pip install -r requirements.txt`
- Clean previous builds: `rmdir /s /q build dist`
- Check `build\FileToDBConverter\warn-FileToDBConverter.txt` for warnings
- Ensure numpy 2.0+ is installed (required for pre-built wheels)

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Commit changes: `git commit -m 'Add feature'`
4. Push to branch: `git push origin feature-name`
5. Open a Pull Request

## 📄 License

This project is open source and available under the MIT License.

## Acknowledgments

Built with:
- **CustomTkinter** - Modern GUI framework
- **pandas** - Data processing and analysis
- **pyodbc** - SQL Server connectivity
- **openpyxl** - Excel file handling
- **cryptography** - Password encryption (Fernet)
- **PyInstaller** - Executable packaging
- **numpy** - Numerical operations
- **Pillow** - Image handling for UI
- **darkdetect** - System theme detection

## 📞 Support

For issues, questions, or feature requests:
- Open an [Issue](../../issues)
- Check [RELEASE.md](RELEASE.md) for build documentation

---

**Made with ❤️ by kru5ty7**
