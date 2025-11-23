# File to Database Table Converter

A powerful Windows GUI application to convert CSV and Excel files to SQL Server database tables with advanced data preview and column customization.

## ✨ Features

- **📁 Batch File Processing** - Add multiple CSV/Excel files to a queue
- **👁️ Data Preview** - Preview first 20 rows before importing
- **✏️ Column Editing** - Rename columns and override data types
- **📊 Statistics** - View row counts, column counts, and NULL value statistics
- **🗄️ Multi-Sheet Support** - Handle Excel files with multiple sheets
- **🔐 Connection Management** - Store encrypted database credentials
- **📈 Progress Tracking** - Real-time progress bars and detailed logging
- **🎯 Auto Type Detection** - Intelligent SQL type inference (BIGINT, FLOAT, NVARCHAR)

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
uv run python gui.py
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
python gui.py
```

## 📖 Usage Guide

### 1. Add Files to Queue
- Click **"Add Files"** button
- Select one or more CSV or Excel files
- Files appear in the queue list

### 2. Preview and Customize (Optional but Recommended!)
- Select a file from the queue
- Click **"Preview File"** button
- In the preview dialog:
  - View **statistics** (rows, columns, NULL counts)
  - See **first 20 rows** of data
  - **Edit column names** in the text fields
  - **Change data types** using dropdowns
  - For multi-sheet Excel files: use the sheet selector
  - Click **"Apply Changes"** to save customizations

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

## 🔧 Requirements

- **Windows 10/11**
- **SQL Server** with ODBC Driver 17 or later
- For source installation: Python 3.10+

### Installing ODBC Driver

Download from: [Microsoft ODBC Driver for SQL Server](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server)

## 🏗️ Building from Source

### Local Build

```bash
# Install PyInstaller
pip install pyinstaller

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

## 📁 Project Structure

```
file_to_database_table/
├── gui.py                      # Main GUI application
├── main.py                     # Core logic (database, file processing)
├── FileToDBConverter.spec      # PyInstaller build configuration
├── build_exe.bat               # Local build script
├── requirements.txt            # Python dependencies
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

## 🐛 Troubleshooting

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
- Large files may take time to load
- Files with many columns may need horizontal scrolling
- Corrupted files will show error messages

### Build Issues
- Ensure all dependencies are installed: `pip install -r requirements.txt`
- Clean previous builds: `rmdir /s /q build dist`
- Check `build\FileToDBConverter\warn-FileToDBConverter.txt` for warnings

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Commit changes: `git commit -m 'Add feature'`
4. Push to branch: `git push origin feature-name`
5. Open a Pull Request

## 📄 License

This project is open source and available under the MIT License.

## 🙏 Acknowledgments

Built with:
- **tkinter** - GUI framework
- **pandas** - Data processing
- **pyodbc** - SQL Server connectivity
- **openpyxl** - Excel file handling
- **cryptography** - Password encryption
- **PyInstaller** - Executable packaging

## 📞 Support

For issues, questions, or feature requests:
- Open an [Issue](../../issues)
- Check [RELEASE.md](RELEASE.md) for build documentation

---

**Made with ❤️ by kru5ty7**
