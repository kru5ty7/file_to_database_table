# File to Database Table

This tool allows you to convert a CSV or Excel file into a database table in MS SQL Server.

## Prerequisites

- Python 3.x
- MS SQL Server
- ODBC Driver 17 for SQL Server

## Setup

### Option 1: Using uv (Recommended)

1.  **Clone the repository:**
    ```bash
    git clone <repository_url>
    cd file_to_database_table
    ```

2.  **Install uv if you haven't already:**
    ```bash
    # On Windows
    powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

    # On macOS/Linux
    curl -LsSf https://astral.sh/uv/install.sh | sh
    ```

3.  **Install dependencies with uv:**
    ```bash
    uv sync
    ```

4.  **Configure the database connection:**
    Copy the template and fill in your MS SQL Server details:
    ```bash
    cp config.json.template config.json
    ```

    Edit `config.json` (supports multiple connections):
    ```json
    {
        "default_connection": "production",
        "connections": {
            "production": {
                "server": "your_server_name",
                "database": "your_database_name",
                "username": "your_username",
                "password": "your_password",
                "driver": "{ODBC Driver 17 for SQL Server}"
            },
            "development": {
                "server": "dev_server_name",
                "database": "dev_database_name",
                "username": "dev_username",
                "password": "dev_password",
                "driver": "{ODBC Driver 17 for SQL Server}"
            }
        }
    }
    ```

    **Legacy single-connection format is also supported:**
    ```json
    {
        "server": "your_server_name",
        "database": "your_database_name",
        "username": "your_username",
        "password": "your_password",
        "driver": "{ODBC Driver 17 for SQL Server}"
    }
    ```

### Option 2: Using pip

1.  **Clone the repository:**
    ```bash
    git clone <repository_url>
    cd file_to_database_table
    ```

2.  **Create a virtual environment:**
    ```bash
    python -m venv venv
    ```

3.  **Activate the virtual environment:**
    -   On Windows:
        ```bash
        .\venv\Scripts\activate
        ```
    -   On macOS/Linux:
        ```bash
        source venv/bin/activate
        ```

4.  **Install the required packages:**
    ```bash
    pip install -r requirements.txt
    ```

5.  **Configure the database connection:**
    Copy the template and fill in your MS SQL Server details:
    ```bash
    cp config.json.template config.json
    ```

    Edit `config.json`:
    ```json
    {
        "server": "your_server_name",
        "database": "your_database_name",
        "username": "your_username",
        "password": "your_password",
        "driver": "{ODBC Driver 17 for SQL Server}"
    }
    ```

## Usage

### Option 1: GUI (Graphical User Interface)

The GUI provides a modern, user-friendly interface for converting files to database tables.

**With uv:**
```bash
uv run python gui.py
```

**With traditional Python:**
```bash
python gui.py
```

**Features:**
- File browser for easy file selection
- **Multiple database connection support** - switch between different environments (production, development, test)
- Database connection testing
- Real-time progress tracking
- Live log output
- Support for CSV and Excel files (single or multiple sheets)

### Option 2: Command Line Interface

**With uv:**
```bash
uv run python main.py
```

**With traditional Python:**
```bash
python main.py
```

When prompted, enter the full path to your CSV or Excel file. The script will create tables in your database with sanitized names and insert all data.

## Features

- **Multi-sheet Excel support**: Automatically processes all sheets in an Excel workbook
- **Data type inference**: Intelligently detects numeric vs text columns
- **Leading zero preservation**: Values like "020435" are preserved as strings
- **Name sanitization**: Table and column names are sanitized for SQL compatibility
- **Comprehensive logging**: Detailed logs saved to `logs/` directory
- **Progress tracking**: Real-time progress updates during conversion
- **Error handling**: Detailed error messages and full tracebacks
- **Password encryption**: Database passwords are encrypted using Fernet symmetric encryption before being stored in config.json

## Security

Passwords in `config.json` are automatically encrypted using Fernet symmetric encryption (from the `cryptography` library). The encryption key is stored in `.encryption_key` file in the project directory.

**Important security notes:**
- Keep `.encryption_key` file secure and never commit it to version control
- The `.gitignore` file is configured to exclude both `.encryption_key` and `config.json`
- If you lose the `.encryption_key` file, you'll need to re-enter all passwords
- For production use, consider additional security measures like environment variables or a secrets manager
