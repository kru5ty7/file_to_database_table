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

### With uv:
```bash
uv run python main.py
```

### With traditional Python:
```bash
python main.py
```

2.  **Enter the file path:**
    When prompted, enter the full path to your CSV or Excel file.

The script will then create a table in your database with the same name as the file (without the extension) and insert the data from the file into the table.
