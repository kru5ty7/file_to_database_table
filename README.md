# File to Database Table

This tool allows you to convert a CSV or Excel file into a database table in MS SQL Server.

## Prerequisites

- Python 3.x
- MS SQL Server
- ODBC Driver 17 for SQL Server

## Setup

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
    Open the `config.json` file and fill in your MS SQL Server details:
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

1.  **Run the script:**
    ```bash
    python main.py
    ```

2.  **Enter the file path:**
    When prompted, enter the full path to your CSV or Excel file.

The script will then create a table in your database with the same name as the file (without the extension) and insert the data from the file into the table.
