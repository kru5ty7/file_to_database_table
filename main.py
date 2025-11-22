
import pandas as pd
import pyodbc
import json
import os
import re
import traceback
import logging
from datetime import datetime
from cryptography.fernet import Fernet
import base64

# Configure logging
def setup_logging():
    """Setup logging configuration with both file and console handlers"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    log_filename = os.path.join(log_dir, f"file_to_db_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

    # Create logger
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)

    # Clear any existing handlers
    logger.handlers = []

    # File handler - detailed logging
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s')
    file_handler.setFormatter(file_formatter)

    # Console handler - less verbose
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(levelname)s - %(message)s')
    console_handler.setFormatter(console_formatter)

    # Add handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logger.info(f"Logging initialized. Log file: {log_filename}")
    return logger

logger = setup_logging()

# Encryption key management
def get_or_create_key():
    """Get or create encryption key for password storage"""
    key_file = '.encryption_key'
    if os.path.exists(key_file):
        with open(key_file, 'rb') as f:
            key = f.read()
    else:
        key = Fernet.generate_key()
        with open(key_file, 'wb') as f:
            f.write(key)
        logger.info("Generated new encryption key")
    return key

def encrypt_password(password):
    """Encrypt password using Fernet symmetric encryption"""
    if not password:
        return ""
    key = get_or_create_key()
    f = Fernet(key)
    encrypted = f.encrypt(password.encode())
    return base64.urlsafe_b64encode(encrypted).decode()

def decrypt_password(encrypted_password):
    """Decrypt password using Fernet symmetric encryption"""
    if not encrypted_password:
        return ""
    try:
        key = get_or_create_key()
        f = Fernet(key)
        decoded = base64.urlsafe_b64decode(encrypted_password.encode())
        decrypted = f.decrypt(decoded)
        return decrypted.decode()
    except Exception as e:
        logger.error(f"Failed to decrypt password: {e}")
        # If decryption fails, assume it's plain text (backward compatibility)
        return encrypted_password

def sanitize_name(name):
    """
    Sanitize table and column names:
    - Convert to lowercase
    - Remove special characters (keep only alphanumeric and underscores)
    - Replace whitespace with underscores
    - Remove leading/trailing underscores
    - Ensure it doesn't start with a number
    """
    original_name = name
    # Convert to lowercase
    name = name.lower()
    # Replace whitespace with underscores
    name = re.sub(r'\s+', '_', name)
    # Remove special characters, keep only alphanumeric and underscores
    name = re.sub(r'[^a-z0-9_]', '', name)
    # Remove leading/trailing underscores
    name = name.strip('_')
    # Ensure it doesn't start with a number
    if name and name[0].isdigit():
        name = 'col_' + name
    # Ensure it's not empty
    if not name:
        name = 'unnamed'

    if original_name != name:
        logger.debug(f"Sanitized name: '{original_name}' -> '{name}'")

    return name

def get_db_connection(connection_name=None):
    logger.info("Connecting to database...")
    try:
        with open('config.json') as config_file:
            config = json.load(config_file)

        # Support both old and new config formats
        if 'connections' in config:
            # New format with multiple connections
            if connection_name is None:
                connection_name = config.get('default_connection', list(config['connections'].keys())[0])

            if connection_name not in config['connections']:
                raise ValueError(f"Connection '{connection_name}' not found in config.json")

            db_config = config['connections'][connection_name]
            logger.info(f"Using connection: {connection_name}")
        else:
            # Old format with single connection
            db_config = config
            logger.debug("Using legacy config format (single connection)")

        logger.debug(f"Database server: {db_config['server']}, Database: {db_config['database']}")

        # Decrypt password if encrypted
        password = decrypt_password(db_config.get("password", ""))

        conn_str = (
            f'DRIVER={db_config["driver"]};'
            f'SERVER={db_config["server"]};'
            f'DATABASE={db_config["database"]};'
            f'UID={db_config["username"]};'
            f'PWD={password};'
        )
        conn = pyodbc.connect(conn_str)
        logger.info("Database connection established successfully")
        return conn
    except FileNotFoundError:
        logger.error("config.json file not found")
        raise
    except Exception as e:
        logger.error(f"Failed to connect to database: {e}")
        raise

def get_available_connections():
    """Get list of available connection names from config"""
    try:
        with open('config.json') as config_file:
            config = json.load(config_file)

        if 'connections' in config:
            return list(config['connections'].keys())
        else:
            return ['default']
    except FileNotFoundError:
        return []

def get_dataframes(file_path):
    """
    Read file and return a dictionary of dataframes.
    For CSV: returns {'sheet1': dataframe}
    For Excel: returns {sheet_name: dataframe} for each sheet
    """
    logger.info(f"Reading file: {file_path}")
    _, file_extension = os.path.splitext(file_path)
    dataframes = {}

    if file_extension.lower() == '.csv':
        logger.debug("File type: CSV")
        # Read CSV with all columns as strings to preserve formatting
        df = pd.read_csv(file_path, dtype=str, keep_default_na=False)
        logger.info(f"CSV loaded: {len(df)} rows, {len(df.columns)} columns")
        # Replace empty strings with NaN for proper NULL handling
        df = df.replace('', pd.NA)
        # Sanitize column names
        df.columns = [sanitize_name(col) for col in df.columns]
        dataframes['sheet1'] = df

    elif file_extension.lower() in ['.xls', '.xlsx']:
        logger.debug("File type: Excel")
        # Read all sheets from Excel file
        excel_file = pd.ExcelFile(file_path)
        logger.info(f"Found {len(excel_file.sheet_names)} sheet(s): {excel_file.sheet_names}")
        for sheet_name in excel_file.sheet_names:
            logger.debug(f"Reading sheet: {sheet_name}")
            # Read each sheet with all columns as strings to preserve leading zeros and formatting
            df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
            logger.info(f"Sheet '{sheet_name}' loaded: {len(df)} rows, {len(df.columns)} columns")
            # Replace empty strings with NaN for proper NULL handling
            df = df.replace('', pd.NA)
            # Sanitize column names
            df.columns = [sanitize_name(col) for col in df.columns]
            # Use sanitized sheet name as key
            dataframes[sanitize_name(sheet_name)] = df
    else:
        logger.error(f"Unsupported file type: {file_extension}")
        raise ValueError(f"Unsupported file type: {file_extension}")

    return dataframes

def infer_column_type(series, column_name):
    """
    Infer the best SQL column type for a series by analyzing its values.
    Returns the SQL type as a string.
    """
    # Remove NA values for analysis
    non_null = series.dropna()

    if len(non_null) == 0:
        logger.debug(f"Column '{column_name}': All NULL values, using NVARCHAR(MAX)")
        return "NVARCHAR(MAX)"

    # Check if all values are numeric (and don't have leading zeros)
    all_numeric = True
    has_decimals = False
    has_leading_zeros = False

    for val in non_null:
        val_str = str(val).strip()

        # Check for leading zeros (except single "0")
        if val_str.startswith('0') and len(val_str) > 1 and val_str[1].isdigit():
            has_leading_zeros = True
            all_numeric = False
            logger.debug(f"Column '{column_name}': Leading zeros detected (e.g., '{val_str}'), using NVARCHAR(MAX)")
            break

        # Try to convert to number
        try:
            float(val_str)
            if '.' in val_str or 'e' in val_str.lower() or 'E' in val_str:
                has_decimals = True
        except (ValueError, TypeError):
            all_numeric = False
            break

    # Determine the appropriate type
    if has_leading_zeros or not all_numeric:
        logger.debug(f"Column '{column_name}': Non-numeric data detected, using NVARCHAR(MAX)")
        return "NVARCHAR(MAX)"
    elif has_decimals:
        logger.debug(f"Column '{column_name}': Decimal values detected, using FLOAT")
        return "FLOAT"
    else:
        # Check if values fit in BIGINT range
        try:
            max_val = max(int(float(str(v))) for v in non_null)
            min_val = min(int(float(str(v))) for v in non_null)
            if -9223372036854775808 <= min_val and max_val <= 9223372036854775807:
                logger.debug(f"Column '{column_name}': Integer values detected, using BIGINT")
                return "BIGINT"
            else:
                logger.debug(f"Column '{column_name}': Values exceed BIGINT range, using NVARCHAR(MAX)")
                return "NVARCHAR(MAX)"
        except:
            logger.debug(f"Column '{column_name}': Error analyzing numeric range, using NVARCHAR(MAX)")
            return "NVARCHAR(MAX)"

def create_table_from_dataframe(df, table_name, cursor):
    logger.info(f"Creating table: {table_name}")

    # Drop table if it exists
    logger.debug(f"Checking if table '{table_name}' already exists")
    cursor.execute(f"IF OBJECT_ID('{table_name}', 'U') IS NOT NULL DROP TABLE {table_name}")

    # Analyze each column to determine the best type
    logger.info("Analyzing column types...")
    sql_columns = []
    column_types = {}

    for column_name in df.columns:
        col_type = infer_column_type(df[column_name], column_name)
        column_types[column_name] = col_type
        sql_columns.append(f"[{column_name}] {col_type}")

    create_table_sql = f"CREATE TABLE {table_name} ({', '.join(sql_columns)})"
    logger.info(f"Table schema: {create_table_sql}")
    cursor.execute(create_table_sql)
    logger.info(f"Table '{table_name}' created successfully")

    # Insert data with proper type handling
    total_rows = len(df)
    logger.info(f"Inserting {total_rows} rows...")

    for idx, row in df.iterrows():
        values = []
        for col_name, v in zip(df.columns, row.values):
            if pd.isna(v):
                values.append("NULL")
            elif column_types[col_name] in ["BIGINT", "FLOAT"]:
                # Insert numeric values without quotes
                values.append(str(v))
            else:
                # Insert strings with proper escaping
                escaped_value = str(v).replace("'", "''")
                values.append(f"'{escaped_value}'")
        insert_sql = f"INSERT INTO {table_name} VALUES ({', '.join(values)})"

        try:
            cursor.execute(insert_sql)
            if (idx + 1) % 100 == 0:
                logger.debug(f"Inserted {idx + 1}/{total_rows} rows")
        except Exception as e:
            logger.error(f"Failed to insert row {idx + 1}: {e}")
            logger.debug(f"SQL: {insert_sql}")
            raise

    cursor.commit()
    logger.info(f"Successfully inserted all {total_rows} rows and committed transaction")


def main():
    logger.info("=" * 60)
    logger.info("File to Database Table Converter - Starting")
    logger.info("=" * 60)

    file_path = input("Enter the path to your file (CSV or Excel): ")

    if not os.path.exists(file_path):
        logger.error(f"File not found: {file_path}")
        print("File not found.")
        return

    # Get base table name from the file name
    base_table_name = sanitize_name(os.path.splitext(os.path.basename(file_path))[0])
    logger.info(f"Base table name: {base_table_name}")

    try:
        dataframes = get_dataframes(file_path)
        conn = get_db_connection()
        cursor = conn.cursor()

        # Process each sheet/dataframe
        logger.info(f"Processing {len(dataframes)} sheet(s)")
        for sheet_name, df in dataframes.items():
            # For single sheet (CSV or single Excel sheet), use base name
            # For multiple sheets, append sheet name to base name
            if len(dataframes) == 1:
                table_name = base_table_name
            else:
                table_name = f"{base_table_name}_{sheet_name}"

            logger.info(f"\n{'='*60}")
            logger.info(f"Processing sheet: {sheet_name}")
            logger.info(f"Target table: {table_name}")
            logger.info(f"{'='*60}")

            create_table_from_dataframe(df, table_name, cursor)

            logger.info(f"✓ Table '{table_name}' completed successfully")

        cursor.close()
        conn.close()
        logger.info("Database connection closed")

        logger.info(f"\n{'='*60}")
        logger.info(f"✓ All tables created successfully! Total: {len(dataframes)}")
        logger.info(f"{'='*60}")
        print(f"\n✓ All tables created successfully! Total: {len(dataframes)}")

    except Exception as e:
        logger.error(f"An error occurred: {e}", exc_info=True)
        print(f"An error occurred: {e}")
        print("\nFull traceback:")
        traceback.print_exc()
        logger.info("Process terminated with errors")

if __name__ == "__main__":
    main()
