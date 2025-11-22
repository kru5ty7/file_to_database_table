
import pandas as pd
import pyodbc
import json
import os
import re
import traceback

def sanitize_name(name):
    """
    Sanitize table and column names:
    - Convert to lowercase
    - Remove special characters (keep only alphanumeric and underscores)
    - Replace whitespace with underscores
    - Remove leading/trailing underscores
    - Ensure it doesn't start with a number
    """
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
    return name

def get_db_connection():
    with open('config.json') as config_file:
        config = json.load(config_file)

    conn_str = (
        f'DRIVER={config["driver"]};'
        f'SERVER={config["server"]};'
        f'DATABASE={config["database"]};'
        f'UID={config["username"]};'
        f'PWD={config["password"]};'
    )
    return pyodbc.connect(conn_str)

def get_dataframes(file_path):
    """
    Read file and return a dictionary of dataframes.
    For CSV: returns {'sheet1': dataframe}
    For Excel: returns {sheet_name: dataframe} for each sheet
    """
    _, file_extension = os.path.splitext(file_path)
    dataframes = {}

    if file_extension.lower() == '.csv':
        # Read CSV with all columns as strings to preserve formatting
        df = pd.read_csv(file_path, dtype=str, keep_default_na=False)
        # Replace empty strings with NaN for proper NULL handling
        df = df.replace('', pd.NA)
        # Sanitize column names
        df.columns = [sanitize_name(col) for col in df.columns]
        dataframes['sheet1'] = df

    elif file_extension.lower() in ['.xls', '.xlsx']:
        # Read all sheets from Excel file
        excel_file = pd.ExcelFile(file_path)
        for sheet_name in excel_file.sheet_names:
            # Read each sheet with all columns as strings to preserve leading zeros and formatting
            df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
            # Replace empty strings with NaN for proper NULL handling
            df = df.replace('', pd.NA)
            # Sanitize column names
            df.columns = [sanitize_name(col) for col in df.columns]
            # Use sanitized sheet name as key
            dataframes[sanitize_name(sheet_name)] = df
    else:
        raise ValueError(f"Unsupported file type: {file_extension}")

    return dataframes

def infer_column_type(series):
    """
    Infer the best SQL column type for a series by analyzing its values.
    Returns the SQL type as a string.
    """
    # Remove NA values for analysis
    non_null = series.dropna()

    if len(non_null) == 0:
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
        return "NVARCHAR(MAX)"
    elif has_decimals:
        return "FLOAT"
    else:
        # Check if values fit in BIGINT range
        try:
            max_val = max(int(float(str(v))) for v in non_null)
            min_val = min(int(float(str(v))) for v in non_null)
            if -9223372036854775808 <= min_val and max_val <= 9223372036854775807:
                return "BIGINT"
            else:
                return "NVARCHAR(MAX)"
        except:
            return "NVARCHAR(MAX)"

def create_table_from_dataframe(df, table_name, cursor):
    # Drop table if it exists
    cursor.execute(f"IF OBJECT_ID('{table_name}', 'U') IS NOT NULL DROP TABLE {table_name}")

    # Analyze each column to determine the best type
    sql_columns = []
    column_types = {}

    for column_name in df.columns:
        col_type = infer_column_type(df[column_name])
        column_types[column_name] = col_type
        sql_columns.append(f"[{column_name}] {col_type}")

    create_table_sql = f"CREATE TABLE {table_name} ({', '.join(sql_columns)})"
    print(create_table_sql)
    cursor.execute(create_table_sql)

    # Insert data with proper type handling
    for _, row in df.iterrows():
        values = []
        for col_name, v in zip(df.columns, row.values):
            if pd.isna(v):
                values.append("NULL")
            elif column_types[col_name] in ["BIGINT", "FLOAT"]:
                # Insert numeric values without quotes
                values.append(str(v))
            else:
                # Insert strings with proper escaping
                values.append(f"'{str(v).replace("'", "''")}'")
        insert_sql = f"INSERT INTO {table_name} VALUES ({', '.join(values)})"
        cursor.execute(insert_sql)

    cursor.commit()


def main():
    file_path = input("Enter the path to your file (CSV or Excel): ")
    if not os.path.exists(file_path):
        print("File not found.")
        return

    # Get base table name from the file name
    base_table_name = sanitize_name(os.path.splitext(os.path.basename(file_path))[0])

    try:
        dataframes = get_dataframes(file_path)
        conn = get_db_connection()
        cursor = conn.cursor()

        # Process each sheet/dataframe
        for sheet_name, df in dataframes.items():
            # For single sheet (CSV or single Excel sheet), use base name
            # For multiple sheets, append sheet name to base name
            if len(dataframes) == 1:
                table_name = base_table_name
            else:
                table_name = f"{base_table_name}_{sheet_name}"

            print(f"\nProcessing sheet: {sheet_name}")
            print(f"Creating table: {table_name}")

            create_table_from_dataframe(df, table_name, cursor)

            print(f"Table '{table_name}' created and data inserted successfully.")

        cursor.close()
        conn.close()

        print(f"\n✓ All tables created successfully! Total: {len(dataframes)}")

    except Exception as e:
        print(f"An error occurred: {e}")
        print("\nFull traceback:")
        traceback.print_exc()

if __name__ == "__main__":
    main()
