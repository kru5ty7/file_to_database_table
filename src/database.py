"""
Database operations and connection management
"""

import pyodbc
import json
import pandas as pd
from .utils import decrypt_password, logger
from .file_processor import infer_column_type

def get_db_connection(connection_name=None):
    """Get database connection using config from config.json"""
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

def create_table_from_dataframe(df, table_name, cursor, column_name_map=None, column_type_map=None):
    """
    Create table from dataframe with optional column name and type overrides.

    Args:
        df: DataFrame to create table from
        table_name: Name of the table to create
        cursor: Database cursor
        column_name_map: Dict mapping original column names to new names (optional)
        column_type_map: Dict mapping column names to SQL types (optional)
    """
    logger.info(f"Creating table: {table_name}")

    # Drop table if it exists
    logger.debug(f"Checking if table '{table_name}' already exists")
    cursor.execute(f"IF OBJECT_ID('{table_name}', 'U') IS NOT NULL DROP TABLE {table_name}")

    # Initialize override maps if not provided
    if column_name_map is None:
        column_name_map = {}
    if column_type_map is None:
        column_type_map = {}

    # Analyze each column to determine the best type
    logger.info("Analyzing column types...")
    sql_columns = []
    column_types = {}
    final_column_names = {}  # Map original to final column names

    for column_name in df.columns:
        # Get final column name (use override if provided, otherwise use original)
        final_col_name = column_name_map.get(column_name, column_name)
        final_column_names[column_name] = final_col_name

        # Get column type (use override if provided, otherwise infer)
        if column_name in column_type_map:
            col_type = column_type_map[column_name]
            logger.debug(f"Using overridden type for '{column_name}': {col_type}")
        else:
            col_type = infer_column_type(df[column_name], column_name)

        column_types[column_name] = col_type
        sql_columns.append(f"[{final_col_name}] {col_type}")

        if final_col_name != column_name:
            logger.info(f"Column renamed: '{column_name}' -> '{final_col_name}'")

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
