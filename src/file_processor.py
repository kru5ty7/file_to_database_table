"""
File reading and processing functions
"""

import pandas as pd
import os
from .utils import sanitize_name, logger

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
