
import pandas as pd
import pyodbc
import json
import os

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

def get_dataframe(file_path):
    _, file_extension = os.path.splitext(file_path)
    if file_extension.lower() == '.csv':
        return pd.read_csv(file_path)
    elif file_extension.lower() in ['.xls', '.xlsx']:
        return pd.read_excel(file_path)
    else:
        raise ValueError(f"Unsupported file type: {file_extension}")

def create_table_from_dataframe(df, table_name, cursor):
    # Drop table if it exists
    cursor.execute(f"IF OBJECT_ID('{table_name}', 'U') IS NOT NULL DROP TABLE {table_name}")

    # Create table
    sql_columns = []
    for column_name, dtype in df.dtypes.items():
        if "object" in str(dtype):
            sql_columns.append(f"[{column_name}] NVARCHAR(MAX)")
        elif "int" in str(dtype):
            sql_columns.append(f"[{column_name}] INT")
        elif "float" in str(dtype):
            sql_columns.append(f"[{column_name}] FLOAT")
        elif "datetime" in str(dtype):
            sql_columns.append(f"[{column_name}] DATETIME")
        else:
            sql_columns.append(f"[{column_name}] NVARCHAR(MAX)")

    create_table_sql = f"CREATE TABLE {table_name} ({', '.join(sql_columns)})"
    cursor.execute(create_table_sql)

    # Insert data
    for index, row in df.iterrows():
        values = ', '.join([f"""'{str(v).replace("'", "''")}'""" if v is not None else "NULL" for v in row.values])
        insert_sql = f"INSERT INTO {table_name} VALUES ({values})"
        cursor.execute(insert_sql)

    cursor.commit()


def main():
    file_path = input("Enter the path to your file (CSV or Excel): ")
    if not os.path.exists(file_path):
        print("File not found.")
        return

    table_name = os.path.splitext(os.path.basename(file_path))[0]
    
    try:
        df = get_dataframe(file_path)
        conn = get_db_connection()
        cursor = conn.cursor()
        
        create_table_from_dataframe(df, table_name, cursor)
        
        print(f"Table '{table_name}' created and data inserted successfully.")
        
        cursor.close()
        conn.close()

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
