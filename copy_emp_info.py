import sqlite3

# Connect to both databases
source_conn = sqlite3.connect('employee.db')
target_conn = sqlite3.connect('Cluster_Expense.db')

try:
    # Get data from source database
    source_cursor = source_conn.cursor()
    source_cursor.execute('SELECT * FROM emp_info')
    data = source_cursor.fetchall()
    
    # Get column names from source table
    source_cursor.execute('PRAGMA table_info(emp_info)')
    columns = [column[1] for column in source_cursor.fetchall()]
    
    # Create table in target database if it doesn't exist
    column_definitions = ', '.join([f'{col} TEXT' for col in columns])
    target_conn.execute(f'CREATE TABLE IF NOT EXISTS emp_info ({column_definitions})')
    
    # Insert data into target database
    target_cursor = target_conn.cursor()
    placeholders = ','.join(['?' for _ in columns])
    target_cursor.executemany(f'INSERT INTO emp_info VALUES ({placeholders})', data)
    
    # Commit the changes
    target_conn.commit()
    print(f"Successfully copied {len(data)} records from employee.db to Cluster_Expense.db")

except sqlite3.Error as e:
    print(f"An error occurred: {e}")

finally:
    # Close the connections
    source_conn.close()
    target_conn.close()
