import sqlite3

# Connect to the SQLite database (or create one if it doesn't exist)
connection = sqlite3.connect('cc_model.db')

# Create a cursor object to execute SQL commands
cursor = connection.cursor()

# SQL command to rename the table
old_table_name = "Sheet1"  # Replace with your old table name
new_table_name = "main_detail"    # Replace with your new table name

try:
    cursor.execute(f"ALTER TABLE {old_table_name} RENAME TO {new_table_name};")
    print(f"Table renamed from '{old_table_name}' to '{new_table_name}'.")
except sqlite3.Error as e:
    print(f"An error occurred: {e}")

# Commit the changes and close the connection
connection.commit()
connection.close()
