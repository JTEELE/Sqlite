import sqlite3
import pandas as pd
import time
import xlwings as xw


database = 'Compensation'#input('Enter database: ')
conn = sqlite3.connect(f'data\\{database}.db')
cursor = conn.cursor()
search_string = input("Enter the search string: ")
tables_query = "SELECT name FROM sqlite_master WHERE type='table'"
cursor.execute(tables_query)
tables = cursor.fetchall()
matched_tables = {}
skipped_tables = {}
filtered_tables = {}
for table in tables:
    table_name = table[0]
    columns_query = f"PRAGMA table_info({table_name})"
    cursor.execute(columns_query)
    columns = [elem[1] for elem in cursor.fetchall() if 'index' not in elem]
    # if "SHKSG" in columns:
    #     print(table)
    for column in columns:
        query = f'SELECT * FROM {table_name} WHERE {column} LIKE ?'
        param = ('%' + search_string + '%',)
        try:
            cursor.execute(query, param)
            results = cursor.fetchall()
        except:
            print(table)
            skipped_tables[table_name] = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        if results:
            print(f"Found in table '{table_name}', column '{column}':")
            matched_tables[table_name] = pd.read_sql_query(f"SELECT * FROM {table_name} WHERE {column} LIKE '%{search_string}%'", conn)
            for row in results:
                print(row)
            # results = None
for table in skipped_tables:
    table_df = skipped_tables[table]
    search_results = table_df.apply(lambda x: x.astype(str).str.contains(search_string, case=False)).any(axis=1)
    filtered_df = table_df[search_results].copy()
    if len(filtered_df)>1 and table not in matched_tables.keys():
        matched_tables[table] = pd.read_sql_query(f"SELECT * FROM {table}", conn)
        filter_columns = filtered_df.columns[filtered_df.astype(str).apply(lambda x: x.str.contains(search_string, case=False)).any()]
        print(filter_columns.values)
        print(f"Found in table {table}, column {filter_columns.values}")
        for index, row in filtered_df.reset_index().iterrows():
            print(tuple(filtered_df.iloc[index]))
conn.close()
# excel_bool = input('Save source tables in Excel? (y/n): ') =='y'
# if excel_bool:
writer = pd.ExcelWriter(f'SEARCH\\{search_string}.xlsx', engine='xlsxwriter')
for source_df in matched_tables:
    matched_tables[source_df].copy().to_excel(writer, sheet_name=f'{source_df[:31]}')
writer.close()
time.sleep(2)
wb = xw.Book(f'SEARCH\\{search_string}.xlsx')
for ws in wb.sheets:
    ws.autofit(axis="columns")
