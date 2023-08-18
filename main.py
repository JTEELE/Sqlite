import glob
import pandas as pd
import sqlite3
import re
NAME =  input('Name this database: ')
FOLDER = input('Upload multiple files? (y/n): ') =='y'
FILE_TYPE = "." + input("Enter File Type (xlsx,csv,dat): ")
DIRECTORY = input("Paste entire folder path: ")
EXCEL_FILES = glob.glob(rf'{DIRECTORY}' + f'/*{FILE_TYPE}')
DROP_COL = ['index','Unnamed: 0','0']
if len(EXCEL_FILES) > 0:
    CONN = sqlite3.connect(f"DB\\{NAME}.db")
    for item in EXCEL_FILES:
        print(item)
    ERROR = False
else:
    ERROR = True
    
class SQLite:
    def __init__(
    self,
    name: str,
    conn: str,
    folder: str,
    directory: str,
    excel_files: str,
    file_type: str,
    drop_cols: list
) -> None:
        self.name = name
        self.conn = conn
        self.folder = folder
        self.directory = directory
        self.file_type = file_type 
        self.excel_files = excel_files
        self.drop_cols = drop_cols

    def pandas_sheets_to_sqlite(self,file_name):
        workbook = pd.read_excel(f'{self.directory}\\{file_name}{self.file_type}', sheet_name=None)
        for sheet_name, df in workbook.items():
            file_name = remove_special_characters(sheet_name)
            file_str = file_name.strip().replace(" ","_")
            df.to_sql(name=file_str, con=self.conn, if_exists='replace')

    def directory_to_db(self):
        print(f'uploading:')
        print(f'{self.excel_files}')
        c = self.conn.cursor()
        csv_data = {}
        for file in self.excel_files:
            if self.file_type == '.csv':
                df = pd.read_csv(file, encoding='latin-1')
            else:
                df = pd.read_excel(file)
            file_name = file.split(f"{self.directory}"+"\\")[1].replace(f'{self.file_type}',"")
            # file_name=file_name.split("-")[1]
            file_name = remove_special_characters(file_name)
            file_name = file_name.strip().replace(" ","_")
            file_name = file_name.replace("\\","")
            # file_name = file_name.replace(self.file_type,"")
            csv_data[file_name] = df.copy()
            cols = [elem for elem in df.columns if elem not in self.drop_cols]
            df[cols].copy().to_sql(name=file_name, con=self.conn, if_exists='replace')
        tables = list(csv_data.keys())
        print(f'{self.name} Tables:')
        print(tables)


def remove_special_characters(string):
    pattern = r"[^\w\s]"
    cleaned_string = re.sub(pattern, "", string)
    return cleaned_string



def main(NAME,CONN,FOLDER,DIRECTORY,EXCEL_FILES,FILE_TYPE,DROP_COL):
    sqlite_class = SQLite(NAME,CONN,FOLDER,DIRECTORY,EXCEL_FILES,FILE_TYPE,DROP_COL)
    if not FOLDER:
        file_name = input("Enter File Name: ")
        sqlite_class.pandas_sheets_to_sqlite(file_name)
    else:
        sqlite_class.directory_to_db()
    cursor = CONN.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    for table in tables:
        print(table[0])
    cursor.close()
    
if __name__ == "__main__":
    print(f"Folder Selection: {FOLDER}")
    if not ERROR:
        main(NAME,CONN,FOLDER,DIRECTORY,EXCEL_FILES,FILE_TYPE,DROP_COL)