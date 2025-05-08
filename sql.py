import tkinter as tk
from tkinter import ttk
import sqlite3
import pandas as pd
import xlwings as xw
import datetime

class DatabaseManager:
    def __init__(self, db_path):
        self.conn = sqlite3.connect(db_path)
        self.cursor = self.conn.cursor()

    def fetch_tables(self):
        self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        return [table[0] for table in self.cursor.fetchall()]

    def fetch_fields(self, table_name):
        self.cursor.execute(f"PRAGMA table_info({table_name});")
        return [field[1] for field in self.cursor.fetchall()]

    def execute_query(self, query):
        self.cursor.execute(query)
        return self.cursor.fetchall(), [desc[0] for desc in self.cursor.description]

    def close(self):
        self.conn.close()

class Application:
    operators = ['E', 'GT', 'GTE', 'LTE', 'LE', 'LIKE', 'BETWEEN', 'IN']
    operator_mapping = {
        'E': '',
        'GT': '>',
        'GTE': '>=',
        'LTE': '<=',
        'LE': '<',
        'LIKE': '',
        'IN': '("","")',
        'BETWEEN': '< >'
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Database Table Selector")
        self.db_manager = DatabaseManager('data/finance.db')

        self.tables = self.db_manager.fetch_tables()
        self.user_fields = []
        self.field_vars = []
        self.field_filters = {}
        self.user_inputs = {}
        self.filter_window = None

        self.create_widgets()

    def create_widgets(self):
        self.root.grid_columnconfigure(0, weight=1)

        self.label_table = ttk.Label(self.root, text="Select a table:")
        self.label_table.grid(row=0, column=0, pady=10, padx=10, sticky='w')

        self.table_var = tk.StringVar()
        self.table_dropdown = ttk.Combobox(self.root, textvariable=self.table_var)
        self.table_dropdown['values'] = self.tables
        self.table_dropdown.grid(row=1, column=0, pady=10, padx=10, sticky='we')
        self.table_dropdown.bind("<<ComboboxSelected>>", self.show_fields)

        self.button_frame = ttk.Frame(self.root)
        self.button_frame.grid(row=2, column=0, pady=10, padx=10, sticky='we')

        self.button_reset = ttk.Button(self.button_frame, text="View/Reset", command=self.reset_fields)
        self.button_reset.grid(row=0, column=0, padx=10)

        self.button_run = ttk.Button(self.button_frame, text="Run Filtered Report", command=self.run_filtered_report)
        self.button_run.grid(row=0, column=1, padx=10)

        self.button_sql = ttk.Button(self.button_frame, text="Write SQL", command=self.open_sql_editor)
        self.button_sql.grid(row=0, column=2, padx=10)

        self.label_fields = ttk.Label(self.root, text="")
        self.label_fields.grid(row=3, column=0, pady=10, padx=10, sticky='w')

        self.fields_frame = ttk.Frame(self.root)
        self.fields_frame.grid(row=5, column=0, pady=10, padx=10, sticky='w')

        self.select_all_button = ttk.Button(self.root, text="Select All", command=self.select_all_fields)
        self.select_all_button.grid(row=4, column=0, pady=10, padx=10)

    def reset_fields(self):
        self.field_filters.clear()
        self.user_fields = []
        for field, var in self.field_vars:
            var.set(False)
        self.show_fields()

    def show_fields(self, event=None):
        selected_table = self.table_var.get()
        if selected_table:
            fields = self.db_manager.fetch_fields(selected_table)
            if 'index' in fields:
                fields.remove('index')

            self.label_fields.config(text=f"Select fields from {selected_table}:")

            for widget in self.fields_frame.winfo_children():
                widget.destroy()

            self.field_vars = []
            for index, field in enumerate(fields):
                var = tk.BooleanVar()
                chk = tk.Checkbutton(self.fields_frame, text=field, variable=var)
                chk.grid(row=index, column=0, sticky='w')
                self.field_vars.append((field, var))

                # Adding user input field
                entry = tk.Entry(self.fields_frame)
                entry.grid(row=index, column=1, padx=10)
                self.user_inputs[field] = entry

    def run_filtered_report(self):
        self.user_fields = [field for field, var in self.field_vars if var.get()]
        if not self.user_fields:
            tk.messagebox.showwarning("No Fields Selected", "Please select at least one field.")
            return

        selected_table = self.table_var.get()
        where_clause = self.build_where_clause()

        query = f"SELECT {', '.join(self.user_fields)} FROM {selected_table} {where_clause}"
        try:
            rows, columns = self.db_manager.execute_query(query)
        except sqlite3.OperationalError as e:
            tk.messagebox.showerror("SQL Error", str(e))
            print(f"Query Error:\n{query}")
            return

        dataframe = pd.DataFrame(rows, columns=columns)

        now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        treeview_window = tk.Toplevel(self.root)
        treeview_window.title(f"Filtered Report at {now}")

        tree = ttk.Treeview(treeview_window, columns=columns, show='headings')
        tree.pack(expand=True, fill='both')

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        for index, row in dataframe.iterrows():
            tree.insert("", "end", values=list(row))

        button_frame = ttk.Frame(treeview_window)
        button_frame.pack(side='bottom', fill='x', padx=10, pady=10)

        refresh_button = ttk.Button(button_frame, text="Refresh Tree", command=lambda: self.refresh_tree(tree, dataframe))
        refresh_button.pack(side='right', padx=5)

        export_button = ttk.Button(button_frame, text="Export to Excel", command=lambda: self.export_to_excel(dataframe, now))
        export_button.pack(side='right', padx=5)

        self.update_filter_window()

    def build_where_clause(self):
        if not self.field_filters:
            return ''
        clauses = []
        for field, (op, value) in self.field_filters.items():
            sql_op = self.operator_mapping.get(op, '=')
            if op == 'IN':
                clauses.append(f"{field} IN {value}")
            elif op == 'BETWEEN':
                val1, val2 = value.split(',')
                clauses.append(f"{field} BETWEEN {val1.strip()} AND {val2.strip()}")
            elif op == 'LIKE':
                clauses.append(f"{field} LIKE '{value}'")
            else:
                clauses.append(f"{field} {sql_op} '{value}'")
        return 'WHERE ' + ' AND '.join(clauses)

    def refresh_tree(self, tree, dataframe):
        for item in tree.get_children():
            tree.delete(item)
        for index, row in dataframe.iterrows():
            tree.insert("", "end", values=list(row))
        print("Tree refreshed")


    def export_to_excel(self, dataframe, now):
        file_path = f"Filtered_Report_{now}.xlsx"
        dataframe.to_excel(file_path, index=False)
        print(f"Exported to {file_path}")
        tk.messagebox.showinfo("Export Successful", f"Data exported to {file_path}")

    def apply_filter(self, column):
        filter_input_window = tk.Toplevel(self.root)
        filter_input_window.title(f"Filter {column}")
        filter_entries = {}

        ttk.Label(filter_input_window, text=f"Enter filter for {column}:").grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        for index, op in enumerate(self.operators):
            ttk.Label(filter_input_window, text=op).grid(row=index+1, column=0, padx=10, pady=5, sticky='e')
            entry = ttk.Entry(filter_input_window)
            previous_value = self.operator_mapping[op] # Prepopulate with the last input value if it exists
            if column in self.field_filters and self.field_filters[column][0] == op:
                previous_value = self.field_filters[column][1]
            entry.insert(0, previous_value)
            entry.grid(row=index+1, column=1, padx=10, pady=5)
            filter_entries[op] = entry  # Use 'op' as the key

        def apply_and_close():
            for op, entry in filter_entries.items():
                if entry.get() != self.operator_mapping[op]:
                    self.field_filters[column] = (op, entry.get())
                    break  # Assume only one operator per field
            else:
                # If no filter is set, remove the field from filters
                if column in self.field_filters:
                    del self.field_filters[column]
            filter_input_window.destroy()
            print(f"Current filters: {self.field_filters}")
            self.update_filter_window()

        # def reset_filters(): 
        #     for entry in filter_entries.values(): 
        #         entry.delete(0, tk.END) 
        #         entry.insert(0, "")
        #         apply_and_close()

        ttk.Button(filter_input_window, text="Apply", command=apply_and_close).grid(row=len(self.operators) + 1, column=1, padx=10, pady=10, sticky='e')
        # ttk.Button(filter_input_window, text="Reset Filters", command=reset_filters).grid(row=len(self.operators) + 1, column=0, columnspan=1, padx=10, pady=10, sticky='w')

    def update_filter_window(self):
        if not self.filter_window or not self.filter_window.winfo_exists():
            self.filter_window = tk.Toplevel(self.root)
            self.filter_window.title("Current Filters")
        else:
            for widget in self.filter_window.winfo_children():
                widget.destroy()

        ttk.Label(self.filter_window, text="Current Filters:").pack(padx=10, pady=10)

        for field, (op, value) in self.field_filters.items():
            ttk.Label(self.filter_window, text=f"{field} {op} {value}").pack(anchor='w', padx=10)

        # Filter buttons to apply or adjust filters
        filter_buttons_frame = ttk.Frame(self.filter_window)
        filter_buttons_frame.pack(pady=10)

        for field in self.user_fields:
            ttk.Button(filter_buttons_frame, text=f"Filter {field}", command=lambda c=field: self.apply_filter(c)).pack(side='left', padx=5)
        
        # Ensure the window remains on top
        self.filter_window.lift()
        self.filter_window.attributes('-topmost', True)
        self.filter_window.after_idle(self.filter_window.attributes, '-topmost', False)

    def select_all_fields(self):
        for _, var in self.field_vars:
            var.set(True)

    def open_sql_editor(self):
        sql_editor_window = tk.Toplevel(self.root)
        sql_editor_window.title("SQL Editor")

        sql_text = tk.Text(sql_editor_window, wrap='word', width=80, height=20)
        sql_text.pack(padx=10, pady=10)

        run_button = ttk.Button(sql_editor_window, text="Run SQL", command=lambda: self.run_sql(sql_text.get("1.0", tk.END)))
        run_button.pack(padx=10, pady=10)

    def run_sql(self, sql_code):
        try:
            rows, columns = self.db_manager.execute_query(sql_code)
            dataframe = pd.DataFrame(rows, columns=columns)
            now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")

            treeview_window = tk.Toplevel(self.root)
            treeview_window.title(f"SQL Query Result at {now}")

            tree = ttk.Treeview(treeview_window, columns=columns, show='headings')
            tree.pack(expand=True, fill='both')

            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=100)

            for index, row in dataframe.iterrows():
                tree.insert("", "end", values=list(row))

            export_button = ttk.Button(treeview_window, text="Export to Excel",
                                    command=lambda: self.export_to_excel(dataframe, now))
            export_button.pack(pady=10)

        except sqlite3.OperationalError as e:
            tk.messagebox.showerror("SQL Error", str(e))
            print(f"SQL Error:\n{sql_code}\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(root)
    root.mainloop()
