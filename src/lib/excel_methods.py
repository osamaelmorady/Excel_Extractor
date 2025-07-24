import os
import pandas as pd
import copy
import re


def excel_col_to_index(col: str) -> int:
    """Convert Excel column letter (e.g., 'A') to zero-based index."""
    col = col.upper()
    index = 0
    for c in col:
        index = index * 26 + (ord(c) - ord('A') + 1)
    return index - 1

def excel_row_to_index(row: str) -> int:
    """Convert Excel row number (e.g., '1') to zero-based index."""
    return int(row) - 1



class WorksheetHandler:
    """
    A class for processing Excel worksheet data using pandas DataFrame.
    """
    def __init__(self, io_handler, sheet_name, sheet_data, file_extension=".xlsx"):
        self.sheet_name = sheet_name
        self.data = pd.DataFrame(copy.deepcopy(sheet_data))
        self.file_extension = file_extension

        self.intermediate_folder_path = io_handler.get_intermediate_folder_path()
        self.output_folder_path = io_handler.get_output_folder_path() 
        

    def save_sheet_as_excel(self, destination="intermediate"):
        if destination == "output":
            output_folder = self.output_folder_path
        elif destination == "intermediate":
            output_folder = self.intermediate_folder_path
        else:
            raise ValueError(f"❌ Unknown destination: {destination}")

        if not output_folder:
            raise ValueError(f"Destination '{destination}' not configured in settings.yaml.")

        output_path = os.path.join(output_folder, f"intermediate_{self.sheet_name}{self.file_extension}")
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            self.data.to_excel(writer, index=False, sheet_name=self.sheet_name)
               
    def save_sheet_as_csv(self, destination="intermediate"):
        if destination == "output":
            output_folder = self.output_folder_path
        elif destination == "intermediate":
            output_folder = self.intermediate_folder_path
        else:
            raise ValueError(f"❌ Unknown destination: {destination}")

        if not output_folder:
            raise ValueError(f"Destination '{destination}' not configured in settings.yaml.")

        output_path = os.path.join(output_folder, f"intermediate_{self.sheet_name}.csv")
        self.data.to_csv(output_path, index=False)
        # with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        #     self.data.to_excel(writer, index=False, sheet_name=self.sheet_name)
            
            
    # --------------------- Add/remove/read methods ---------------------
    def read_region(self, start_point, end_point):   
        def parse_cell(cell):
            match = re.match(r"([A-Za-z]+)([0-9]+)", cell)
            if not match:
                raise ValueError(f"Invalid Excel cell format: '{cell}'")
            col_letters, row_number = match.groups()
            return int(row_number) - 1, excel_col_to_index(col_letters)  # row, col
    
        start_row, start_col = parse_cell(start_point)
        end_row, end_col = parse_cell(end_point)
    
        return self.data.iloc[start_row:end_row + 1, start_col:end_col + 1].copy(deep=True)
    
    
    # Add a new column at a given Excel column label (e.g., 'D')
    def add_column(self, col_letter, default_value=None):
        col_index = excel_col_to_index(col_letter)
        if col_index > len(self.data.columns):
            raise IndexError(f"Column '{col_letter}' is out of bounds.")

        new_col_name = f"Column_{col_letter}"
        self.data.insert(col_index, new_col_name, default_value)


    # Add a new row at a given Excel row number (e.g., '5')
    def add_row(self, row_number, default_value=None):
        row_index = excel_row_to_index(row_number)
        if row_index > len(self.data):
            raise IndexError(f"Row '{row_number}' is out of bounds.")

        new_row = pd.Series(default_value, index=self.data.columns)
        top = self.data.iloc[:row_index]
        bottom = self.data.iloc[row_index:]
        self.data = pd.concat([top, pd.DataFrame([new_row]), bottom], ignore_index=True)


    # Remove column by Excel letter, e.g., 'C'
    def remove_column(self, col_letter):     
        col_index = excel_col_to_index(col_letter)
        if col_index >= len(self.data.columns):
            raise IndexError(f"Column '{col_letter}' is out of bounds.")
        
        col_name = self.data.columns[col_index]
        self.data.drop(columns=[col_name], inplace=True)
    
    # Remove row by Excel number, e.g., '5'
    def remove_row(self, row_number):        
        row_index = excel_row_to_index(row_number)
        if 0 <= row_index < len(self.data):
            self.data.drop(index=row_index, inplace=True)
            self.data.reset_index(drop=True, inplace=True)
        else:
            raise IndexError(f"Row '{row_number}' is out of bounds.")


    # --------------------- Cell update methods ---------------------
    def set_cell_value(self, row_index, column_name, value=None):
        self.data.at[row_index, column_name] = value

    def get_cell_value(self, row_index, column_name):
        return self.data.at[row_index, column_name]

    def set_row_values(self, row_index, new_row):
        for col, val in new_row.items():
            self.data.at[row_index, col] = val

    def get_row_values(self, row_index):
        return self.data.iloc[row_index].to_dict()

    def set_column_values(self, column_name, new_values):
        if len(new_values) != len(self.data):
            raise ValueError("Length of new_values must match number of rows.")
        self.data[column_name] = new_values

    def get_column_values(self, column_name):
        return self.data[column_name].tolist()

    # --------------------- Manipulation methods ---------------------
    def trim_spaces_in_row(self, row_index):
        for col in self.data.columns:
            val = self.data.at[row_index, col]
            if isinstance(val, str):
                self.data.at[row_index, col] = val.strip()

    def trim_spaces_in_column_by_index(self, col_index):
        col_name = self.data.columns[col_index]
        self.data[col_name] = self.data[col_name].apply(lambda x: x.strip() if isinstance(x, str) else x)

    def find_text_in_column(self, column, search_text):
        return self.data[self.data[column].str.contains(search_text, case=False, na=False)]

    def find_text_in_row(self, row_index, search_text):
        return any(search_text.lower() in str(val).lower() for val in self.data.iloc[row_index])

    # --------------------- Transformations ---------------------
    def sort_rows(self, sort_key, reverse=False):
        self.data = self.data.sort_values(by=sort_key, ascending=not reverse).reset_index(drop=True)

    def sort_columns(self, column_order, reverse=False):
        if reverse:
            column_order = list(reversed(column_order))
        self.data = self.data[column_order]

    def filter_columns(self, columns_to_keep):
        self.data = self.data[columns_to_keep]

    def filter_rows(self, rows_to_keep):
        self.data = self.data.iloc[rows_to_keep].reset_index(drop=True)

    # --------------------- Lookup ---------------------
    def vlookup(self, lookup_df, data_key, lookup_key, result_columns):
        lookup_df = lookup_df.set_index(lookup_key)[result_columns]
        self.data = self.data.join(lookup_df, on=data_key)

    def compare_excel(self, df2, key_column):
        set1 = set(self.data[key_column])
        set2 = set(df2[key_column])
        return {
            'in_both': set1 & set2,
            'only_in_first': set1 - set2,
            'only_in_second': set2 - set1
        }

    # --------------------- Cleaning ---------------------
    def remove_empty_rows(self):
        self.data.dropna(how='all', inplace=True)
        self.data.reset_index(drop=True, inplace=True)

    def fill_missing_values(self, column, fill_value):
        self.data[column].fillna(fill_value, inplace=True)

    def convert_column_type(self, column, target_type):
        self.data[column] = pd.to_numeric(self.data[column], errors='coerce') if target_type in [int, float] else self.data[column].astype(str)

    # --------------------- Advanced ---------------------
    def group_by(self, group_column, aggregate_func='sum'):
        return self.data.groupby(group_column).agg(aggregate_func)

    def merge_excel(self, df2):
        self.data = pd.concat([self.data, df2], ignore_index=True)

    def pivot_table(self, index_col, columns_col, values_col, aggfunc='sum'):
        return pd.pivot_table(self.data, index=index_col, columns=columns_col, values=values_col, aggfunc=aggfunc)
