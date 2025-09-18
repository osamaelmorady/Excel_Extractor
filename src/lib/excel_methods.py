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

def parse_excel_cell(cell):
    """
    Parses Excel-style cell (e.g., 'B4') into (row_index, column_label).
    """
    match = re.fullmatch(r"([A-Za-z]+)([0-9]+)", cell.strip())
    if not match:
        raise ValueError(f"Invalid Excel cell format: '{cell}'")
    col_letters, row_letters = match.groups()
    return row_letters, col_letters.upper()




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
            
    #######################################################################        
    # --------------------- Add/remove/read methods --------------------- #
    #######################################################################
    def set_row_values(self, row_str, values):
        """ Example sheet.set_row_values("4", ["x", 1.1, "Enabled"]) """
        """Set full row using Excel-style row number, filling from left (first column)."""
        row_idx = excel_row_to_index(row_str)
        if len(values) > len(self.data.columns):
            raise ValueError("❌ Too many values for row length.")

        for i, value in enumerate(values):
            col = self.data.columns[i]
            self.data.at[row_idx, col] = value

    def get_row_values(self, row_str):
        """Get full row as list from left to right."""
        row_idx = excel_row_to_index(row_str)
        return self.data.iloc[row_idx].tolist()

    def set_column_values(self, col_str, values):
        """ Example sheet.set_column_values("B", [1, 2, 3, 4]) """
        """Set full column using Excel-style letter, top-down."""
        col_label = excel_col_to_index(col_str)
        if len(values) != len(self.data):
            raise ValueError("❌ Column length must match number of rows.")
        self.data[col_label] = values

    def get_column_values(self, col_str):
        """Get full column as list from top to bottom."""
        col_label = excel_col_to_index(col_str)
        return self.data[col_label].tolist()

    
    def get_region(self, start_point, end_point):
        """Example : self.get_region("A1", "C3")"""
        start_row, start_col = parse_excel_cell(start_point)
        end_row, end_col = parse_excel_cell(end_point)
        
        start_row_loc = excel_row_to_index(start_row)
        end_row_loc = excel_row_to_index(end_row)
        start_col_loc = excel_col_to_index(start_col)
        end_col_loc = excel_col_to_index(end_col)

        return self.data.iloc[start_row_loc:end_row_loc + 1, start_col_loc:end_col_loc + 1].copy(deep=True)

    
    def set_region(self, start_cell, end_cell, values_2d):
        """Example : self.set_region("A1", "C3", values_2d)"""
        start_row, start_col = parse_excel_cell(start_cell)
        end_row, end_col = parse_excel_cell(end_cell)
        
        start_row_loc = excel_row_to_index(start_row)
        end_row_loc = excel_row_to_index(end_row)
        start_col_loc = excel_col_to_index(start_col)
        end_col_loc = excel_col_to_index(end_col)
        
        for i, row_val in enumerate(range(start_row_loc, end_row_loc + 1)):
            for j, col_val in enumerate(range(start_col_loc, end_col_loc + 1)):
                self.data.iat[row_val, col_val] = values_2d[i][j]
    
    
    # Add a new column at a given Excel column label (e.g., 'D')
    def add_column(self, col_str, default_value=None, col_new_name=None):
        """Example : self.add_column("A", values_2d)"""
        col_index = excel_col_to_index(col_str)

        if col_index > len(self.data.columns):
            raise IndexError(f"Column '{col_str}' is out of bounds.")

        if col_new_name is None:
            col_new_name = f"{col_str}"  

        self.data.insert(col_index, col_new_name, default_value)

    # Add a new row at a given Excel row number (e.g., '5')
    def add_row(self, row_str, default_value=None):
        """Example : self.add_row("A1", "C3", values_2d)"""
        row_index = excel_row_to_index(row_str)

        if row_index > len(self.data):
            raise IndexError(f"Row '{row_str}' is out of bounds.")

        new_row = pd.Series(default_value, index=self.data.columns)
        top = self.data.iloc[:row_index]
        bottom = self.data.iloc[row_index:]

        # Maintain original index if needed
        self.data = pd.concat([top, pd.DataFrame([new_row]), bottom]).reset_index(drop=True)

    # Remove column by Excel letter, e.g., 'C'
    def remove_column(self, col_letter):     
        """Example : self.remove_column("A1", "C3", values_2d)"""
        col_index = excel_col_to_index(col_letter)
        if col_index >= len(self.data.columns):
            raise IndexError(f"Column '{col_letter}' is out of bounds.")
        
        col_name = self.data.columns[col_index]
        self.data.drop(columns=[col_name], inplace=True)
    
    # Remove row by Excel number, e.g., '5'
    def remove_row(self, row_str): 
        """Example : self.remove_row("A1", "C3", values_2d)"""       
        row_index = excel_row_to_index(row_str)
        if 0 <= row_index < len(self.data):
            self.data.drop(index=row_index, inplace=True)
            self.data.reset_index(drop=True, inplace=True)
        else:
            raise IndexError(f"Row '{row_str}' is out of bounds.")
        
    #######################################################################
    # --------------------- Cell update methods ---------------------
    #######################################################################
    def set_cell_value(self, cell, value=None):
        """Example : self.set_cell_value("B4", "ON")"""
        row_str, column_str = parse_excel_cell(cell)
        
        row_index = excel_row_to_index(row_str)
        col_index = excel_col_to_index(column_str)
        
        if row_index >= len(self.data) or col_index >= len(self.data.columns):
            raise IndexError(f"Row '{row_str}' or column '{column_str}' is out of bounds.")
        
        col_label = self.data.columns[col_index]
        
        # Convert column to object dtype if not already (to support string assignment)
        if self.data[col_label].dtype != 'object':
            self.data[col_label] = self.data[col_label].astype('object')

        self.data.at[row_index, col_label] = value
        
    def get_cell_value(self, cell):
        """Example : self.get_cell_value("B4")"""
        row_str, column_str = parse_excel_cell(cell)
        row_index = excel_row_to_index(row_str)
        col_index = excel_col_to_index(column_str)
        if row_index >= len(self.data) or col_index >= len(self.data.columns):
            raise IndexError(f"Row '{row_str}' or column '{column_str}' is out of bounds.")
        
        col_label = self.data.columns[col_index]
        
        return self.data.at[row_index, col_label]

    #######################################################################
    # --------------------- Manipulation methods ---------------------
    #######################################################################
    def trim_spaces_in_cell(self, cell):
        """Trim spaces from a specific cell if it's a string (e.g., 'B5')."""
        row_str, col_str = parse_excel_cell(cell)
        row_index = excel_row_to_index(row_str)
        col_name = self.data.columns[excel_col_to_index(col_str)]
        val = self.data.at[row_index, col_name]
        if isinstance(val, str):
            self.data.at[row_index, col_name] = val.strip()

    def trim_spaces_in_row(self, row_str):
        """Trims spaces in all string cells of a given Excel-style row (e.g., '5')."""
        row_index = excel_row_to_index(row_str)
        for col in self.data.columns:
            val = self.data.at[row_index, col]
            if isinstance(val, str):
                self.data.at[row_index, col] = val.strip()

    def trim_spaces_in_column(self, col_str):
        """Trims spaces in all string cells of a given Excel-style column (e.g., 'B')."""
        col_name = self.data.columns[excel_col_to_index(col_str)]
        self.data[col_name] = self.data[col_name].apply(lambda x: x.strip() if isinstance(x, str) else x)


    def find_text_in_cell(self, cell, search_text):
        """Check if a cell contains the given search text (case-insensitive)."""
        row_str, col_str = parse_excel_cell(cell)
        row_index = excel_row_to_index(row_str)
        col_name = self.data.columns[excel_col_to_index(col_str)]
        val = self.data.at[row_index, col_name]
        return isinstance(val, str) and search_text.lower() in val.lower()


    def find_text_in_column(self, col_str, search_text):
        """Finds rows where a given Excel-style column (e.g., 'C') contains search_text."""
        col_name = self.data.columns[excel_col_to_index(col_str)]
        return self.data[self.data[col_name].str.contains(search_text, case=False, na=False)]

    def find_text_in_row(self, row_str, search_text):
        """Checks if search_text exists in any cell of a given Excel-style row (e.g., '4')."""
        row_index = excel_row_to_index(row_str)
        return any(search_text.lower() in str(val).lower() for val in self.data.iloc[row_index])

    def find_text_in_region(self, start_cell, end_cell, search_text):
        """Search for text in a rectangular Excel-style region, case-insensitive.

        Example: find_text_in_region("B4", "D10", "error")
        Returns: True if found anywhere in the region
        """
        start_row, start_col = parse_excel_cell(start_cell)
        end_row, end_col = parse_excel_cell(end_cell)

        region = self.data.iloc[start_row:end_row + 1, start_col:end_col + 1]

        return region.astype(str).applymap(lambda x: search_text.lower() in x.lower()).any().any()

    #######################################################################
    # --------------------- Transformations ---------------------
    #######################################################################
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
