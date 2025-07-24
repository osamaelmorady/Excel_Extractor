import os
import pandas as pd
from lib.io_handler import IOHandler

class ExcelHandler:
    #### Constructor
    #######################
    def __init__(self, io_handler):
        super().__init__()  # Loads YAML and sets paths
        self.workbook_path = io_handler.get_input_file_info()
        self.intermediate_folder_path = io_handler.get_intermediate_folder_path()
        self.output_folder_path = io_handler.get_output_folder_path()
        self.sheets_info= io_handler.get_sheets_info()
        
    #### Workbook methods
    #######################
    def read_excel_workbook(self):
        """Reads the Excel workbook as a whole (not yet parsed)."""
        return pd.ExcelFile(self.workbook_path)

    def write_workbook(self, sheet_dict, destination="output"):
        """Writes multiple DataFrames to a single Excel workbook (intermediate folder by default)."""
        output_filename="Generated_workbook.xlsx"
        if destination == "output":
            output_folder = self.output_folder_path
        elif destination == "intermediate":
            output_folder = self.intermediate_folder_path
        else:
            raise ValueError(f"❌ Unknown destination: {destination}")    
        
        output_path = os.path.join(output_folder, output_filename)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in sheet_dict.items():
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)


    #### worksheets methods
    ##########################
    def read_excel_sheet(self, sheet_id):
        """Reads a single sheet from the Excel workbook using sheet ID and config."""
        if sheet_id not in self.sheets_info:
            raise KeyError(f"❌ Sheet ID '{sheet_id}' not found in sheet configuration.")

        sheet_cfg = self.sheets_info[sheet_id]
        sheet_name = sheet_cfg["name"]
        header_row = sheet_cfg.get("header", 0)
        skiprows = sheet_cfg.get("skiprows", None)  # optional
        
        try:
            return pd.read_excel(
                self.workbook_path,
                sheet_name=sheet_name,
                header=header_row,
                skiprows=skiprows
            )
        except ValueError as e:
            raise ValueError(f"❌ Sheet '{sheet_name}' not found in the Excel file.") from e


    def write_excel_sheet(self, df, sheet_id, destination="output"):
        """Writes a single DataFrame to an Excel file using sheet ID for naming."""
        if sheet_id not in self.sheets_info:
            raise KeyError(f"❌ Sheet ID '{sheet_id}' not found in sheet configuration.")

        sheet_name = self.sheets_info[sheet_id]["name"]

        if destination == "output":
            output_folder = self.output_folder_path
        elif destination == "intermediate":
            output_folder = self.intermediate_folder_path
        else:
            raise ValueError(f"❌ Unknown destination: {destination}")

        output_file = os.path.join(output_folder, f"Generated_{sheet_id}.xlsx")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
       
            
    def read_all_input_sheets(self):
        """Reads Excel sheets defined in sheets_info and returns them as a dictionary keyed by ID."""
        excel = pd.ExcelFile(self.workbook_path)
        sheet_dict = {}

        for sheet_id, sheet_config in self.sheets_info.items():
            sheet_name = sheet_config["name"]
            header_row = sheet_config.get("header", 0)

            if sheet_name in excel.sheet_names:
                df = pd.read_excel(self.workbook_path, sheet_name=sheet_name, header=header_row)
                sheet_dict[sheet_id] = df
            else:
                msg = f"⚠️ Sheet '{sheet_name}' not found in workbook."
                if hasattr(self, "logger"):
                    self.logger.warning(msg)
                else:
                    print(msg)
                    
        return sheet_dict


