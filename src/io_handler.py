# src/io_handler.py
import pandas as pd
import yaml
import os
import datetime

class IOHandler:
    def __init__(self):
        with open("config/settings.yaml") as f:
            self.config = yaml.safe_load(f)

    def read_excel(self):
        if not os.path.exists(self.config['excel_input_path']):
            raise FileNotFoundError(f"Excel file not found: {self.config['excel_input_path']}")
        return pd.ExcelFile(self.config['excel_input_path'])

    def write_csv(self, df, sheet_name):
        output_path = f"{self.config['intermediate_folder']}/{sheet_name}.csv"
        df.to_csv(output_path, index=False)
        
    def write_genetation_summary(self, missing):
        summary_path = f"{self.config['summary_file']}"
        with open(summary_path, "w", encoding="utf-8") as summary_file:
            summary_file.write(f"Run Timestamp: {datetime.datetime.now()}\n\n")
            if not missing:
                message = "✅ Generation is done successfully."
                print(message)
                summary_file.write(message + "\n")
            else:
                message = f"❌ Error: Missing sheet output(s): {', '.join(missing)}"
                print(message)
                summary_file.write(message + "\n")
                raise RuntimeError(message)
        