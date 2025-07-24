# src/io_handler.py
import sys
import pandas as pd
import yaml
import os
import datetime
from pathlib import Path

class IOHandler:
    def __init__(self):
        config_path =sys.argv[1]
        
        if len(sys.argv) < 2:
            print("❌ Usage: python main.py <config/settings.yaml>")
            sys.exit(1)
            
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"❌ YAML config file not found: {config_path}")        
        
        with open(config_path) as f:
            self.config_yaml = yaml.safe_load(f)
            self.input_path = self.get_input_folder_path()
            self.intermediate_path = self.get_intermediate_folder_path()
            self.output_path = self.get_output_folder_path()
            self.summary_path = self.get_summary_file_path()
            self.log_file_path = self.get_log_file_path()

    def get_input_file_info(self):
        """Returns full path to the single Excel input file."""
        file_path = os.path.join(
            self.config_yaml["input_folder"],
            self.config_yaml["input_files_to_process"]["excel_workbook"]["name"]
        )
        if not os.path.exists(file_path):
            raise FileNotFoundError(f" ❌ Excel input file not found: {file_path}")
        return file_path
    
    def get_sheets_info(self):
        """Returns sheet configuration dictionary, with guard."""
        try:
            worksheets = self.config_yaml["input_files_to_process"]["excel_worksheets"]
            if not isinstance(worksheets, dict):
                raise TypeError("❌ 'excel_worksheets' must be a dictionary.")
            return worksheets
        except KeyError:
            raise KeyError("❌ 'excel_worksheets' not found in configuration.")
            
    def get_input_folder_path(self):
        path = self.config_yaml["input_folder"]
        if not os.path.exists(path):
            raise FileNotFoundError(f" ❌ Input folder does not exist: {path}")
        return path

    def get_intermediate_folder_path(self):
        path = self.config_yaml['intermediate_folder']
        if not os.path.exists(path):
            raise FileNotFoundError(f" ❌Intermediate folder does not exist: {path}")
        return path

    def get_output_folder_path(self):
        path = self.config_yaml['output_folder']
        if not os.path.exists(path):
            raise FileNotFoundError(f" ❌ Output folder does not exist: {path}")
        return path

    def get_summary_file_path(self):
        """Returns the summary file path (creates folder if needed)."""
        summary_path = Path(self.config_yaml['generation_summary_file'])
        summary_path.parent.mkdir(parents=True, exist_ok=True)
        return summary_path  
    
    def get_log_file_path(self):
        """Returns the log file path (creates folder if needed)."""
        log_path = Path(self.config_yaml['log_file'])
        log_path.parent.mkdir(parents=True, exist_ok=True)
        return log_path 
    



        