import sys
import os
import datetime

def validate_output_files(io_handler):
    generated_files = [f.lower() for f in os.listdir(io_handler.config['intermediate_folder']) if f.endswith('.csv')]
    missing = []
    for sheet in io_handler.config['sheets_to_process']:
        expected_file = f"{sheet}.csv".lower()
        if expected_file not in generated_files:
            missing.append(sheet)
    return missing
