import sys
import os
import datetime

# Add the project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.logger import setup_logger
from src.io_handler import IOHandler
from src.parser import ExcelParser
from src.utils import *



def main():
    logger = setup_logger()
    io_handler = IOHandler()
    parser = ExcelParser(io_handler, logger)

    try:
        parser.parse_sheets()
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")
        raise

    # Check generated files
    missing = validate_output_files(io_handler)

    # Write a summary log
    # summary_path = os.path.join(io_handler.config['logs'], "generation_summary.txt")
    io_handler.write_genetation_summary(missing)

if __name__ == "__main__":
    main()
