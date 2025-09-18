###############################################################################################
######################################## SYS LIB ##############################################
###############################################################################################
import sys
import os
import datetime
import pandas as pd
from pathlib import Path
import pprint

# # Add the project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))


###############################################################################################
######################################## USER LIB #############################################
###############################################################################################
# Import proper libraries
from lib.io_handler import IOHandler
from lib.logger import  LoggerHandler
from lib.excel_handler import ExcelHandler
# from lib.csv_methods import CSVWorksheetHandler
from lib.excel_methods import WorksheetHandler
from app import *

###############################################################################################
####################################### Variables  ############################################
###############################################################################################
### i/o file class
io_handler = IOHandler()  # ✅ Pass path explicitly
### excel sheet class
excel_handler = ExcelHandler(io_handler) 





###############################################################################################
########################################### Code  #############################################
###############################################################################################
def main():
    
    App_main(io_handler, excel_handler) 
    
  
    
    
    
    
if __name__ == "__main__":
    main()
