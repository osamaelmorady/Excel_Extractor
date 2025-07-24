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
    
    
    
    in_defect_db_df = excel_handler.read_excel_sheet("database_matrix")
    in_running_cond_df = excel_handler.read_excel_sheet("running_condition_matrix")
    in_retry_cond_df = excel_handler.read_excel_sheet("retry_condition_matrix")
    in_frozen_cond_df = excel_handler.read_excel_sheet("frozen_condition_matrix")
    in_stop_cond_df = excel_handler.read_excel_sheet("stop_condition_matrix")
    in_reinit_cond_df = excel_handler.read_excel_sheet("reinitialize_condition_matrix")
    in_fim_action_df = excel_handler.read_excel_sheet("fim_action_matrix")
    in_retry_action_df = excel_handler.read_excel_sheet("retry_action_matrix")

    

    # in_defect_db_df = excel_handler.read_excel_sheet("Defect Database")
    # in_running_cond_df = excel_handler.read_excel_sheet("Defect running condition matrix")
    # in_retry_cond_df = excel_handler.read_excel_sheet("Defect retry condition matrix")
    # in_frozen_cond_df = excel_handler.read_excel_sheet("Defect Frozen condition matrix")
    # in_stop_cond_df = excel_handler.read_excel_sheet("Defect Stop condition matrix")
    # in_reinit_cond_df = excel_handler.read_excel_sheet("Defect Reinitialize condition m")
    # in_fim_action_df = excel_handler.read_excel_sheet("FIM action matrix")
    # in_retry_action_df = excel_handler.read_excel_sheet("Defect retry action matrix")



    in_defect_db = WorksheetHandler( io_handler, 'defect_db' , in_defect_db_df)    
    in_running_cond = WorksheetHandler( io_handler, 'running_cond' , in_running_cond_df)
    in_retry_action = WorksheetHandler( io_handler, 'retry_action' , in_retry_action_df)
    in_frozen_cond = WorksheetHandler( io_handler, 'frozen_cond' , in_frozen_cond_df)
    in_stop_cond = WorksheetHandler( io_handler, 'stop_cond' , in_stop_cond_df)
    in_reinit_cond = WorksheetHandler( io_handler, 'reinit_cond' , in_reinit_cond_df)
    in_fim_action = WorksheetHandler( io_handler, 'fim_action' , in_fim_action_df)
    in_retry_cond = WorksheetHandler( io_handler, 'retry_cond' , in_retry_cond_df)
    




    # in_stop_cond.remove_row(0)
    # in_stop_cond.remove_row(1)
    
    in_defect_db.save_sheet_as_csv()
    in_running_cond.save_sheet_as_csv() 
    in_retry_action.save_sheet_as_csv()    
    in_frozen_cond.save_sheet_as_csv() 
    in_stop_cond.save_sheet_as_csv() 
    in_reinit_cond.save_sheet_as_csv() 
    in_fim_action.save_sheet_as_csv() 
    in_retry_cond.save_sheet_as_csv() 

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ### external logger class  
    # logger_handler = LoggerHandler(config_path) 
    # logger = logger_handler.logger
    # logger.info("Logger initialized successfully.")


############ Test IO Handler ###############
    # print(io_handler.input_path)
    # print(io_handler.get_input_folder_path())
    # print(io_handler.get_intermediate_folder_path())
    # print(io_handler.get_output_folder_path())
    # print(io_handler.get_summary_file_path())
    # print(io_handler.get_log_file_path())
    # print(io_handler.get_input_sheets_names())
    
    
    
    
############ Test Excel Extractor , csv generator ###############
    # sheet_dict = excel_handler.read_all_input_sheets()
 ############### generate output sheets ###############
    # excel_handler.write_csv_sheet(sheet_dict[input_sheets_map['in_running_cond']],input_sheets_map['in_running_cond'])
    # excel_handler.write_csv_sheet(sheet_dict[input_sheets_map['in_retry_cond']],input_sheets_map['in_retry_cond'])  
    # excel_handler.write_csv_sheet(sheet_dict[input_sheets_map['in_frozen_cond']],input_sheets_map['in_frozen_cond'])  
    # excel_handler.write_csv_sheet(sheet_dict[input_sheets_map['in_stop_cond']],input_sheets_map['in_stop_cond'])  
    # excel_handler.write_csv_sheet(sheet_dict[input_sheets_map['in_reinit_cond']],input_sheets_map['in_reinit_cond'])  
    
    
    

    

    
    
    
    
    
if __name__ == "__main__":
    main()
