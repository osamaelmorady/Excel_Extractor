
from lib.excel_methods import WorksheetHandler



def App_main(io_handler, excel_handler):
    
    in_defect_db_df = excel_handler.read_excel_sheet("database_matrix")
    in_running_cond_df = excel_handler.read_excel_sheet("running_condition_matrix")
    in_retry_cond_df = excel_handler.read_excel_sheet("retry_condition_matrix")
    in_frozen_cond_df = excel_handler.read_excel_sheet("frozen_condition_matrix")
    in_stop_cond_df = excel_handler.read_excel_sheet("stop_condition_matrix")
    in_reinit_cond_df = excel_handler.read_excel_sheet("reinitialize_condition_matrix")
    in_fim_action_df = excel_handler.read_excel_sheet("fim_action_matrix")
    in_retry_action_df = excel_handler.read_excel_sheet("retry_action_matrix")

 
    in_defect_db = WorksheetHandler( io_handler, 'defect_db' , in_defect_db_df)    
    in_running_cond = WorksheetHandler( io_handler, 'running_cond' , in_running_cond_df)
    in_retry_action = WorksheetHandler( io_handler, 'retry_action' , in_retry_action_df)
    in_frozen_cond = WorksheetHandler( io_handler, 'frozen_cond' , in_frozen_cond_df)
    in_stop_cond = WorksheetHandler( io_handler, 'stop_cond' , in_stop_cond_df)
    in_reinit_cond = WorksheetHandler( io_handler, 'reinit_cond' , in_reinit_cond_df)
    in_fim_action = WorksheetHandler( io_handler, 'fim_action' , in_fim_action_df)
    in_retry_cond = WorksheetHandler( io_handler, 'retry_cond' , in_retry_cond_df)
    

    Handle_Defect_db_matrix(in_defect_db)
    Handle_Running_cond_matrix(in_running_cond)
    Handle_Retry_cond_matrix(in_retry_cond)
    Handle_Frozen_cond_matrix(in_frozen_cond)
    Handle_Stop_cond_matrix(in_stop_cond)
    Handle_Reinit_cond_matrix(in_reinit_cond)
    Handle_Fim_action_matrix(in_fim_action)
    Handle_Retry_action_matrix(in_retry_action)

    
    
    
    
    
    
def Handle_Defect_db_matrix(sheet): 
    sheet.save_sheet_as_csv()

    pass

def Handle_Running_cond_matrix(sheet): 
    sheet.add_column("E",50)
    sheet.add_row("5",5)
    sheet.remove_row("1")
    sheet.remove_row("4")     
    sheet.set_cell_value("B5","ON")
    
    sheet.save_sheet_as_csv()

def Handle_Retry_cond_matrix(sheet): 
    sheet.add_column("E",50)
    sheet.add_row("5",5)
    sheet.remove_row("1")
    sheet.remove_row("4")     
    sheet.set_cell_value("B5","ON")
    
    sheet.save_sheet_as_csv()

def Handle_Frozen_cond_matrix(sheet): 
    sheet.add_column("E",50)
    sheet.add_row("5",5)
    sheet.remove_row("1")
    sheet.remove_row("4")     
    sheet.set_cell_value("B5","ON")
    
    sheet.save_sheet_as_csv()



def Handle_Stop_cond_matrix(sheet): 
    sheet.add_column("E",50)
    sheet.add_row("5",5)
    sheet.remove_row("1")
    sheet.remove_row("4")     
    sheet.set_cell_value("B5","ON")
    
    sheet.save_sheet_as_csv()


def Handle_Reinit_cond_matrix(sheet): 
    sheet.add_column("E",50)
    sheet.add_row("5",5)
    sheet.remove_row("1")
    sheet.remove_row("4")     
    sheet.set_cell_value("B5","ON")
    
    sheet.save_sheet_as_csv()

def Handle_Fim_action_matrix(sheet): 
    sheet.add_column("E",50)
    sheet.add_row("5",5)
    sheet.remove_row("1")
    sheet.remove_row("4")     
    sheet.set_cell_value("B5","ON")
    
    sheet.save_sheet_as_csv()

def Handle_Retry_action_matrix(sheet): 
    sheet.add_column("E",50)
    sheet.add_row("5",5)
    sheet.remove_row("1")
    sheet.remove_row("4")     
    sheet.set_cell_value("B5","ON")
    
    sheet.save_sheet_as_csv()


    
    
