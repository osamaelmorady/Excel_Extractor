# src/parser.py

# class SheetParser:
#     def __init__(self, io_handler, excel_handler):
#         self.io_handler = io_handler
#         self.excel_handler = excel_handler
#         self.sheet_dict = self.excel_handler.read_all_input_sheets()
#         self.sheets_map = self._map_sheet_names()

#     def _map_sheet_names(self):
#         """
#         Creates a dictionary mapping local variable names to original sheet names.
#         Update the keys below to represent your preferred variable names.
#         """
#         return {
#             "defect_db": "Defect Database",
#             "running_cond": "Defect running condition matrix",
#             "retry_cond": "Defect retry condition matrix",
#             "frozen_cond": "Defect Frozen condition matrix",
#             "stop_cond": "Defect Stop condition matrix",
#             "fim_action": "FIM action matrix",
#             "retry_action": "Defect retry action matrix"
#         }

#     def get_sheet(self, alias):
#         """Returns the DataFrame of the sheet using its local alias."""
#         original_name = self.sheets_map.get(alias)
#         if original_name and original_name in self.sheet_dict:
#             return self.sheet_dict[original_name]
#         raise KeyError(f"Alias '{alias}' not found or sheet is missing.")

# # Usage:
# # parser = SheetParser(io_handler, excel_handler)
# # df = parser.get_sheet("defect_db")