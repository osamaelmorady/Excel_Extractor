import sys
import os
import datetime


# class UTILs:

#     def validate_output_files(io_handler):
#         generated_files = [f.lower() for f in os.listdir(io_handler.config['intermediate_folder']) if f.endswith('.csv')]
#         missing = []
#         for sheet in io_handler.config["input_files_to_process"]['sheets']:
#             expected_file = f"{sheet}.csv".lower()
#             if expected_file not in generated_files:
#                 missing.append(sheet)
#         return missing



#     def write_genetation_summary(self, missing):
#         summary_path = self.get_summary_path()
#         with open(summary_path, "w", encoding="utf-8") as generation_summary_file:
#             generation_summary_file.write(f"Run Timestamp: {datetime.datetime.now()}\n\n")
#             if not missing:
#                 message = "✅ Generation is done successfully."
#                 print(message)
#                 generation_summary_file.write(message + "\n")
#             else:
#                 message = f"❌ Error: Missing sheet output(s): {', '.join(missing)}"
#                 print(message)
#                 generation_summary_file.write(message + "\n")
#                 raise RuntimeError(message)