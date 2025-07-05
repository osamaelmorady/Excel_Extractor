# src/parser.py
class ExcelParser:
    def __init__(self, io_handler, logger):
        self.io_handler = io_handler
        self.logger = logger

    def parse_sheets(self):
        try:
            excel = self.io_handler.read_excel()
        except FileNotFoundError as e:
            self.logger.error(str(e))
            return

        for sheet_name in self.io_handler.config['sheets_to_process']:
            if sheet_name in excel.sheet_names:
                df = excel.parse(sheet_name)
                if df.empty:
                    self.logger.warning(f"Sheet {sheet_name} is empty.")
                    continue
                self.io_handler.write_csv(df, sheet_name)
                self.logger.info(f"Processed sheet: {sheet_name}")
            else:
                self.logger.warning(f"Sheet {sheet_name} not found in Excel file.")