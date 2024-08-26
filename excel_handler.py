from openpyxl import Workbook, load_workbook


class ExcelHandler:
    def handler(self, tools_log_file, svod_excel_filem, output_path):
        wb = Workbook()
