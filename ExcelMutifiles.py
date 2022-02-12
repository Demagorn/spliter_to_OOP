import openpyxl


class ExcelCopy:

    def __init__(self, source_file, **kwargs):
        self.source_file = source_file
        self.source_worksheet_name = kwargs.get("worksheet", None)

    def __enter__(self):
        self.source_wb = openpyxl.load_workbook(self.source_file)
        if self.source_worksheet_name:
            self.source_ws = self.source_wb[self.source_worksheet_name]
        else:
            self.source_ws = self.source_wb.active

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def create_sheet_copy(self, greed_lines=False, **kwargs):
        """creates a copy and delete the sample sheet"""
        new_worksheet = self.source_wb.copy_worksheet(self.source_ws)
        new_worksheet.title = kwargs.get("title", None)
        new_worksheet.sheet_properties.tabColor = kwargs.get("title_color", None)
        new_worksheet.sheet_view.showGridLines = greed_lines

    def create_workbook(self, workbook_name, *args):
        """create file according to format, for multiple sheet provide sheet names as list to *args"""

