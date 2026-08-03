import os
import win32com.client


class PDFEngine:

    def __init__(self):
        self.excel = None

    def start(self):

        if self.excel is None:

            self.excel = win32com.client.DispatchEx(
                "Excel.Application"
            )

            self.excel.Visible = False
            self.excel.DisplayAlerts = False

    def stop(self):

        if self.excel:

            self.excel.Quit()
            self.excel = None

    def export_cover_sheet(
        self,
        workbook_path,
        worksheet_name,
        output_folder,
    ):

        self.start()

        workbook = self.excel.Workbooks.Open(
            os.path.abspath(workbook_path)
        )

        try:

            worksheet = workbook.Worksheets(worksheet_name)

            output_pdf = os.path.join(
                output_folder,
                f"{worksheet_name} Cover Sheet.pdf"
            )

            worksheet.ExportAsFixedFormat(
                Type=0,
                Filename=os.path.abspath(output_pdf)
            )

            return output_pdf

        finally:

            workbook.Close(False)
            self.stop()