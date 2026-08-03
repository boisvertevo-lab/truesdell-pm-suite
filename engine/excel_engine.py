import os
from pypdf import PdfReader, PdfWriter
import win32com.client


class PDFEngine:

    def __init__(self):
        self.excel = None

    ############################################################

    def start_excel(self):

        if self.excel is None:
            self.excel = win32com.client.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False

    ############################################################

    def stop_excel(self):

        if self.excel:
            self.excel.Quit()
            self.excel = None

    ############################################################

    def export_cover_sheet(
        self,
        workbook_path,
        worksheet_name,
        output_folder,
    ):

        self.start_excel()

        workbook = self.excel.Workbooks.Open(
            os.path.abspath(workbook_path)
        )

        try:

            sheet = workbook.Worksheets(worksheet_name)

            output_pdf = os.path.join(
                output_folder,
                f"{worksheet_name} Cover Sheet.pdf"
            )

            sheet.ExportAsFixedFormat(
                Type=0,
                Filename=os.path.abspath(output_pdf),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )

            workbook.Close(False)

            return output_pdf

        finally:

            self.stop_excel()

    ############################################################

    def find_attachment(
        self,
        pdf_library,
        submittal_number,
    ):

        search = f"Submittal #{int(submittal_number)}"

        for root, _, files in os.walk(pdf_library):

            for file in files:

                if not file.lower().endswith(".pdf"):
                    continue

                if file.startswith(search):

                    return os.path.join(root, file)

        return None

    ############################################################

    def merge_pdfs(
        self,
        cover_pdf,
        attachment_pdf,
        output_pdf,
    ):

        writer = PdfWriter()

        cover = PdfReader(cover_pdf)

        for page in cover.pages:
            writer.add_page(page)

        attach = PdfReader(attachment_pdf)

        for page in attach.pages:
            writer.add_page(page)

        with open(output_pdf, "wb") as f:
            writer.write(f)

        return output_pdf

    ############################################################

    def build_submittal(
        self,
        workbook_path,
        worksheet_name,
        pdf_library,
        output_folder,
    ):

        cover = self.export_cover_sheet(
            workbook_path,
            worksheet_name,
            output_folder,
        )

        attachment = self.find_attachment(
            pdf_library,
            worksheet_name,
        )

        if attachment is None:

            raise Exception(
                f"Could not locate PDF for Submittal #{int(worksheet_name)}"
            )

        final_pdf = os.path.join(
            output_folder,
            os.path.basename(attachment),
        )

        self.merge_pdfs(
            cover,
            attachment,
            final_pdf,
        )

        return final_pdf