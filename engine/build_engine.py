from engine.pdf_engine import PDFEngine


class BuildEngine:

    def __init__(self):

        self.pdf = PDFEngine()

    def build_selected(
        self,
        workbook_path,
        worksheet_name,
        output_folder,
    ):

        return self.pdf.export_cover_sheet(
            workbook_path,
            worksheet_name,
            output_folder,
        )