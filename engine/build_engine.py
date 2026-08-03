from engine.pdf_engine import PDFEngine


class BuildEngine:

    def __init__(self):
        self.pdf = PDFEngine()

    def build_selected(
        self,
        workbook_path,
        worksheet_name,
        pdf_library,
        output_folder,
    ):

        return self.pdf.build_submittal(
            workbook_path,
            worksheet_name,
            pdf_library,
            output_folder,
        )