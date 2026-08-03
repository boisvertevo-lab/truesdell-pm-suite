def buildSelected(self):

    if self.table.currentRow() < 0:

        QMessageBox.information(
            self,
            "Select Submittal",
            "Please select a row."
        )
        return

    if self.output_folder == "":

        QMessageBox.information(
            self,
            "Output Folder",
            "Please choose an output folder."
        )
        return

    if self.pdf_library == "":

        QMessageBox.information(
            self,
            "PDF Library",
            "Please choose the PDF Library."
        )
        return

    row = self.table.currentRow()

    sheet = str(
        self.table.item(row, 0).text()
    ).zfill(2)

    try:

        pdf = self.pdf.build_submittal(
            self.workbook_path,
            sheet,
            self.pdf_library,
            self.output_folder
        )

        QMessageBox.information(
            self,
            "Complete",
            f"Finished!\n\n{pdf}"
        )

    except Exception as ex:

        QMessageBox.critical(
            self,
            "Build Error",
            str(ex)
        )