from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from engine.build_engine import BuildEngine
from engine.excel_engine import ExcelEngine


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("ProjectFlow")
        self.resize(1200, 750)

        self.workbook_path = ""
        self.output_folder = ""
        self.pdf_library = ""

        self.excel = None
        self.builder = BuildEngine()

        self.project_value = QLabel("")
        self.job_value = QLabel("")
        self.description_value = QLabel("")

        self.open_workbook_btn = QPushButton("Open Workbook")
        self.pdf_library_btn = QPushButton("PDF Library")
        self.output_folder_btn = QPushButton("Output Folder")
        self.build_selected_btn = QPushButton("Build Selected")
        self.build_all_btn = QPushButton("Build All")

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(
            ["#", "Description", "Supplier", "Spec", "Approved"]
        )
        self.table.setSelectionBehavior(self.table.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(self.table.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.open_workbook_btn.clicked.connect(self.open_workbook)
        self.pdf_library_btn.clicked.connect(self.choose_pdf_library)
        self.output_folder_btn.clicked.connect(self.choose_output_folder)
        self.build_selected_btn.clicked.connect(self.build_selected)
        self.build_all_btn.clicked.connect(self.build_all)

        info_layout = QFormLayout()
        info_layout.addRow("Project", self.project_value)
        info_layout.addRow("Job", self.job_value)
        info_layout.addRow("Description", self.description_value)

        project_info_group = QGroupBox("Project Information")
        project_info_group.setLayout(info_layout)

        button_row = QHBoxLayout()
        button_row.addWidget(self.open_workbook_btn)
        button_row.addWidget(self.pdf_library_btn)
        button_row.addWidget(self.output_folder_btn)
        button_row.addStretch()
        button_row.addWidget(self.build_selected_btn)
        button_row.addWidget(self.build_all_btn)

        layout = QVBoxLayout(self)
        layout.addWidget(project_info_group)
        layout.addLayout(button_row)
        layout.addWidget(self.table)

    def open_workbook(self):
        workbook_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Workbook",
            "",
            "Excel Workbooks (*.xlsx *.xlsm *.xls)",
        )

        if not workbook_path:
            return

        try:
            self.excel = ExcelEngine(workbook_path)
            self.workbook_path = workbook_path

            self.project_value.setText(self.excel.project_name)
            self.job_value.setText(self.excel.job_number)
            self.description_value.setText(self.excel.description)

            self.populate_submittals()

        except Exception as exc:
            QMessageBox.critical(self, "Open Workbook Error", str(exc))

    def populate_submittals(self):
        if self.excel is None:
            return

        submittals = self.excel.get_submittals()
        self.table.setRowCount(len(submittals))

        for row_index, submittal in enumerate(submittals):
            self.table.setItem(row_index, 0, QTableWidgetItem(str(submittal.get("number", ""))))
            self.table.setItem(row_index, 1, QTableWidgetItem(str(submittal.get("description", ""))))
            self.table.setItem(row_index, 2, QTableWidgetItem(str(submittal.get("supplier", ""))))
            self.table.setItem(row_index, 3, QTableWidgetItem(str(submittal.get("spec", ""))))
            self.table.setItem(row_index, 4, QTableWidgetItem(str(submittal.get("approved", ""))))

    def choose_pdf_library(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select PDF Library Folder",
        )

        if folder:
            self.pdf_library = folder

    def choose_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder",
        )

        if folder:
            self.output_folder = folder

    def build_selected(self):
        if self.excel is None or not self.workbook_path:
            QMessageBox.information(
                self,
                "Workbook Required",
                "Please open a workbook first.",
            )
            return

        if self.table.currentRow() < 0:
            QMessageBox.information(
                self,
                "Select Submittal",
                "Please select a row to build.",
            )
            return

        if not self.output_folder:
            QMessageBox.information(
                self,
                "Output Folder",
                "Please choose an output folder.",
            )
            return

        row = self.table.currentRow()
        worksheet_name = self.table.item(row, 0).text().strip().zfill(2)

        try:
            result = self.builder.build_selected(
                self.workbook_path,
                worksheet_name,
                self.pdf_library,
                self.output_folder,
            )
            QMessageBox.information(
                self,
                "Complete",
                f"Build finished successfully.\n\n{result}",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Build Error",
                str(exc),
            )

    def build_all(self):
        if self.excel is None or not self.workbook_path:
            QMessageBox.information(
                self,
                "Workbook Required",
                "Please open a workbook first.",
            )
            return

        if not self.output_folder:
            QMessageBox.information(
                self,
                "Output Folder",
                "Please choose an output folder.",
            )
            return

        if self.table.rowCount() == 0:
            QMessageBox.information(
                self,
                "No Submittals",
                "There are no submittals to build.",
            )
            return

        errors = []

        for row in range(self.table.rowCount()):
            worksheet_name = self.table.item(row, 0).text().strip().zfill(2)

            try:
                self.builder.build_selected(
                    self.workbook_path,
                    worksheet_name,
                    self.pdf_library,
                    self.output_folder,
                )
            except Exception as exc:
                errors.append(f"{worksheet_name}: {exc}")

        if errors:
            QMessageBox.critical(
                self,
                "Build Error",
                "\n".join(errors),
            )
        else:
            QMessageBox.information(
                self,
                "Complete",
                "All submittals were built successfully.",
            )
