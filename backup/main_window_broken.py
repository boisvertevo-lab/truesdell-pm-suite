from PySide6.QtWidgets import *
from engine.excel_engine import ExcelEngine
from engine.pdf_engine import PDFEngine
from engine.build_engine import BuildEngine

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ProjectFlow")
        self.resize(1200,750)

        self.workbook_path=""
        self.output_folder=""
        self.pdf_library=""

        self.excel=None
        self.pdf=PDFEngine()
        self.builder=BuildEngine()

        self.projectLabel=QLabel("Project:")
        self.jobLabel=QLabel("Job:")
        self.descriptionLabel=QLabel("Description:")

        self.openBtn=QPushButton("Open Workbook")
        self.pdfBtn=QPushButton("PDF Library")
        self.outputBtn=QPushButton("Output Folder")
        self.buildBtn=QPushButton("Build Selected")
        self.buildAllBtn=QPushButton("Build All")

        self.table=QTableWidget(0,5)
        self.table.setHorizontalHeaderLabels(
            ["#","Description","Supplier","Spec","Approved"]
        )
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.openBtn.clicked.connect(self.openWorkbook)
        self.pdfBtn.clicked.connect(self.choosePdfLibrary)
        self.outputBtn.clicked.connect(self.chooseOutputFolder)
        self.buildBtn.clicked.connect(self.buildSelected)

        grid=QGridLayout()
        grid.addWidget(self.projectLabel,0,0)
        grid.addWidget(self.jobLabel,1,0)
        grid.addWidget(self.descriptionLabel,2,0)

        info=QGroupBox("Project")
        info.setLayout(grid)

        top=QHBoxLayout()
        top.addWidget(self.openBtn)
        top.addWidget(self.pdfBtn)
        top.addWidget(self.outputBtn)
        top.addStretch()
        top.addWidget(self.buildBtn)
        top.addWidget(self.buildAllBtn)

        lay=QVBoxLayout(self)
        lay.addWidget(info)
        lay.addLayout(top)
        lay.addWidget(self.table)

    # Remaining methods:
    # openWorkbook()
    # choosePdfLibrary()
    # chooseOutputFolder()
    # buildSelected()

