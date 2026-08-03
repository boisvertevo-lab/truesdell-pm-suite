from openpyxl import load_workbook


class ExcelEngine:

    def __init__(self, filename):

        self.filename = filename

        self.workbook = load_workbook(
            filename,
            data_only=True
        )

        self.log = self.workbook["Log"]

    @property
    def project_name(self):

        return self.log["B3"].value or ""

    @property
    def description(self):

        return self.log["B4"].value or ""

    @property
    def job_number(self):

        return self.log["B5"].value or ""

    def get_submittals(self):

        rows = []

        row = 15
        blank = 0

        while True:

            number = self.log[f"A{row}"].value
            description = self.log[f"B{row}"].value

            if number is None and description is None:

                blank += 1

                if blank >= 5:
                    break

                row += 1
                continue

            blank = 0

            rows.append({

                "number": number,

                "description": description or "",

                "supplier": self.log[f"C{row}"].value or "",

                "spec": self.log[f"D{row}"].value or "",

                "submitted": self.log[f"E{row}"].value,

                "returned": self.log[f"F{row}"].value,

                "approved": self.log[f"G{row}"].value or ""

            })

            row += 1

        return rows