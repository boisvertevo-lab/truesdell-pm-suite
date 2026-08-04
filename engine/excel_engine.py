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

        while True:

            # Read values from the Log sheet
            number = self.log[f"A{row}"].value
            description = self.log[f"B{row}"].value

            # Convert None to strings
            number = "" if number is None else str(number).strip()
            description = "" if description is None else str(description).strip()

            # A blank description means there are no more submittals
            if description == "":
                break

            # Skip rows that don't contain a numeric submittal number
            if not number.isdigit():
                row += 1
                continue

            rows.append({
                "number": number,
                "description": description,
                "supplier": self.log[f"C{row}"].value or "",
                "spec": self.log[f"D{row}"].value or "",
                "submitted": self.log[f"E{row}"].value,
                "returned": self.log[f"F{row}"].value,
                "approved": self.log[f"G{row}"].value or "",
            })

            row += 1

        return rows