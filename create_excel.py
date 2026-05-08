from openpyxl import Workbook
from openpyxl.styles import Font


def create_excel_file(filename: str = "excel project.xlsx") -> None:
    workbook = Workbook()
    sheet = workbook["Sheet"]
    sheet.title = "Sheet1"

    cell = sheet["B12"]
    cell.value = "we are good"
    cell.font = Font(name="Calibri", size=12, bold=True)

    workbook.save(filename)


if __name__ == "__main__":
    create_excel_file()
