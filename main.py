from openpyxl import Workbook, workbook

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "hello"
sheet["A2"] = "world!"

workbook.save(filename = "hello_world.xlsx")
