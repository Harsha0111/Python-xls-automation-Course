from openpyxl import load_workbook

# Load an existing workbook
wb = load_workbook("output/01_Creating_and_Saving_a_Workbook.xlsx")
ws = wb.active
print(f"Loaded sheet title: {ws.title}")
