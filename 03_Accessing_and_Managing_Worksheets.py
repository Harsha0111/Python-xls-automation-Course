from openpyxl import Workbook


wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

# Create a new sheet
new_sheet = wb.create_sheet("NewSheet")

# Rename a sheet
new_sheet.title = "RenamedSheet"

# Copy a sheet
copied_sheet = wb.copy_worksheet(new_sheet)
copied_sheet.title = "CopiedSheet"

# Delete a sheet
wb.remove(new_sheet)

# Save changes
wb.save("output/03_Accessing_and_Managing_Worksheets.xlsx")
print("Sheet operations completed and saved as '03_Accessing_and_Managing_Worksheets.xlsx'")
