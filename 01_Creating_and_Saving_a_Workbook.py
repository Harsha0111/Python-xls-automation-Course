from openpyxl import Workbook

# Create a new workbook and select the active sheet
wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

# Save the workbook
wb.save("output/01_Creating_and_Saving_a_Workbook.xlsx")
print("Workbook created and saved as '01_Creating_and_Saving_a_Workbook.xlsx'")
