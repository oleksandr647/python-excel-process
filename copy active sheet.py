import win32com.client as win32

# Connect to the running instance of Excel
excel = win32.Dispatch("Excel.Application")

# Specify the target file name (e.g., "MyWorkbook.xlsx")
target_filename = "CO3-Spreadsheet.xls"

# Find the workbook by name among all open workbooks
workbook = None
for wb in excel.Workbooks:
    if wb.Name == target_filename:
        workbook = wb
        break

if workbook:
    # Get the active sheet in the specified workbook
    active_sheet = workbook.ActiveSheet

    # Copy the active sheet within the same workbook
    # To avoid creating a new workbook, explicitly reference the workbook
    active_sheet.Copy(Before=active_sheet)
    
    # Optionally rename the duplicated sheet to avoid naming conflicts
    duplicated_sheet = workbook.Sheets(active_sheet.Index - 1)
    duplicated_sheet.Name = f"{active_sheet.Name}_Copy"

    print(f"Duplicated sheet '{active_sheet.Name}' within the same workbook as '{duplicated_sheet.Name}'.")
else:
    print(f"Workbook '{target_filename}' not found among open Excel files!")

# Optional: Make Excel visible if it's not already
excel.Visible = True
