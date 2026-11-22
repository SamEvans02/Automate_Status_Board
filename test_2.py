import openpyxl

wb = openpyxl.load_workbook("Team_Status_Board 2026.xlsx")

print("All worksheet names:", wb.sheetnames)

for ws in wb.worksheets:
    print(
        f"Name: {ws.title!r} | "
        f"State: {ws.sheet_state} | "
        f"Dimensions: {ws.max_row} rows x {ws.max_column} columns"
    )
    print("  A1:", repr(ws["A1"].value))
