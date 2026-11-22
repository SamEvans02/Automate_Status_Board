
import openpyxl
from pathlib import Path

# Get the directory where this script is running
print("Current working directory:", Path.cwd())

input_file = Path("Team_Status_Board 2026.xlsx")
output_file = Path("EDITED_Team_Status_Board 2026.xlsx")

# Check whether the source file exists
print("Input file exists:", input_file.exists())

# Load the existing workbook
wb = openpyxl.load_workbook(input_file)

# Inspect the worksheets
print("Worksheets:", wb.sheetnames)
print("Active worksheet:", wb.active.title)

sheet = wb.active

# Check existing data
print("B7 before editing:", sheet["B7"].value)
print("Worksheet dimensions:", sheet.max_row, "rows x", sheet.max_column, "columns")

# Make the change
sheet["B7"] = "Test"
print("B7 after editing:", sheet["B7"].value)

# Save to a separate file
wb.save(output_file)

print("Saved file:", output_file.resolve())
print("Output file size:", output_file.stat().st_size, "bytes")

# Reopen the saved file to verify it
check_wb = openpyxl.load_workbook(output_file)
check_sheet = check_wb[sheet.title]

print("B7 after reopening:", check_sheet["B7"].value)
print("Reopened worksheet dimensions:",
      check_sheet.max_row, "rows x", check_sheet.max_column, "columns")
