import os
import openpyxl

FOLDER_CREATION_LOCATION = r"C:\Users\Andrius\Python\_Projects\YT\Creating_folders_from_excel\result"
EXCEL_FILE_NAME = r"C:\Users\Andrius\Python\_Projects\YT\Creating_folders_from_excel\data\test1.xlsx"

workbook = openpyxl.load_workbook(EXCEL_FILE_NAME)
sheet = workbook['Sheet1']

column_values = [cell.value for col in sheet.iter_cols(
    min_row=2, max_row=None, min_col=2, max_col=2) for cell in col]

print(column_values)

# for value in column_values:
#     print("Creating folder: ", value)

# for value in column_values:
#     folderName = value
#     baseDir = FOLDER_CREATION_LOCATION
#     os.makedirs(os.path.join(baseDir, folderName))
#     print("Created folder: ", folderName)