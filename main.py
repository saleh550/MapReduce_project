import pandas as pd
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
# from openpyxl import calculation
wb = load_workbook('./DataFile.xlsx', data_only=True)
ws = wb.active
Multi_year_average = 0
Multi_year_sum = 0
row_sum = 0
for row in range(1, 122):
    for col in range(1, 14):
        char = get_column_letter(col)
        if char != 'A':
            row_sum += ws[char+str(row)].value
            Multi_year_sum += ws[char+str(row)].value
        else:
            row_sum = 0

Multi_year_average = Multi_year_sum/121

print("The multi year average is: " + str(Multi_year_average))
# column_letter = 'A'
# row_number = 2
# cell = f"{column_letter}{row_number}"
# calculated_value = ws[cell].value
#
# print(calculated_value)
