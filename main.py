import pandas as pd
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


# from openpyxl import calculation

def getMonth():
    return 5


wb = load_workbook('./DataFile.xlsx', data_only=True)
ws = wb.active

max_investors_number_inYear = {
    "year": '',
    "investors": 0,
}
min_investors_number_inYear = {
    "year": '',
    "investors": 0,
}
max_investors_number_inMonth = {
    "month": '',
    "year": '',
    "investors": 0

}
min_investors_number_inMonth = {
    "month": '',
    "year": '',
    "investors": 0

}
Multi_year_average = 0
Multi_year_sum = 0
min_number_inYear = 350 * 12
max_number_inYear = 0
min_number_inMonth = 350
max_number_inMonth = 0
row_sum = 0
my_list = []
for row in range(1, 122):
    for col in range(1, 14):
        char = get_column_letter(col)
        if char != 'A':
            row_sum += ws[char + str(row)].value
            Multi_year_sum += ws[char + str(row)].value
            if ws[char + str(row)].value > max_number_inMonth:
                max_number_inMonth = ws[char + str(row)].value
                max_investors_number_inMonth['year'] = ws['A'+str(row)].value
                max_investors_number_inMonth['month'] = char
                max_investors_number_inMonth['investors'] = max_number_inMonth

        if char == 'M':
            # print(ws['A'+str(row)].value)
            # print(row_sum)
            # print("<------------------>")
            if row_sum > max_number_inYear:
                max_number_inYear = row_sum
                max_investors_number_inYear['year'] = ws['A'+str(row)].value
                max_investors_number_inYear['investors'] = max_number_inYear
            if row_sum < min_number_inYear:
                min_number_inYear = row_sum
                min_investors_number_inYear['year'] = ws['A'+str(row)].value
                min_investors_number_inYear['investors'] = min_number_inYear

            row_sum = 0


Multi_year_average = Multi_year_sum / 121


print("The multi year average is: " + str(Multi_year_average))
print("In " + str(max_investors_number_inYear['year']) + " was the max of investors of such " + str(max_investors_number_inYear['investors']))
print("In " + str(min_investors_number_inYear['year']) + " was the min of investors of such " + str(min_investors_number_inYear['investors']))
print(max_investors_number_inMonth['year'],max_investors_number_inMonth['month'],max_investors_number_inMonth['investors'])
# column_letter = 'A'
# row_number = 2
# cell = f"{column_letter}{row_number}"
# calculated_value = ws[cell].value
#
# print(calculated_value)
