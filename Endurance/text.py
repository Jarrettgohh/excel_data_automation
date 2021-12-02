# import re

# with open('test.txt', 'rt') as file:

#     for line in file:
#         print(line)


# from Excel.excel_functions import append_df_to_excel
from Excel.excel_functions import append_df_to_excel
import pandas
import xlwt
import re
import sys
import os


cwd = os.getcwd()
# sys.path.insert(0, cwd.replace('Endurance', 'excel_functions'))
# sys.path.insert(
#     0, "C:/Users/gohja/Desktop/excel_data_automation")
# sys.path.append("C:/Users/gohja/Desktop/excel_data_automation/Excel")


# print(cwd.replace('Endurance', 'excel_functions'))

# file_to_read = 'endurance_test_data.txt'
# file_to_write = 'endurance_test_data.xls'
# transfer_file = 'data_transfer.xls'

# book = xlwt.Workbook()
# ws = book.add_sheet('data')  # Add a sheet

# f = open(file_to_read, 'r+')

# data = f.readlines()  # read all lines at once


# for i in range(len(data)):
#     # This will return a line of string data, you may need to convert to other formats depending on your use case
#     row = data[i].split()

#     for j in range(len(row)):
#         row_data = re.sub('Â', '', row[j])
#         ws.write(i, j, row_data)  # Write to cell i, j

# book.save(file_to_write)
# f.close()


# Read with pandas
# df = pandas.read_excel(file_to_write, sheet_name='data',
#                        usecols='C:D')

# voltage_polarization_data = df.iloc[40:542]

# append_df_to_excel(df=voltage_polarization_data, filename='transfer_file',
#                    sheet_name='data_transfer', startrow='2', startcol='B')
