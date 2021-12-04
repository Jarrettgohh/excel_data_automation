import openpyxl
from Excel.excel_functions import append_df_to_excel
import pandas
import xlwt
import re
import json

config_json = open('config.json', 'r')
config_json = json.load(config_json)

txt_file = './Endurance/endurance_test_data_PV.txt'
excel_file = './Endurance/endurance_test_data.xlsx'

sheet_name = 'data_transfer'

wb = openpyxl.Workbook()
ws = wb.worksheets[0]

# book = xlwt.Workbook()
# ws = book.add_sheet()  # Add a sheet

file = open(txt_file, 'r+')

data = file.readlines()  # read all lines at once

for row_index in range(len(data)):
    # This will return a line of string data
    row = data[row_index].split()

    for col_index in range(len(row)):
        row_data = re.sub('Â', '', row[col_index])
        ws.cell(row=row_index + 1, column=col_index + 1).value = row_data

        # ws.write(row_index, data_index, row_data)  # Write to cell i, j

wb.save(excel_file)
file.close()

# Read with pandas
df = pandas.read_excel(excel_file, sheet_name=sheet_name, usecols='C:D')

voltage_polarization_data = df.iloc[40:542]

append_df_to_excel(
    df=voltage_polarization_data,
    filename=excel_file,
    sheet_name=sheet_name,
    startrow=2,
    startcol=2,
)
