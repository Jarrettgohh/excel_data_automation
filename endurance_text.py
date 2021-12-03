
from Excel.excel_functions import append_df_to_excel
import pandas
import xlwt
import re

txt_file = './Endurance/endurance_test_data.txt'
excel_file = './Endurance/endurance_test_data.xls'
excel_transfer_file = './Endurance/data_transfer.xls'

book = xlwt.Workbook()
ws = book.add_sheet('data')  # Add a sheet

file = open(txt_file, 'r+')

data = file.readlines()  # read all lines at once


for row_index in range(len(data)):
    # This will return a line of string data
    row = data[row_index].split()

    for data_index in range(len(row)):
        row_data = re.sub('Â', '', row[data_index])
        ws.write(row_index, data_index, row_data)  # Write to cell i, j


book.save(excel_file)
file.close()


# Read with pandas
df = pandas.read_excel(excel_file, sheet_name='data',
                       usecols='C:D')

voltage_polarization_data = df.iloc[40:542]

append_df_to_excel(df=voltage_polarization_data, filename=excel_transfer_file,
                   sheet_name='data_transfer', startrow=2, startcol=2,)
