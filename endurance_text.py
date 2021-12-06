import openpyxl
from Excel.excel_functions import append_df_to_excel
import pandas
import re
import json
import subprocess

config_json = open('config.json', 'r')
config_json = json.load(config_json)

endurance_test_config = config_json['endurance_test']

txt_file = './Endurance/endurance_test_data_PV.txt'
excel_file = './Endurance/endurance_test_data.xlsx'

sheet_name = 'data_transfer'

print('-----------------------------')
print('\nRemeber to edit the config.json file\n')
print('-----------------------------')
print('-----------------------------')
print('\nRun options\n')
print('-----------------------------')
print('\n')
print(
    '1. Transfer data from text (.txt) file to excel (.xlsx) file. Select this option if you wish to update the config.json file according to a new excel file cell format.'
)
print(
    '2. Read excel (.xlsx) files as configured in config.json and re-format to new excel files.'
)
print('\n')

print('-----------------------------')
user_selection = input('Enter your choice: ')


def execute_powershell(command: str):
    subprocess.Popen(['powershell.exe', command])


def transfer_txt_to_xlsx():

    for file in endurance_test_config['txt_files_to_transfer_to_excel']:

        wb = openpyxl.Workbook()
        ws = wb.worksheets[0]

        file_name = file['file_path']
        file = open(file_name, 'r+')

        data = file.readlines()  # read all lines at once

        for row_index in range(len(data)):
            # This will return a line of string data
            row = data[row_index].split()

            for col_index in range(len(row)):
                row_data = re.sub('Â', '', row[col_index])
                ws.cell(row=row_index + 1,
                        column=col_index + 1).value = row_data

        excel_file_path = f"{file_name.replace('.txt', '')}.xlsx"

        wb.save(excel_file_path)
        file.close()

        print(f'Transferred data to {excel_file_path}')
        print('Opening the file...')

        # Open the new Excel file after data is written to it
        execute_powershell(f'Invoke-Item \"{excel_file_path}\"')


def reformat_xlsx():
    return

    df = pandas.read_excel(excel_file, sheet_name=sheet_name, usecols='C:D')

    voltage_polarization_data = df.iloc[40:542]

    append_df_to_excel(
        df=voltage_polarization_data,
        filename=excel_file,
        sheet_name=sheet_name,
        startrow=2,
        startcol=2,
    )


if user_selection == "1":
    transfer_txt_to_xlsx()

elif user_selection == 2:
    reformat_xlsx()
