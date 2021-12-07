import openpyxl
from Excel.excel_functions import append_df_to_excel
import pandas
import re
import json
import subprocess

config_json = open('config.json', 'r')
config_json = json.load(config_json)

endurance_test_config = config_json['endurance_test']

sheet_name = 'Sheet'

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
print('\n')


def execute_powershell(command: str):
    subprocess.Popen(['powershell.exe', command])


def transfer_txt_to_xlsx():

    text_files_to_transfer = endurance_test_config[
        'txt_files_to_transfer_to_excel']

    for file in text_files_to_transfer:

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
        # print('Opening the file...')

        # # Open the new Excel file after data is written to it
        # execute_powershell(f'Invoke-Item \"{excel_file_path}\"')


def reformat_xlsx():
    excel_files_to_format = endurance_test_config['excel_files_to_read']

    for file in excel_files_to_format:

        number_of_cycles = file['number_of_cycles']
        number_of_points = file['number_of_points']
        row_margin_buffer = file['row_margin_buffer']
        rows_to_read = file['rows_to_read']

        file_path = file['file_path']

        print(f"\n\nFormating excel file from path: {file_path}")

        for cycle_number in range(int(number_of_cycles)):
            start_row = file['start_row'] + (
                cycle_number * (number_of_points + row_margin_buffer + 1))

            print(f"Formatting row: {start_row}")

            df = pandas.read_excel(file_path,
                                   sheet_name=sheet_name,
                                   usecols=rows_to_read)

            voltage_polarization_data = df.iloc[start_row:start_row +
                                                number_of_points]

            col_to_write = (cycle_number * 2) + 1

            append_df_to_excel(
                df=voltage_polarization_data,
                filename=f'{file_path.replace(".xlsx", "")}_transfer.xlsx',
                sheet_name=sheet_name,
                startrow=2,
                startcol=col_to_write,
            )


def transfer_and_reformat():
    transfer_txt_to_xlsx()
    print('\n')
    reformat_xlsx()


if user_selection == "1":
    transfer_txt_to_xlsx()

elif user_selection == "2":
    reformat_xlsx()

elif user_selection == "3":
    transfer_and_reformat()