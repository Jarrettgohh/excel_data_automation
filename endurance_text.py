from numpy import number
import pandas as pd
import openpyxl
from Excel.excel_functions import append_df_to_excel
import pandas
import re
import json
import subprocess
import os
import sys

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
print(
    '3. Transfer data from text (.txt) file to excel (.xlsx) file and extract the relevant data (configured in config.json) and re-format to a new excel file. (Combination of option 1 and 2)'
)
print('\n')

print('-----------------------------')
user_selection = input('Enter your choice: ')
print('\n')


def execute_powershell(command: str):
    subprocess.Popen(['powershell.exe', command])


def format_to_xlsx(file_path: str,
                   file_config,
                   file_to_write,
                   initial_col: number = 1):
    # f'{file_path.replace(".xlsx", "")}_transfer.xlsx'

    number_of_cycles = file_config['number_of_cycles']
    number_of_points = file_config['number_of_points']
    row_margin_buffer = file_config['row_margin_buffer']
    rows_to_read = file_config['rows_to_read']
    header_text = file_config['header_text']

    print(f"Formating excel file from path: {file_path}")

    for cycle_number in range(int(number_of_cycles)):

        start_row = file_config['start_row'] + (
            cycle_number * (number_of_points + row_margin_buffer + 1)
        ) if row_margin_buffer != None else file_config['start_row'] + (
            cycle_number * number_of_points)

        # print(f"Formatting row: {start_row}")

        df = pandas.read_excel(
            file_path,
            sheet_name=sheet_name,
            usecols=rows_to_read,
        )
        voltage_polarization_data = df.iloc[start_row:start_row +
                                            number_of_points]

        col_to_write = initial_col + (cycle_number * 2)

        if header_text != None:
            header_df = pd.DataFrame(data=[header_text])

            append_df_to_excel(
                df=header_df,
                filename=file_to_write,
                sheet_name=sheet_name,
                startrow=1,
                startcol=col_to_write,
            )

        append_df_to_excel(
            df=voltage_polarization_data.astype('float'),
            filename=file_to_write,
            sheet_name=sheet_name,
            startrow=2,
            startcol=col_to_write,
        )


def transfer_single_txt_to_xlsx(file_path: str,
                                folder_directory_to_transfer: str):
    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]

    file = open(file_path, 'r+')

    data = file.readlines()  # read all lines at once

    for row_index in range(len(data)):
        # This will return a line of string data
        row = data[row_index].split()

        for col_index in range(len(row)):
            row_data = re.sub('Â', '', row[col_index])
            ws.cell(row=row_index + 1, column=col_index + 1).value = row_data

    excel_file_path = f"{file_path.replace('txt', 'xlsx')}"

    try:
        os.makedirs(folder_directory_to_transfer)

    except FileExistsError:
        # directory already exists
        pass

    xlsx_file_name_match = re.search(r'(/\w*.xlsx)$', excel_file_path)

    if xlsx_file_name_match == None:
        print(
            'Invalid "file_path_to_read" argument in the config.json. Include the "/" to indicate file directories and include the ".txt" file extension.'
        )
        sys.exit()

    xlsx_file_name_index = xlsx_file_name_match.start()
    xlsx_file_name = excel_file_path[xlsx_file_name_index:]

    path_to_transfer = folder_directory_to_transfer + xlsx_file_name
    wb.save(path_to_transfer)

    file.close()

    print(f'Transferred data to {excel_file_path}')
    # print('Opening the file...')

    # # Open the new Excel file after data is written to it
    # execute_powershell(f'Invoke-Item \"{excel_file_path}\"')


def transfer_txt_to_xlsx():

    text_files_to_transfer = endurance_test_config[
        'txt_files_to_transfer_to_excel']

    for file in text_files_to_transfer:
        file_name = file['file_path']
        transfer_single_txt_to_xlsx(file_name)


def format_txt_files():
    files_to_format = endurance_test_config['files_to_format']

    files_to_format_names = list(files_to_format.keys())

    # Transfer the .txt file to .xlsx file (Text to excel)
    for file_to_format_name in files_to_format_names:
        files = files_to_format[file_to_format_name]

        # Hard transfer each file; direct transfer line by line from .txt to .xlsx
        for file in files:
            file_path = file['file_path_to_read']
            folder_directory_to_transfer = file['folder_directory_to_transfer']
            transfer_single_txt_to_xlsx(
                file_path,
                folder_directory_to_transfer=folder_directory_to_transfer)

        # Transfer and extract each file
        for file_index, file in enumerate(files):
            file_path = file['file_path_to_read']
            folder_directory_to_transfer = file['folder_directory_to_transfer']
            folder_path_to_write = file['folder_path_to_write']

            initial_col = (file_index * 2) + 1
            file_path_xlsx_results = folder_path_to_write + '/' + file_to_format_name

            file_path_xlsx_data = f'{file_path.replace(".txt", ".xlsx")}'

            # Getting the directory to the .xlsx files where the .txt data is transferred to
            xlsx_file_name_match = re.search(r'(/\w*.xlsx)$',
                                             file_path_xlsx_data)

            if xlsx_file_name_match == None:
                print(
                    'Invalid "file_path_to_read" argument in the config.json. Include the "/" to indicate file directories and include the ".txt" file extension.'
                )
                sys.exit()

            xlsx_file_name = file_path_xlsx_data[xlsx_file_name_match.start():]
            path_to_read = folder_directory_to_transfer + xlsx_file_name

            try:
                os.makedirs(folder_path_to_write)

            except FileExistsError:
                # directory already exists
                pass

            format_to_xlsx(file_path=path_to_read,
                           file_config=file,
                           file_to_write=file_path_xlsx_results,
                           initial_col=initial_col)


if user_selection == "1":
    transfer_txt_to_xlsx()

elif user_selection == "2":
    format_txt_files()
