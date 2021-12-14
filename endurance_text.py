from numpy import number
import pandas as pd
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
            df=voltage_polarization_data,
            filename=file_to_write,
            sheet_name=sheet_name,
            startrow=2,
            startcol=col_to_write,
        )


def transfer_single_txt_to_xlsx(file_path: str):
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

    excel_file_path = f"{file_path.replace('.txt', '')}.xlsx"

    wb.save(excel_file_path)
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


def reformat_xlsx():
    excel_files_to_format = endurance_test_config['excel_files_to_read']

    for file in excel_files_to_format:
        format_to_xlsx(
            file_path=file['file_path'],
            file_config=file,
        )


def transfer_and_reformat():
    transfer_txt_to_xlsx()
    print('\n')
    reformat_xlsx()


def format_txt_files():
    files_to_format = config_json['files_to_format']

    files_to_format_names = list(files_to_format.keys())

    # Transfer the .txt file to .xlsx file (Text to excel)
    for file_to_format_name in files_to_format_names:
        files = files_to_format[file_to_format_name]

        # Hard transfer each file
        for file in files:
            file_path = file['file_path_to_read']
            transfer_single_txt_to_xlsx(file_path)

        # Transfer and extract each file
        for file_index, file in enumerate(files):
            file_path = file['file_path_to_read']
            file_path_xlsx_data = f'{file_path.replace(".txt", ".xlsx")}'

            file_path_xlsx_results = file_to_format_name
            initial_col = (file_index * 2) + 1

            format_to_xlsx(file_path=file_path_xlsx_data,
                           file_config=file,
                           file_to_write=file_path_xlsx_results,
                           initial_col=initial_col)


if user_selection == "1":
    transfer_txt_to_xlsx()

elif user_selection == "2":
    # reformat_xlsx()
    format_txt_files()

elif user_selection == "3":
    transfer_and_reformat()