import openpyxl
import os
import subprocess


def transfer_single_csv_to_xlsx(path_to_csv: str, folder_dir_to_write: str,
                                file_name_to_write: str):

    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]

    file = open(path_to_csv, 'r+')

    data = file.readlines()  # read all lines at once

    for row_index in range(len(data)):
        # This will return a line of string data
        row = data[row_index].split()

        for col_index in range(len(row)):
            row_data = row[col_index]

            if 'E' in row_data:
                split = row_data.split('E')
                value = split[0]
                exp = split[1]

                # Format the float exponent value (With E)
                try:
                    value = '{:,.2f}'.format(float(value))
                    row_data = f'{value}E{exp}'

                except:
                    pass

            row_data = row_data.replace(',', '')
            ws.cell(row=row_index + 1, column=col_index + 1).value = row_data

    try:
        os.makedirs(folder_dir_to_write)

    except FileExistsError:
        # directory already exists
        pass

    excel_file_path = f'{folder_dir_to_write}\\{file_name_to_write}'

    wb.save(excel_file_path)
    file.close()


def execute_powershell(command: str):
    subprocess.Popen(['powershell.exe', command])
