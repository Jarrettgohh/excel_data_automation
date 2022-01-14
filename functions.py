import openpyxl
import os
import subprocess
import sys
import shutil

import pandas as pd

from Excel.excel_functions import append_df_to_excel


def transfer_files_to_new_folder(current_file_dir: str, target_dir: str,
                                 target_file_name: str):

    try:
        os.makedirs(target_dir)

    except FileExistsError:
        pass

    try:
        shutil.copyfile(current_file_dir, target_dir + "/" + target_file_name)

    except:
        pass


def transfer_single_csv_to_xlsx(path_to_csv: str, folder_dir_to_write: str,
                                file_name_to_write: str):

    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]

    try:
        file = open(path_to_csv, 'r+')

    except:
        print(f'Failed to open "{path_to_csv}", file path does not exist.\n')
        sys.exit()

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


def create_file_and_append_df_to_xlsx(xlsx_folder_dir: str,
                                      xlsx_file_name: str, df: pd.DataFrame,
                                      startrow: int, startcol: int):
    try:
        df = df.astype('float')

    except:
        pass

    xlsx_file_path_to_write = f'{xlsx_folder_dir}{xlsx_file_name}'

    try:
        os.makedirs(xlsx_file_path_to_write)

    except FileExistsError:
        pass

    try:
        # Append dataframe to main excel file
        append_df_to_excel(df=df,
                           filename=xlsx_file_path_to_write,
                           startrow=startrow,
                           startcol=startcol)

    except:
        print(
            f'Failed to write to excel file. Ensure that the target file path "{xlsx_file_path_to_write}" is not running/open.\n'
        )
        sys.exit()


def execute_powershell(command: str):
    try:
        subprocess.check_output(['powershell.exe', command])

    except subprocess.CalledProcessError:
        raise Exception()


def execute_powershell_function(file_dir: str, fn_name: str, fn_args: str):
    try:
        cmd = ["powershell.exe", f". \"{file_dir}\";", f"&{fn_name} {fn_args}"]
        subprocess.check_output(cmd)

    except subprocess.CalledProcessError:
        raise Exception()