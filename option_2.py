import json
import sys
import os
import re
import pandas as pd

from openpyxl.utils.cell import column_index_from_string
from functions import execute_powershell, transfer_single_csv_to_xlsx
from Excel.excel_functions import append_df_to_excel, xlsx_read_col_row

#
# Extract data from .txt, .csv or .xls file into a .xlsx file in a certain format
#

config_json = open('config.json', 'r')
config_json = json.load(config_json)

option_2_configs = config_json['option_2_configs']


def option_2():

    for config in option_2_configs:

        excel_file_paths_to_open = []

        # Root directory
        root_dir = config['ROOT_DIRECTORY']

        # To read
        to_read = config['TO_READ']
        file_type_to_read = to_read['file_type']
        relative_folder_directories = to_read['relative_folder_directories']
        files = to_read['files']
        cols_to_read = to_read['cols']
        rows_to_read = to_read['rows']

        # To transfer
        transfer_folder_dir = "\\transfer"

        # To write
        to_write = config['TO_WRITE']
        relative_folder_directory = to_write['relative_folder_directory']
        xlsx_file_name_to_write = to_write['file_name']
        to_write_row_settings = to_write['row_settings']
        to_write_col_settings = to_write['col_settings']

        xlsx_file_path_to_write = f'{root_dir}{relative_folder_directory}\\{xlsx_file_name_to_write}'

        # To convert the cols_to_read from letters to number values
        for index, col_value in enumerate(cols_to_read):
            if type(col_value) == int:
                continue

            col_index = column_index_from_string(col_value)
            cols_to_read[index] = col_index - 1

        for folder_dir_index, folder_dir in enumerate(
                relative_folder_directories):

            #
            # Writing of headers
            #

            to_write_start_row = to_write_row_settings['start_row']
            to_write_start_col_setting = to_write_col_settings['start_col']

            for file_index, file_name in enumerate(files):

                cols_to_read_len = len(cols_to_read)

                file_index_start_row = cols_to_read_len * file_index
                folder_index_start_row = ((folder_dir_index * len(files)) *
                                          cols_to_read_len)

                to_write_start_col_setting = column_index_from_string(
                    to_write_start_col_setting) - 1 if type(
                        to_write_start_col_setting
                    ) == str else to_write_start_col_setting
                to_write_cols = to_write_col_settings['cols']

                start_col_to_write = (
                    to_write_start_col_setting + folder_index_start_row +
                    file_index_start_row
                ) if to_write_cols == 'auto' else to_write_cols[file_index]

                header_df = pd.DataFrame([folder_dir.replace("\\", "")])

                if file_index == 0:
                    # Append the headers
                    append_df_to_excel(df=header_df,
                                       filename=xlsx_file_path_to_write,
                                       startrow=to_write_start_row,
                                       startcol=start_col_to_write)

                file_path_to_read = f'{root_dir}{folder_dir}\\{file_name}'

                if file_type_to_read == 'csv':
                    if '.csv' not in file_path_to_read:
                        print(
                            'Invalid "files" list argument in the config.json. Ensure that the file extensions follows the "file_type".'
                        )
                        sys.exit()

                    folder_dir_to_transfer = f'{root_dir}{folder_dir}{transfer_folder_dir}'

                    file_name_to_transfer = file_name.replace(".csv", ".xlsx")
                    file_dir_to_transfer = f'{folder_dir_to_transfer}\\{file_name_to_transfer}'

                    if folder_dir_index == 0:
                        print(
                            '\nConverting .csv files into .xlsx format, and transferring into new folder...'
                        )

                    transfer_single_csv_to_xlsx(
                        path_to_csv=file_path_to_read,
                        folder_dir_to_write=folder_dir_to_transfer,
                        file_name_to_write=file_name_to_transfer)

                    df = xlsx_read_col_row(xlsx_file=file_dir_to_transfer,
                                           rows_to_read=rows_to_read,
                                           cols_to_read=cols_to_read)

                    try:
                        df = df.astype('float')

                    except:
                        pass

                    try:
                        os.makedirs(f'{root_dir}{relative_folder_directory}')

                    except FileExistsError:
                        pass

                    try:
                        if folder_dir_index == 0:
                            print(
                                f'Appending data to file at path: {xlsx_file_path_to_write}\n'
                            )

                        # Append dataframe to main excel file
                        append_df_to_excel(df=df,
                                           filename=xlsx_file_path_to_write,
                                           startrow=to_write_start_row + 2,
                                           startcol=start_col_to_write)

                    except:
                        print(
                            f'Failed to write to excel file. Ensure that the target file path "{xlsx_file_path_to_write}" is not running/open.\n'
                        )
                        sys.exit()

        excel_file_paths_to_open.append(xlsx_file_path_to_write)

    for file_path in excel_file_paths_to_open:
        try:
            print(f'\n\nOpening excel file at path: {file_path}\n\n')

            # Open the new Excel file after data is written to it
            execute_powershell(f'Invoke-Item \"{file_path}\"')

        except:
            print(
                f'Failed to open xlsx file at path: "{xlsx_file_path_to_write}"\n'
            )
            continue
