import json
import sys
import os

from functions import transfer_single_csv_to_xlsx
from Excel.excel_functions import append_df_to_excel, xlsx_read_col_row

#
# Extract data from .txt, .csv or .xls file into a .xlsx file in a certain format
#

config_json = open('config.json', 'r')
config_json = json.load(config_json)

option_2_configs = config_json['option_2_configs']


def option_2():

    for config in option_2_configs:

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
        folder_directory_to_write = to_write['folder_directory']
        xlsx_file_name_to_write = to_write['file_name']
        to_write_row_settings = to_write['row_settings']
        to_write_col_settings = to_write['col_settings']

        for folder_dir_index, folder_dir in enumerate(
                relative_folder_directories):
            for file_index, file_name in enumerate(files):

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
                        os.makedirs(folder_directory_to_write)

                    except FileExistsError:
                        pass

                    to_write_start_col_setting = to_write_col_settings[
                        'start_col']
                    to_write_cols = to_write_col_settings['cols']

                    cols_to_read_len = len(cols_to_read)
                    start_col_to_write = (
                        (folder_dir_index * len(files)) * cols_to_read_len
                    ) + to_write_start_col_setting + (
                        cols_to_read_len * file_index
                    ) if to_write_cols == 'auto' else to_write_cols[file_index]

                    # Append dataframe to main excel file
                    append_df_to_excel(
                        df=df,
                        filename=
                        f'{folder_directory_to_write}\\{xlsx_file_name_to_write}',
                        startrow=to_write_row_settings['start_row'],
                        startcol=start_col_to_write)
