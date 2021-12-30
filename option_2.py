import json
import sys

from functions import transfer_single_csv_to_xlsx
from Excel.excel_functions import xlsx_read_col_row

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

        for folder_dir in relative_folder_directories:
            for file_name in files:

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

                    print(df)

            # for index, file_path in enumerate(files):

            #     if file_type == 'csv':
            #         transfer_single_csv_to_xlsx(
            #             file_path_to_read=
            #             f'{root_dir}{relative_folder_path_to_read}\{file_path}',
            #             folder_dir_to_write=
            #             f'{root_dir}{relative_folder_path_transfer_for_csv_files}',
            #             file_path_to_write=
            #             f'{root_dir}{relative_folder_path_transfer_for_csv_files}\\{file_path.replace("csv", "xlsx")}',
            #         )

            #         excel_file_path_to_read = f'{root_dir}{relative_folder_path_transfer_for_csv_files}\{file_path}'.replace(
            #             'csv', 'xlsx')
            #         df = excel_read_col_row(excel_file=excel_file_path_to_read,
            #                                 rows_to_read=rows_to_read,
            #                                 cols_to_read=cols_to_read)

            #         try:
            #             os.makedirs(folder_dir_to_write)

            #         except FileExistsError:
            #             # directory already exists
            #             pass

            #         try:
            #             df = df.astype('float')

            #         except:
            #             pass

            #         # Append dataframe to main excel file
            #         append_df_to_excel(
            #             df=df,
            #             filename=
            #             f'{folder_dir_to_write}\\{excel_file_name_to_write}',
            #             startrow=start_row_to_write,
            #             startcol=hardcode_cols_to_write[index])

            #     elif file_type == 'xlsx':
            #         df = excel_read_col_row(excel_file=file_path,
            #                                 rows_to_read=rows_to_read,
            #                                 cols_to_read=cols_to_read,
            #                                 sheet_name="Sheet1")
