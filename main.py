import sys
import json
import pandas as pd
# import numpy as np
import re

lines = sys.stdin.readlines()


config_json = open('config.json', 'r')
config_json = json.load(config_json)


file_names = []
directory = ''


for line in lines:
    line = line.replace('\n', '')

    # To check if its a file name rather than the directory by checking for '\'
    if '\\' not in line:
        file_names.append(line)
        continue

     # Folder name
    else:
        directory = line
        continue


# print(directory)
# print(file_names)

# To get the new excel file name from PowerShell
excel_file_to_write = f'{directory}\Test_1_data_calculations.xlsx'


def excel_read_col_row(excel_file, row_col_dict):

    rows = row_col_dict['rows']
    cols = row_col_dict['cols']

    return pd.read_excel(excel_file,
                         skiprows=rows[0]-1,
                         skipfooter=rows[-1]-rows[0], usecols=cols)


# def write_to_excel():

capacitance_values_dict = {}
device_dimensions = list(config_json['device_dimensions'])


# To sort the different wafer dimensions to each key in the dict, through the file name
# To also sort the data to the respective device key
for file_name in file_names:
    row_col_dict = config_json
    row_col_dict = row_col_dict['index_to_read']

    for device_dimension in device_dimensions:
        if device_dimension not in file_name:
            continue

        if device_dimension not in capacitance_values_dict:
            capacitance_values_dict[device_dimension] = {}

        device_dimension_regex = re.escape(f'{device_dimension}_')
        device_num = re.sub(device_dimension_regex, '', file_name, 1)

        df = excel_read_col_row(
            f'{directory}\{file_name}.xlsx', row_col_dict=row_col_dict)

        df_numpy = df.to_numpy()

        device_num_list = []

        for data in df_numpy:
            device_num_list.append(data[0])

        capacitance_values_dict[device_dimension][device_num] = device_num_list

# print(json.dumps(capacitance_values_dict, indent=2))


df = pd.DataFrame(
    data=['R200'],
)


try:
    df.to_excel(excel_file_to_write,
                header=None,
                index=False,
                startcol=0,
                startrow=0
                )

except:
    print('\nFailed to write to new excel file. Make sure the file is not open.')

# for device_size in capacitance_values_dict:
#     devices_in_each_size = capacitance_values_dict[device_size]

#     df = pd.DataFrame(
#         data=devices_in_each_size,
#         index=[list(range(1, 11))],
#     )

#     print(df)

#     try:
#         df.to_excel(excel_file_to_write)

#     except:
#         print('\n\nFailed to write to new excel file. Make sure the file is not open.')
