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


def excel_read_col_row(excel_file, row_col_dict):

    rows = row_col_dict['rows']
    cols = row_col_dict['cols']

    return pd.read_excel(excel_file,
                         skiprows=rows[0]-1,
                         skipfooter=rows[-1]-rows[0], usecols=cols)


capacitance_values_dict = {}
device_dimensions = config_json['device_dimensions']

# To sort the different wafer dimensions through the file name
for file_name in file_names:

    for index, device_dimension in enumerate(device_dimensions):
        print(device_dimension)
        # device_dimension_regex = f'\{device_dimension}'
        # print(device_dimension_regex)
        # device_num = re.sub(device_dimension_regex, '', file_name, 1)
        # print(device_num)


# # Loop through each file in folder
# for file_name in file_names:
#     row_col_dict = config_json
#     row_col_dict = row_col_dict['index_to_read']

#     df = excel_read_col_row(
#         f'{directory}\{file_name}.xlsx', row_col_dict=row_col_dict)

#     capacitance_values_dict[file_name] = df.to_numpy()


# capacitance_values_dict (EXAMPLE):
# {
#  'S100_D1':
# [[0.451], [0.545], [0.5435], [0.6969]],
#  'S100_D2':
# [[0.411], [0.525], [0.5455], [0.689]],
# }


# # Loop through each wafer size to
# for wafer_size in capacitance_values_dict:
#     capacitance_data_nested = capacitance_values_dict[wafer_size]

#     capacitance_data = []

#     # Replace nested array with normal array (capacitance_values_dict); for each wafer size
#     for data in capacitance_data_nested:
#         capacitance_data.append(data[0])

#     capacitance_values_dict[wafer_size] = capacitance_data


# device_numbers = []

# for device in capacitance_values_dict:
#     device_regex = '\S100_'

#     device_num = re.sub(device_regex, '', device, 1)

#     device_numbers.append(device_num)


# df = pd.DataFrame(
#     data=capacitance_values_dict,
#     # index=[list(range(1, 11))],
#     # columns=device_numbers
# )

# print(df)


# try:
#     df.to_excel(f'{directory}\\Test_1_data_calculations.xlsx')

# except:
#     print('\n\nFailed to write to new excel file. Make sure the file is not open.')
