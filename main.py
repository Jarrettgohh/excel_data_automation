import subprocess
import sys
import json
import pandas as pd
import re

from typing import OrderedDict
from pandas.core.frame import DataFrame
from excel_functions import append_df_to_excel

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
excel_file_to_write = f'{directory}\Test_3_data_calculations.xlsx'


def excel_read_col_row(excel_file, row_col_dict):
    rows = row_col_dict['rows']
    cols = row_col_dict['cols']

    return pd.read_excel(excel_file,
                         skiprows=rows[0]-1,
                         skipfooter=rows[-1]-rows[0], usecols=cols)


def append_to_new_excel(df: DataFrame, **args):
    try:
        append_df_to_excel(filename=excel_file_to_write,
                           df=df, **args
                           )
    except Exception as e:
        print(e)
        print(
            'Failed to write to new Excel file. Make sure that the Excel file is not open.')


def execute_powershell(command: str):
    subprocess.Popen(
        ['powershell.exe', command])


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


capacitance_values_ordered = OrderedDict(capacitance_values_dict)
capacitance_values_ordered_list = list(capacitance_values_ordered.keys())

sheet_name = 'Calculations_1'

for index, device_size in enumerate(capacitance_values_dict):
    devices_in_each_size = capacitance_values_dict[device_size]

    if index == 0:
        col_to_skip = index

    else:
        index_prev_device = index - 1

        prev_device_dict_key = capacitance_values_ordered_list[index_prev_device]
        prev_device_dict = capacitance_values_dict[prev_device_dict_key]

        col_to_skip = 0

        for i in range(index):

            prev_device_dict_key = capacitance_values_ordered_list[index - (
                i + 1)]
            prev_device_dict = capacitance_values_dict[prev_device_dict_key]
            col_to_skip += len(prev_device_dict)

    # Writing the device size at the top left of each data section
    device_size_start_col = col_to_skip + (index * 3)
    capacitance_values_start_col = col_to_skip + (index * 3) + 1

    df = pd.DataFrame(
        data=[device_size],
    )

    append_to_new_excel(df=df,
                        sheet_name=sheet_name,
                        startcol=0 if index == 0 else device_size_start_col,
                        startrow=0)

    df = pd.DataFrame(
        data=devices_in_each_size,
        index=[list(range(1, 11))],
    )

    append_to_new_excel(df=df,
                        sheet_name=sheet_name,
                        startcol=1 if index == 0 else capacitance_values_start_col,
                        startrow=0)

    df = pd.DataFrame(
        data=['Average',
              'Max',
              'Min',
              'Range',
              'Range/Min(%)',
              '',
              'Average k (F/cm^2)',
              'Max k',
              'Min k',
              'Range k',
              'Range/Min(%)'
              ]
    )

    append_to_new_excel(df=df,
                        sheet_name=sheet_name,
                        startcol=0 if index == 0 else device_size_start_col,
                        startrow=config_json['measure_times'] + 1
                        )


# Open the new Excel file after data is written to it
execute_powershell(f'Invoke-Item \"{excel_file_to_write}\"')
