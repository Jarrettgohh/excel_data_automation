import sys
import json
import pandas as pd

lines = sys.stdin.readlines()

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

# Loop through each file in folder
for file_name in file_names:
    json_file = open('config.json', 'r')
    row_col_dict = json.load(json_file)
    row_col_dict = row_col_dict['index_to_read']

    df = excel_read_col_row(
        f'{directory}\{file_name}.xlsx', row_col_dict=row_col_dict)

    capacitance_values_dict[file_name] = df.to_numpy()


for wafer_name in capacitance_values_dict:
    capacitance_data_nested = capacitance_values_dict[wafer_name]

    capacitance_data = []

    for data in capacitance_data_nested:
        capacitance_data.append(data[0])

    capacitance_values_dict[wafer_name] = capacitance_data


print(capacitance_values_dict)
