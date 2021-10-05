import subprocess
import json

from excel_cell_ranges import get_excel_cell_ranges


config_json = open('../config.json', 'r')
config_json = json.load(config_json)


def execute_powershell(command: str):
    subprocess.Popen(
        ['powershell.exe', command])


def stringify_dict_to_powershell_object(dict_to_convert: dict):
    powershell_object = json.dumps(dict_to_convert).replace(' ', '').replace(
        "\":", "\"=").replace(",", ";").replace("{\"", "{").replace("\"=", "=").replace(";\"", ";")
    return f'@{powershell_object}'


cell_ranges = config_json['cell_ranges']
cell_range_with_row_list = get_excel_cell_ranges(
    row_number=13, cell_ranges=cell_ranges)


target_file_directory = 'C:\\Users\gohja\Desktop\\excel_data_automation\\data_calculations.xlsm'


powershell_array_str = ''


# Generate PowerShell type array
for index, cell_range_with_row in enumerate(cell_range_with_row_list):
    powershell_object = stringify_dict_to_powershell_object(
        cell_range_with_row)

    powershell_array_str += (f', {powershell_object}' if index !=
                             0 else f'@({powershell_object}')


powershell_array_str = f'{powershell_array_str})'
# print(powershell_array_str)

# powershell_array = json.dumps(cell_ranges_with_row_list).replace(
#     "[", "{").replace("]", "}")
# powershell_array = f"@{powershell_array}"

# -cell_ranges_with_row {cell_ranges_with_row_list}
# -target_file_directory {target_file_directory}
powershell_command = f"../PowerShell/macro.ps1 -cell_ranges_with_row {powershell_array_str}"
execute_powershell(powershell_command)
