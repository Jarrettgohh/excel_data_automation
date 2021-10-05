import subprocess
import json

from excel_cell_ranges import get_excel_cell_ranges


config_json = open('../config.json', 'r')
config_json = json.load(config_json)


def execute_powershell(command: str):
    subprocess.Popen(
        ['powershell.exe', command])


cell_ranges = config_json['cell_ranges']
cell_range_with_row_list = get_excel_cell_ranges(
    row_number=13, cell_ranges=cell_ranges)


print(cell_range_with_row_list)

# target_file_directory = 'C:\\Users\gohja\Desktop\\excel_data_automation\\data_calculations.xlsm'

# powershell_command = f"../PowerShell/macro.ps1 -target_file_directory {target_file_directory}"
# execute_powershell(powershell_command)
