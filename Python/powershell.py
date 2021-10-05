import subprocess


def execute_powershell(command: str):
    subprocess.Popen(
        ['powershell.exe', command])


target_file_directory = 'C:\\Users\gohja\Desktop\\excel_data_automation\\data_calculations.xlsm'

powershell_command = f"../PowerShell/macro.ps1 -file_directory {target_file_directory}"
execute_powershell(powershell_command)
