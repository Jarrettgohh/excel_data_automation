import subprocess


def execute_powershell(command: str):
    subprocess.Popen(
        ['powershell.exe', command])


execute_powershell("../excel_macro.ps1")
