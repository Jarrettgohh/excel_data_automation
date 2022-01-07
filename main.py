import subprocess
import sys
import re

from option_2 import option_2
from functions import execute_powershell, execute_powershell_function

print('-----------------------------')
print('\nRemeber to edit the config.json file\n')
print('-----------------------------')
print('-----------------------------')
print('\nRun options\n')
print('-----------------------------')
print('\n')
print(
    '1. Transfer data from text (.txt) file to excel (.xlsx) file. Select this option if you wish to update the config.json file according to a new excel file cell format.'
)
print(
    '2. Extract data from .txt, .csv or .xls file into a .xlsx file in a certain format.'
)

print('-----------------------------')
user_selection = input('Enter your choice: ')
print('\n')

if user_selection == "1":
    sys.exit()

elif user_selection == "2":
    option_2()

elif user_selection == "3":

    fname = 'C:/hry Users/gohja/Desktop/ASTAR intenship data/CV Hysteresis_LOT3/PV/HZO_LOT3_WAFER1_PV/convert_xls_to_xlsx_test'

    try:
        matches = re.findall(r"/*[\w|\s]+/*", fname)

        folder_dir = ''

        for index, match in enumerate(matches):
            if re.search('\s', match):
                slash_index = re.findall(r'/', match)
                match = '"' + match.replace('/', '') + '"'
                match = '/' + match + '/' if len(slash_index) == 2 else (
                    '/' + match if slash_index == 0 else match + '/')

            folder_dir = folder_dir + match

        print(folder_dir)

    except:
        pass
