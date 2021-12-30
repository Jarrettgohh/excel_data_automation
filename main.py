import sys

from option_2 import option_2

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