import sys
import pandas as pd

lines = sys.stdin.readlines()

file_names = []
directory = ''


for line in lines:
    line = line.replace('\n', '')

    # To check if its .xlsx file and add in array
    if '.xlsx' in line:
        file_names.append(line)
        continue

     # Folder name
    else:
        directory = line
        continue

print(f'folder name: {directory}')
print(f'file names: {file_names}')


data = pd.read_excel(f'{directory}\{file_names[0]}')

print(data)
