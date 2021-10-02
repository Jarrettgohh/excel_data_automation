import sys

lines = sys.stdin.readlines()

file_names = []
folder_name = ''


for line in lines:
    line = line.replace('\n', '')

    # To check if its .xlsx file and add in array
    if '.xlsx' in line:
        file_names.append(line)
        continue

     # Folder name
    else:
        folder_name = line
        continue
