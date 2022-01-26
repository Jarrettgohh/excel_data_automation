import re

from copy import copy

ordered_values_config = ["100Hz", "200Hz", "1.2V", "1.5V"]
# "2V" matches "1.2V" too -> to fix
files_to_order = ["100Hz_1.5V", "200Hz_1.2V", "200Hz_1.5V", "100Hz_1.2V"]

filtered_files_to_order = []
ordered_value_fields = {}

for config in ordered_values_config:
    ordered_value_fields[config] = []

for file in files_to_order:
    for config in ordered_values_config:

        match = re.search(f'(^|\-|\_|\.){config}(\-|\_|\.|$)', file)

        if match:
            ordered_value_field_list = ordered_value_fields[config]
            ordered_value_field_list.append(file)

            ordered_value_fields[config] = ordered_value_field_list

            if file not in filtered_files_to_order:
                filtered_files_to_order.append(file)

ordered_files = copy(filtered_files_to_order)
# print(ordered_value_fields)
file_expected_index = 0

for file in filtered_files_to_order:
    if file != '200Hz_1.5V':
        break
    print('\n')
    print(f'file: {file}')
    for ordered_file_index, ordered_file in reversed(
            list(enumerate(ordered_files))):

        print(f'ordered file: {ordered_file}')

        # If same file name
        if ordered_file == file:
            continue

        for field in ordered_value_fields:
            ordered_file_matches = ordered_value_fields[field]

            if ordered_file in ordered_file_matches and file in ordered_file_matches:
                print('pass')
                continue

            elif ordered_file not in ordered_file_matches and file not in ordered_file_matches:
                print('pass')
                continue

            elif ordered_file not in ordered_file_matches:
                print('0')
                file_expected_index = ordered_file_index
                # ordered_files.remove(file)
                # ordered_files.insert(ordered_file_index, file)
                # break

            elif file not in ordered_file_matches:
                print('1')
                file_expected_index = ordered_file_index + 1
                # ordered_files.remove(file)
                # ordered_files.insert(ordered_file_index + 1, file)
                break

        else:
            continue  # only executed if the inner loop did NOT break
        break  # only executed if the inner loop DID break

    print(ordered_files)
    # print(file)
    print(str(file_expected_index) + '\n')
    ordered_files.remove(file)
    ordered_files.insert(file_expected_index, file)
    print(ordered_files)

# print(ordered_files)
