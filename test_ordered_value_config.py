import re

from copy import copy

ordered_values_config = [
    "100Hz", "200Hz", "300Hz", "500Hz", "1000Hz", "1.2V", "1.5V", "3V", "5V"
]
# "2V" matches "1.2V" too -> to fix
files_to_order = [
    "500Hz_3V", "200Hz_5V", "300Hz_1.2V", "100Hz_3V", "1000Hz_1.5V",
    "100Hz_1.2V", "100Hz_5V", "100Hz_1.5V", "200Hz_1.2V", "200Hz_3V",
    "500Hz_1.2V", "200Hz_1.5V", "random", "1000Hz_1.2V"
]



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
file_expected_index = 0

for file in filtered_files_to_order:

    for ordered_file_index, ordered_file in reversed(
            list(enumerate(ordered_files))):

        # If same file name
        if ordered_file == file:
            file_expected_index = ordered_file_index
            continue

        pos_status = None

        for field in ordered_value_fields:
            ordered_file_matches = ordered_value_fields[field]

            if ordered_file in ordered_file_matches and file in ordered_file_matches:
                continue

            elif ordered_file not in ordered_file_matches and file not in ordered_file_matches:
                continue

            elif ordered_file not in ordered_file_matches:
                # Break this loop but continue in the outer loop
                pos_status = '0'
                file_expected_index = ordered_file_index

                break

            elif file not in ordered_file_matches:

                # Break this loop and the outer loop too
                pos_status = '1'
                file_expected_index = ordered_file_index + 1

                break

        else:
            continue  # only executed if the inner loop did NOT break

        # Only executed if the inner loop DID break
        if pos_status == None:
            continue

        else:
            if pos_status == '1':
                break

    prev_index = ordered_files.index(file)

    ordered_files.insert(file_expected_index, file)
    del ordered_files[prev_index]
