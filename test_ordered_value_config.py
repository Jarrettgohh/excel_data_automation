import re

ordered_values_config = ["100Hz", "200Hz", "1.2V", "1.5V"]
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

# print(filtered_files_to_order)
print(ordered_value_fields)
