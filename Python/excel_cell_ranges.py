import json
import re


def get_excel_cell_ranges(cell_ranges=[], row_number: int = 0):
    config_json = open('../config.json', 'r')
    config_json = json.load(config_json)

    cell_range_with_row_list = []

    for cell_range in cell_ranges:
        cell_range_with_row = dict(config_json['cell_range_with_row'])

        before_colon_regex = ':[A-Z]+'
        after_colon_regex = '[A-Z]+:'

        cell_range_before_colon = re.sub(before_colon_regex, '', cell_range, 1)
        cell_range_after_colon = re.sub(after_colon_regex, '', cell_range, 1)

        cell_range_with_row['cell_select'] = f'{cell_range_before_colon}{row_number}'
        cell_range_with_row['cell_range'] = f'{cell_range_before_colon}{row_number}:{cell_range_after_colon}{row_number}'

        cell_range_with_row_list.append(cell_range_with_row)

    # print(json.dumps(cell_range_with_row_list, indent=2))
    return cell_range_with_row_list
