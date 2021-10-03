import pandas as pd
import json


def excel_read_col_row(excel_file, row_col_dict):

    rows = row_col_dict['rows']
    rows_range = range(rows[0]-2, rows[-1]-1)

    usecols = row_col_dict['cols']

    df = pd.read_excel(excel_file,  usecols=usecols)

    df = df.loc[rows_range]

    print(df)


json_file = open('config.json', 'r')
row_col_dict = json.load(json_file)
row_col_dict = row_col_dict['index_to_read']


excel_read_col_row('pandas.xlsx', row_col_dict=row_col_dict)

# df = data frame
# skiprows (-2)
# usecols row A = pandas col 0
# df = pd.read_excel('pandas.xlsx',
#                    # skiprows=25,
#                    usecols=[2])


# COL
# -> Row A in Excel is considered as row index 0 in pandas


# ROW
# -> df row given from pandas is (-2) index for the row number
# -> The second row index (row 2 in Excel) starts from 0 in pandas
# -> Due to pandas counting the row 1 in Excel to be the header


# range = for each index (-skiprows value given)
# loc = df.loc[range(0, 10)]
# print(df)
# print(data_frame.loc[27])
