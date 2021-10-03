import pandas as pd


def excel_read_col_row(excel_file, ):
    df = pd.read_excel(excel_file, skiprows=25, usecols=[2])


# df = data frame
# skiprows (+1)
# usecols row A = pandas col 0
df = pd.read_excel('pandas.xlsx', skiprows=25, usecols=[2])
# COL
# -> Row A in Excel is considered as row index 0 in pandas


# ROW
# -> df row given from pandas is (-2) index for the row number
# -> The second row index (row 2 in Excel) starts from 0 in pandas
# -> Due to pandas counting the row 1 in Excel to be the header


# range = for each index (-skiprows value given)
loc = df.loc[range(0, 10)]
print(loc)
# print(data_frame.loc[27])
