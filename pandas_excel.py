import pandas as pd

data = pd.read_excel('./R100_D5.xlsx', skiprows=27, usecols=[2])
print(data)
