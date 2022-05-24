import os
import pandas as pd





data = pd.read_excel("output.xlsx")
data.fillna(0, inplace=True)
data.to_excel("output.xlsx", index=False)