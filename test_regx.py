# a = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
# a = set(a)
# b = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
# b = set(b)
# res = b.difference(a)
# res.add(21)
# print(res)


import pandas as pd

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

data = pd.read_excel('20_20220516102136_6612907_2346086.xls')
col_names = ["Database"] + data.iloc[:, 0].to_list()



print(data)


