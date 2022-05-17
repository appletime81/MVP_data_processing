# a = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
# a = set(a)
# b = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
# b = set(b)
# res = b.difference(a)
# res.add(21)
# print(res)


import pandas as pd
import re
from pprint import pprint

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

ORIG_COL_NAMES = ['Database',
                  'AI_Operator',
                  'Machine',
                  'Start_Date',
                  'Start_Time',
                  'End_Date',
                  'End_Time',
                  'Lot',
                  'Product_Code',
                  'Total_Number_of_Strips',
                  'Number_of_Strips_Inspected',
                  'Number_of_Strips_Not_Inspected',
                  'False_Call_Device_',
                  'Number_of_Devices_False_Called',
                  'Total_Quantity_of_Devices_Per_Strip',
                  'Quantity_Visible_of_Devices_Per_Strip',
                  'Devices_Per_Hour_',
                  'Number_of_Strips_Inspected.1',
                  'Total_Devices_Count',
                  'Number_of_Devices_Passed',
                  'Number_of_Devices_Rejected',
                  'Number_of_Devices_Skipped',
                  'Yield_by_Device_',
                  'Punch_Count',
                  'Punch_Not_Punched',
                  'Punch_Algo_Pass',
                  'Punch_Algo_Fail',
                  'Punch_Algo_No_Result',
                  'Number_of_Strips_No_Images',
                  ]


def get_column_names(col_names):
    for i, col_name in enumerate(col_names):
        if re.search(r"\(\w+\)", col_name):
            col_name = re.sub(r"\(\w+\)", "", col_name)
        if re.search(r"\(\W+\)", col_name):
            col_name = re.sub(r"\(\W+\)", "", col_name)
        col_names[i] = col_name.replace(" ", "_")
    return col_names


def convert_to_dataframe(data):
    data.dropna(inplace=True)
    data_dict = data.to_dict()
    new_data_dict = dict()
    keys = list(data_dict.keys())
    key1 = "Database"
    keys.pop(keys.index(key1))
    key2 = keys[0]

    col_names = get_column_names(list(data_dict[key1].values()))

    for key, value in zip(col_names, list(data_dict[key2].values())):
        new_data_dict[key] = [value]
    new_data_dict[key1] = [key2]
    df = pd.DataFrame(new_data_dict)
    new_col_names = ORIG_COL_NAMES + list(set(col_names) - set(ORIG_COL_NAMES))
    df = df.reindex(columns=new_col_names)
    df.to_excel("test.xlsx", index=False)


    # pprint(data_dict)
    # pprint(data.head())


if __name__ == '__main__':
    data = pd.read_excel('20_20220516102136_6612907_2346086.xls')
    convert_to_dataframe(data)
