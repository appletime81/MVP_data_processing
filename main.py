import re
import os
import time
import pandas as pd

from pprint import pprint

ORIG_COL_NAMES = [
    "Database",
    "AI_Operator",
    "Machine",
    "Start_Date",
    "Start_Time",
    "End_Date",
    "End_Time",
    "Lot",
    "Product_Code",
    "Total_Number_of_Strips",
    "Number_of_Strips_Inspected",
    "Number_of_Strips_Not_Inspected",
    "False_Call_Device_",
    "Number_of_Devices_False_Called",
    "Total_Quantity_of_Devices_Per_Strip",
    "Quantity_Visible_of_Devices_Per_Strip",
    "Devices_Per_Hour_",
    "Number_of_Strips_Inspected.1",
    "Total_Devices_Count",
    "Number_of_Devices_Passed",
    "Number_of_Devices_Rejected",
    "Number_of_Devices_Skipped",
    "Yield_by_Device_",
    "Punch_Count",
    "Punch_Not_Punched",
    "Punch_Algo_Pass",
    "Punch_Algo_Fail",
    "Punch_Algo_No_Result",
    "Number_of_Strips_No_Images",
    "Defect_Type"
]


def get_column_names(col_names):
    for i, col_name in enumerate(col_names):
        if re.search(r"\(\w+\)", col_name):
            col_name = re.sub(r"\(\w+\)", "", col_name)
        if re.search(r"\(\W+\)", col_name):
            col_name = re.sub(r"\(\W+\)", "", col_name)
        col_names[i] = col_name.replace(" ", "_")
    return col_names


def get_all_file_names(path):
    all_file_names = []
    for file_name in sorted(os.listdir(path)):
        if file_name.endswith(".xls"):
            all_file_names.append(file_name)
    return all_file_names


def convert_to_dataframe(data):
    # get original data keys
    data.dropna(inplace=True)
    data_dict = data.to_dict()
    new_data_dict = dict()
    keys = list(data_dict.keys())
    key1 = "Database"
    keys.pop(keys.index(key1))
    key2 = keys[0]

    # generate new dictionary of data
    col_names = get_column_names(list(data_dict[key1].values()))
    for key, value in zip(ORIG_COL_NAMES, list(data_dict[key2].values())):
        new_data_dict[key] = [value]
    new_data_dict[key1] = [key2]

    # convert dictionary to dataframe
    df = pd.DataFrame(new_data_dict)

    # recompose column names
    new_col_names = ORIG_COL_NAMES + list(set(col_names) - set(ORIG_COL_NAMES))

    # rename column names
    df = df.reindex(columns=new_col_names)
    return df


if __name__ == "__main__":
    start = time.time()
    column_name_file = "MVP_Data_Summary.xlsx"
    output_file_name = "test.xlsx"
    # pprint(get_column_names(column_name_file))

    # get all excel file name
    all_excel_file_names = get_all_file_names(os.getcwd())

    # merge all excel
    for i, file_name in enumerate(all_excel_file_names):
        if i == 0:
            data = pd.read_excel(file_name)
            data = convert_to_dataframe(data)
        else:
            tempdata = pd.read_excel(file_name)
            tempdata = convert_to_dataframe(tempdata)
            data = pd.concat([data, tempdata], axis=0)


    # save excel file
    data.to_excel("test.xlsx", index=False)

    # print execution time
    print("Execution time: {} seconds".format(time.time() - start))
    print("Done")
