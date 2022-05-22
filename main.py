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


def get_all_excel_file_col_name(path):
    all_file_names = get_all_file_names(path)
    all_diff_col_names_list = []
    for i, file_name in enumerate(all_file_names):
        data = pd.read_excel(os.path.join(path, file_name), sheet_name=0, header=0)
        data.dropna(axis=0, how="all", inplace=True)
        tmp_col_names_list = ["Database"] + data.iloc[:, 0].to_list()
        tmp_col_names_list = get_column_names(tmp_col_names_list)
        tmp_col_names_set = set(tmp_col_names_list)

        # get difference between original column names and current column names
        diff_col_names_list = list(tmp_col_names_set.difference(set(ORIG_COL_NAMES)))

        # get all different column names
        all_diff_col_names_list += diff_col_names_list

    all_diff_col_names_list = list(set(all_diff_col_names_list))
    all_col_names = ORIG_COL_NAMES + all_diff_col_names_list
    return all_col_names


def convert_to_dataframe(data):
    data_dict = {}

    pass






if __name__ == "__main__":
    start = time.time()

    # set output file name
    output_file_name = "test.xlsx"

    # Step1: Get all column names
    path = os.getcwd()
    all_col_names = get_all_excel_file_col_name(path)
    pprint(all_col_names)  # test print

    # Step2: Convert all excel files to dataframe










