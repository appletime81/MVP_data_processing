import re
import os
import shutil
import time
import numpy as np
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
    "Defect_Type",
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
        if file_name.endswith(".xlsx"):
            all_file_names.append(file_name)
    return all_file_names


def get_all_excel_file_col_name(path):
    all_file_names = get_all_file_names(path)
    all_diff_col_names_list = []
    for i, file_name in enumerate(all_file_names):
        print(f"{file_name}")
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


def combine_all_excel_data(excel_file_names, col_names):
    data_dict = dict([(col_name, []) for col_name in col_names])

    for i, file_name in enumerate(excel_file_names):
        data = pd.read_excel(os.path.join(path, file_name), sheet_name=0, header=0)
        # get column name
        data.dropna(axis=0, how="all", inplace=True)
        tmp_col_names = ["Database"] + data.iloc[:, 0].to_list()
        tmp_col_names = get_column_names(tmp_col_names)
        tmp_values = [data.columns.values.tolist()[1]] + data.iloc[:, 1].to_list()
        tmp_data_dict = dict(zip(tmp_col_names, tmp_values))
        for col_name in col_names:
            if col_name in tmp_data_dict.keys():
                data_dict[col_name].append(tmp_data_dict[col_name])
            else:
                data_dict[col_name].append(np.nan)

    # convert dictionary to dataframe
    data_df = pd.DataFrame(data_dict)

    return data_df


def combine_old_excel_and_new_data(old_data, new_data):
    old_col_names = old_data.columns.values.tolist()
    new_col_names = new_data.columns.values.tolist()

    # old data row numbers
    old_row_nums = len(old_data)

    # new data row numbers
    new_row_nums = len(new_data)

    col_names = list(set(old_col_names + new_col_names))
    col_names = ORIG_COL_NAMES + list(set(col_names).difference(set(ORIG_COL_NAMES)))
    data_dict = dict([(col_name, []) for col_name in col_names])

    for col_name in col_names:
        if col_name in old_col_names and col_name in new_col_names:
            data_dict[col_name] += (
                old_data[col_name].to_list() + new_data[col_name].to_list()
            )
        if col_name in new_col_names and col_name not in old_col_names:
            data_dict[col_name] += [np.nan] * old_row_nums + new_data[col_name].to_list()
        if col_name in old_col_names and col_name not in new_col_names:
            data_dict[col_name] += old_data[col_name].to_list() + [np.nan] * new_row_nums

    # convert dictionary to dataframe
    combined_data_df = pd.DataFrame(data_dict)
    return combined_data_df


if __name__ == "__main__":
    start = time.time()

    # set output file name
    output_file_name = "output.xlsx"

    # set current directory
    path = os.getcwd()

    # Step1: Get all column names

    all_col_names = get_all_excel_file_col_name(path)
    pprint(all_col_names)  # test print

    # Step2: Combine all excel data
    all_excel_file_names = get_all_file_names(path)
    data_df = combine_all_excel_data(all_excel_file_names, all_col_names)

    # Step3: move excel files to Old_Raw_Date folder
    for file_name in all_excel_file_names:
        shutil.move(
            os.path.join(path, file_name), os.path.join(path, "Old_Raw_Date", file_name)
        )

    # Step4: Combine old excel and new data, if old data(file name: output.xlsx) exists
    if os.path.exists(os.path.join(path, output_file_name)):
        combined_data_df = combine_old_excel_and_new_data(
            pd.read_excel(os.path.join(path, output_file_name), sheet_name=0, header=0),
            data_df,
        )
        combined_data_df.to_excel(os.path.join(path, output_file_name), index=False)
        data = pd.read_excel(output_file_name)
        data.fillna(0, inplace=True)
        data.to_excel(output_file_name, index=False)
    else:
        data_df.to_excel(os.path.join(path, output_file_name), index=False)
        data = pd.read_excel(output_file_name)
        data.fillna(0, inplace=True)
        data.to_excel(output_file_name, index=False)

    print("Time: {}".format(time.time() - start))
