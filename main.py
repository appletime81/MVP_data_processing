import re
import os
import pandas as pd

from pprint import pprint

ALL_COL_NAMES = list()


def get_column_names(file_name):
    data = pd.read_excel(file_name)
    col_names = data.columns.values.tolist()
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


def get_all_column_names(file_names, col_names):
    """
    :param file_names: all_excel_file_names
    :param col_names: get from MVP_Data_Summary.xlsx
    :return: pd.DataFrame
    """
    new_col_names = list()
    for file_name in file_names:
        tempdata = pd.read_excel(file_name)
        tempdata.dropna(inplace=True)
        temp_col_names = (["Database"] + tempdata.iloc[:, 0].to_list())[29:]
        for col_name in temp_col_names:
            new_col_names.append(col_name)
    new_col_names = col_names + list(set(new_col_names))
    pprint(new_col_names)


if __name__ == "__main__":
    column_name_file = "MVP_Data_Summary.xlsx"
    all_excel_file_names = get_all_file_names(os.getcwd())
    get_all_column_names(all_excel_file_names, get_column_names(column_name_file))
