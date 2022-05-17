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
        all_file_names.append(file_name)
    return all_file_names


def get_all_column_names(file_name):
    global ALL_COL_NAMES
    data = pd.read_excel(file_name)
    col_names = get_column_names(file_name)


if __name__ == "__main__":
    column_name_file = "MVP_Data_Summary.xlsx"
    col_names = get_column_names(column_name_file)
    # pprint(col_names)

    all_excel_file_names = get_all_file_names(os.getcwd())
    all_excel_file_names = [
        file_name for file_name in all_excel_file_names if file_name != column_name_file
    ]
