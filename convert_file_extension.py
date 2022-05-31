import win32com.client as win32
import os


def convert_xls_to_xlsx(file_name):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(os.getcwd() + "\\" + file_name)

    wb.SaveAs(
        os.getcwd() + f"\\{file_name}x", FileFormat=51
    )  # FileFormat = 51 is for .xlsx extension, FileFormat = 56 is for .xls extension
    wb.Close()
    excel.Application.Quit()


if __name__ == "__main__":
    for file in os.listdir(os.path.dirname(__file__)):
        if file.endswith(".xls"):
            print(f"Converting {file} to {file}x")
            try:
                convert_xls_to_xlsx(file)
            except Exception:
                print(f"Error converting {file}")
