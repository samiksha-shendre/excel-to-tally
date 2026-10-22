import os
from pathlib import Path
from openpyxl import load_workbook
from datetime import datetime

#converts date to "20261019" i.e., "YYYYmmdd" format
def format_date(value):
    if not value:
        return ""
    if isinstance(value, str):
        return datetime.strptime(value, "%d/%m/%Y").strftime("%Y%m%d")
    return value.strftime("%Y%m%d")

# get 2B file of particular company
def get_2B_file(return_month, gstNo, gstr_2B_download_date):
    gstr_2B_filename = f"{return_month}_{gstNo}_GSTR2B_{gstr_2B_download_date}.xlsx"
    folder_path = Path("./CURRENT MONTH GST WORKINGS/GSTR-2B")

    if folder_path.exists():
        for file_path in folder_path.glob("*.xlsx"):
            if gstNo in file_path.name:
                return load_workbook(f"{folder_path}/{gstr_2B_filename}", data_only=True)

# get working file of particular company
def get_working_file(companyName, month, year):
    working_filename = f"GST WORKING_{companyName}_{month} {year}.xlsx"
    folder_path = Path("./CURRENT MONTH GST WORKINGS/WORKINGS")

    if folder_path.exists():
        for file_path in folder_path.glob("*.xlsx"):
            if companyName in file_path.name:
                return [load_workbook(f"{folder_path}/{working_filename}", data_only=True), folder_path]

def give_bg_color(sheet_name, min_row, max_row, min_col, max_col, bg_color):
    for row in sheet_name.iter_rows(min_row=min_row, max_row = max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.fill = bg_color





# get_2B_file("082026", "27AAXFB1785G1ZY", "15092026")







#function to give specific a column width of b
# def give_fixed_column_width(a, b):
#     working_wb_purchases_sheet.column_dimensions[a].width = b