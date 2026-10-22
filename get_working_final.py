from openpyxl import load_workbook
from GST_NO_AND_NAMES import companies
import functions

details_wb = load_workbook("./CURRENT MONTH GST WORKINGS/DETAILS.xlsx", data_only=True)
details_sheet = details_wb["Sheet1"]
return_month = details_sheet["B8"].value
gstr_2B_download_date = details_sheet["B9"].value
gst_month = details_sheet["B1"].value
gst_year = details_sheet["B2"].value

for company in companies:
    compName = company["name"]
    gstNo = company["GST No."]

    wb_2B = functions.get_2B_file(return_month, gstNo, gstr_2B_download_date)
    wb_working_details = functions.get_working_file(compName, gst_month, gst_year)
    wb_working = wb_working_details[0]
    file_path = wb_working_details[1]
    wb_working_filename = f"{file_path}/GST WORKING_{compName}_{gst_month} {gst_year}.xlsx"

    sheet_2B = wb_2B["B2B"]
    sheet_wb = wb_working.create_sheet("PURCHASES")

    m = sheet_2B.cell(row=5, column=3).value
    sheet_wb.cell(row=2,column=2,value=m)

    wb_working.save(wb_working_filename)
