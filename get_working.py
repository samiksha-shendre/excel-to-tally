from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter
from datetime import datetime
import functions

 

return_month = "082026"
dt = datetime.strptime(return_month, "%m%Y")
return_month_in_words = dt.strftime("%B")
return_year = dt.strftime("%Y")
gstr_2B_download_date = "15092026"
_gst_no = "27AAXFB1785G1ZY"
firm_name = "BEEJAL INFRA"

gstr_2B_filename = f"{return_month}_{_gst_no}_GSTR2B_{gstr_2B_download_date}.xlsx"
gstr_2B_wb = load_workbook(gstr_2B_filename, data_only=True)
working_wb_name = f"GST WORKING_{firm_name}_{return_month_in_words.upper()} {return_year}.xlsx"
working_wb = load_workbook(working_wb_name)

b2b_sheet = gstr_2B_wb["B2B"]

heading_bg_color = PatternFill(start_color="538DD5", fill_type="solid")
total_bg_color = PatternFill(start_color="FFC000", fill_type="solid")

for row in working_wb_purchases_sheet.iter_rows(min_row=1, min_col=1, max_col=14):
    for cell in row:
        cell.fill = heading_bg_color




#function to assign particular cell the particular value
def assign_value_to_cell(a,b,c):
    working_wb_purchases_sheet.cell(row=a,column=b,value=c)

heading_names = ["SR. NO.", "INVOICE NO.", "DATE", "PARTY NAME", "GST NO.", "NET AMOUNT", "%", "CGST", "%", "SGST", "%","IGST", "R/OFF", "TOTAL AMOUNT"]

for row in working_wb_purchases_sheet['A1:A14']:
    for cell in row:
        assign_value_to_cell(cell.column,cell.row,heading_names[cell.row-1])

for row in range(7, b2b_sheet.max_row + 1):
    # Skip completely empty rows
    if not any(b2b_sheet.cell(row=row, column=col).value for col in range(1, 21)):
        continue

    

    invoice_no = b2b_sheet.cell(row=row, column=3).value or ""
    invoice_date = b2b_sheet.cell(row=row, column=5).value or ""
    creditor_ledger = b2b_sheet.cell(row=row, column=2).value or ""
    gst_no = b2b_sheet.cell(row=row, column=1).value or ""
    basic_amount = b2b_sheet.cell(row=row, column=9).value or ""
    cgst_amount = b2b_sheet.cell(row=row, column=11).value or ""
    sgst_amount = b2b_sheet.cell(row=row, column=12).value or ""
    igst_amount = b2b_sheet.cell(row=row, column=10).value or ""
    invoice_value = b2b_sheet.cell(row=row, column=6).value or ""
    rounding_off_value = int(invoice_value)-int(basic_amount)-int(cgst_amount or 0)-int(sgst_amount or 0)-int(igst_amount or 0)

    assign_value_to_cell(row-4, 2, invoice_value)
    assign_value_to_cell(row-4, 3, invoice_date)
    assign_value_to_cell(row-4, 4, creditor_ledger)
    assign_value_to_cell(row-4, 5, gst_no)
    assign_value_to_cell(row-4, 6, float(basic_amount) if basic_amount else 0)
    assign_value_to_cell(row-4, 7, float(cgst_amount / basic_amount if cgst_amount else 0))
    assign_value_to_cell(row-4, 8, float(cgst_amount) if cgst_amount else 0)
    assign_value_to_cell(row-4, 9, float(sgst_amount / basic_amount if sgst_amount else 0))
    assign_value_to_cell(row-4, 10, float(sgst_amount) if sgst_amount else 0)
    assign_value_to_cell(row-4, 11, float(igst_amount / basic_amount if igst_amount else 0))
    assign_value_to_cell(row-4, 12, float(igst_amount) if igst_amount else 0)
    assign_value_to_cell(row-4, 13, float(invoice_value) if invoice_value else 0)
    assign_value_to_cell(row-4, 14, float(rounding_off_value) if rounding_off_value else 0)

for row_idx in range(1, working_wb_purchases_sheet.max_row + 3):
    working_wb_purchases_sheet.row_dimensions[row_idx].height = 20

for col in working_wb_purchases_sheet.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)  

    for cell in col:
        if cell.value is not None:
            # Cast the value to a string to check its character length
            cell_len = len(str(cell.value))
            if cell_len > max_len:
                max_len = cell_len
                
    working_wb_purchases_sheet.column_dimensions[col_letter].width = max_len + 3


center_alignment = Alignment(horizontal="center", vertical="center")
right_alignment = Alignment(horizontal="right", vertical="center")
last_row = working_wb_purchases_sheet.max_row + 2
working_wb_purchases_sheet[f'A{last_row}'] = "TOTAL"
working_wb_purchases_sheet.merge_cells(start_row = last_row, start_column=1, end_row=last_row, end_column=5)
for row in working_wb_purchases_sheet.iter_rows(
    min_row=1, max_row=working_wb_purchases_sheet.max_row, min_col=1, max_col=working_wb_purchases_sheet.max_column
):
    for cell in row:
        cell.alignment = center_alignment

#give sr. no.
for i in range(3, working_wb_purchases_sheet.max_row-1):
  working_wb_purchases_sheet[f"A{i}"] = i-2

#give orange background color to last row of total
for row in working_wb_purchases_sheet.iter_rows(min_row=working_wb_purchases_sheet.max_row, min_col=1, max_col=14, max_row= working_wb_purchases_sheet.max_row):
    for cell in row:
        cell.fill = total_bg_color
    
#give 2 decimal places rounding off and right alignment to rows from taxable value to invoice value
for row in working_wb_purchases_sheet.iter_rows(min_row=3, max_row=working_wb_purchases_sheet.max_row, min_col=6, max_col=14):
    for cell in row:
        if isinstance(cell.value, (int, float)):
            cell.number_format = '0.00'
            cell.alignment = right_alignment

#formatting for the Month cell B2
working_wb_purchases_sheet["B2"] = return_month_in_words
working_wb_purchases_sheet["B2"].fill = PatternFill(start_color="92D050", fill_type="solid")
working_wb_purchases_sheet["B2"].font = Font(bold=True)

#function to calculate total and assign value to cells
def getSum(a):
    working_wb_purchases_sheet[f"{a}{working_wb_purchases_sheet.max_row}"].value = f"=SUM({a}2:{a}{working_wb_purchases_sheet.max_row-1})"

getSum("F")
getSum("H")
getSum("J")
getSum("L")
getSum("M")
getSum("N")



give_column_width("G",7)
give_column_width("I",7)
give_column_width("K",7)
give_column_width("M",9)



#save the workbook
working_wb.save(working_wb_name)

