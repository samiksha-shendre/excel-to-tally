import openpyxl
from playwright.sync_api import sync_playwright

file_name = "GST LOGIN ID AND PASSWORD.xlsx"
sheet_name = "MESSAGES"
name_start_row = 10
name_end_row = 50

wb = openpyxl.load_workbook(file_name, data_only=True)
sheet = wb[sheet_name]

for i in range(name_start_row, name_end_row+1):
    wp_link = sheet.cell(row=i, column=3).value
    print(wp_link)
    
    with sync_playwright() as p:
    # Launch Chromium (can also use firefox or webkit)
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
    
        page.goto(wp_link)
        page.screenshot(path=f"example_playwright{i}.png")
    
        browser.close()