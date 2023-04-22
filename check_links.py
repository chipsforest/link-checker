import sys
import pandas as pd
import requests
from tqdm import tqdm
import xlsxwriter

# Check if the Excel file name is provided as an argument
if len(sys.argv) != 2:
    print('Usage: python script.py [23-WFS-creatives-briefing-social.xlsx]')
    sys.exit()

# Load the Excel file
excel_file = pd.ExcelFile(sys.argv[1])

# Initialize variables to store URLs and their statuses
urls = []
statuses = []

# Loop through each sheet and cell to extract the URLs
for sheet_name in excel_file.sheet_names:
    sheet = excel_file.parse(sheet_name)
    for i, row in sheet.iterrows():
        for cell in row:
            if isinstance(cell, str) and cell.startswith('http'):
                urls.append(cell)

# Send a GET request to each URL to check its status code
for url in tqdm(urls):
    response = requests.get(url)
    statuses.append(response.status_code)

# Create a new Excel workbook and write the results to a new sheet
workbook = xlsxwriter.Workbook('results.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'URL')
worksheet.write('B1', 'Status Code')
for i, (url, status) in enumerate(zip(urls, statuses), start=1):
    row = i + 1
    worksheet.write(f'A{row}', url)
    worksheet.write(f'B{row}', status)

    # Color the status code cell red for 404s and green for other codes
    cell_format = workbook.add_format()
    if status == 404:
        cell_format.set_bg_color('#FFC7CE')  # Light red
    else:
        cell_format.set_bg_color('#C6EFCE')  # Light green
    worksheet.write(f'B{row}', status, cell_format)

# Close the workbook to save the changes
workbook.close()
