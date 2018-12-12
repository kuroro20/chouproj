from openpyxl import load_workbook

wb2 = load_workbook('PDM.xlsx')
sheet_ranges = wb2['range names']
print (sheet_ranges)