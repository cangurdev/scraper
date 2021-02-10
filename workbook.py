import openpyxl
from openpyxl.worksheet.dimensions import ColumnDimension

#Getting existing file 
wb = openpyxl.load_workbook('Product Detail URL.xlsx') #Gets xlsx file
sheet = wb['Sheet4'] #Gets sheet

#Creating new file for results
newWorkbook = openpyxl.Workbook()
newSheet = newWorkbook.active
dest_filename = 'results.xlsx'

#Setting columns width
newSheet.column_dimensions["A"].width = 75.0
newSheet.column_dimensions["B"].width = 18.0
newSheet.column_dimensions["C"].width = 35.0
newSheet.column_dimensions["D"].width = 20.0
newSheet.column_dimensions["E"].width = 10.0
newSheet.column_dimensions["F"].width = 12.0

#Setting columns name
newSheet['A1'] = "URL"
newSheet['B1'] = "BRAND"
newSheet['C1'] = "NAME"
newSheet['D1'] = "CODE"
newSheet['E1'] = "PRICE"
newSheet['F1'] = "AVAILABILITY"