 import openpyxl
 from openpyxl.styles import Font
 import os

 os.chdir(r"/Users/lisatingchang/Downloads")
 wb = openpyxl.load_workbook('produceSales.xlsx')
 sheet = wb.worksheets[0]

 price_updates_dict = {'Garlic': 3.07,
                        'lemon': 1.27}

for rowNum in range(2, sheet.max_row, 1): 
    produceName = sheet.cell(rowNum, 1).value
    if produceName in price_updates_dict:
        sheet.cell(rowNum, 2).value = price_updates.dict[produceName] #改哪一行
        sheet.cell(rowNum, 2).font = Font(color='FF0000')  #改字體

wb.save('produceSales_updates.xlsx') #另存新檔