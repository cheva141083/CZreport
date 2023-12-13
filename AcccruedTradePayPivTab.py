import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Border, Side



# Search a file in a specific directory
folderUbication ='E:/Resources/NAM US/Python/QA Accrued trade payables'

#Get list of files in the folder
filesOnFolder = os.listdir(folderUbication)

#filter excel files
excelFiles = [file for file in filesOnFolder if file.endswith('.xlsx')]

#Check if there are excel files in the folder
if excelFiles:
    # Get the first file 
    first_file = os.path.join(folderUbication, excelFiles[0])

    # Open the excel file con OpenPyXL
    book = load_workbook(first_file)

    # Operations to apply in file
    worksheet = book['GL']
    print(worksheet.title)

    
    # Cierra el libro después de trabajar con él
    book.close()
else:
    print("There are not excel files in this folder")