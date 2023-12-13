import pandas as pd
import os
import warnings
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Border, Side



# Suppress openpyxl Data Validation warning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet._reader")
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet.header_footer")
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.reader.workbook")

# Search a file in a specific directory
folderUbication ='E:/Resources/NAM US/Python/QA Altair to HFM'

#Get list of files in the folder
filesOnFolder = os.listdir(folderUbication)

#filter excel files
excelFiles = [file for file in filesOnFolder if file.endswith('.xlsm')]

#Check if there are excel files in the folder
if excelFiles:
    # Get the first file 
    first_file = os.path.join(folderUbication, excelFiles[0])

    # Open the excel file con OpenPyXL
    book = load_workbook(first_file, read_only=False, keep_vba=True)

    # Operations to apply in file
    
    worksheet = book.active
    print(worksheet)

    # Cierra el libro después de trabajar con él
    book.close()
else:
    print("There are not excel files in this folder")