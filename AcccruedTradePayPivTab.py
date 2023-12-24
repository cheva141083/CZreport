import pandas as pd
import os
from datetime import datetime
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

    #Get the name of first file
    file_name = os.path.basename(first_file).split('.')[0]
    print(file_name)


    # Open the excel file con OpenPyXL
    book = load_workbook(first_file)

    # Operations to apply in file
    worksheet = book['GL']
    print(worksheet.title)

    #Add new column 
    new_column = 'AX'

    #Cell reference
    cell_new_column = worksheet[new_column + '1' ]

    #Name new column
    new_column_name = "Aging in days"
    cell_new_column.value = new_column_name

    max_row = worksheet.max_row

    for row in range(2, max_row + 1):  # Assuming the data starts from the second row
    # Get the date from column X
        date_cell = worksheet.cell(row=row, column=24)  # Column X is the 24th column
        cell_value = date_cell.value
        if cell_value is not None:
            date_value = datetime.strptime(cell_value, "%m/%d/%Y")
            print(f"Type {type(date_value)}")
            print(date_value)

            if date_value is not None and isinstance(date_value, datetime):
                # Calculate the number of days between the date and 11/30/2023
                target_date = datetime(2023, 11, 30)
                print(target_date)
                days_difference = (target_date - date_value).days
                print(days_difference)

                # Update the value in column AX (50th column)
                worksheet.cell(row=row, column=50, value=days_difference)


    #Save file
   
    name_specifications = folderUbication + "/" + file_name + " modified" + ".xlsx"
    book.save(name_specifications)



    
    # Cierra el libro después de trabajar con él
    book.close()
else:
    print("There are not excel files in this folder")