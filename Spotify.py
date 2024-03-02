import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

url = "https://charts.spotify.com/charts/view/regional-global-daily/latest"

# Realiza la solicitud a la página web y obtén el contenido
response = requests.get(url)
html_content = response.content

# Analiza el contenido HTML con BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

# Crea un nuevo libro de Excel y selecciona la hoja activa
workbook = Workbook()
sheet = workbook.active

class_table='Table__TableElement-evwssh-0 jaKCLL styled__StyledTable-sc-135veyd-7 keAEM'

# Encuentra los elementos en la página web que deseas copiar
# (esto es solo un ejemplo, debes adaptarlo a tu caso específico)
table = soup.find('table', class_=class_table)

if table:
# Itera sobre las filas y celdas de la tabla y copia la información a Excel
    for row_num, row in enumerate(table.find_all('tr'), start=1):
        for col_num, cell in enumerate(row.find_all(['td', 'th']), start=1):
            sheet.cell(row=row_num, column=col_num, value=cell.text)

else:
     print('No se encontró ninguna tabla en la página.')
     
# Guarda el libro de Excel
workbook.save('spotify.xlsx')