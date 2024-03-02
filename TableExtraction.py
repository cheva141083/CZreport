import requests
from bs4 import BeautifulSoup
import openpyxl

# URL de la página web
url = 'http://www.mediatraffic.de/tracks.htm'

# Realizar la solicitud HTTP y obtener el contenido de la página
response = requests.get(url)
print(url)

html_content = response.content

# Analizar el HTML con BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

# Especificar la clase de la tabla (ajusta según la página)
width_tabla = '600'



# Encontrar la tabla por su clase
tabla = soup.find('table', {'width': width_tabla})

# Crear un nuevo libro de Excel y una hoja de trabajo
workbook = openpyxl.Workbook()
sheet = workbook.active

# Añadir encabezados
sheet.append(["Numero", "Semanas", "Cancion","Movimiento"])


# Verificar si se encontró la tabla
if tabla:
    # Iterar sobre las filas de la tabla
    for fila in tabla.find_all('tr'):
        # Obtener las celdas de la fila
        celdas = fila.find_all(['th', 'td'])

        # Imprimir el contenido de cada celda
        for celda in celdas:
            print(celda.text.strip())  # Puedes ajustar esto según tus necesidades
            numero = celdas[0].text.strip()
            semanas= celdas[1].text.strip()
            cancion=  celdas[2].text.strip()
            movimiento = celdas[3].text.strip()
            sheet.append([numero,semanas,cancion,movimiento])

else:
    print('No se encontró ninguna tabla en la página.')

workbook.save("mediatraffic.xlsx")    