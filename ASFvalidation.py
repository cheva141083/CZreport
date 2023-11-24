import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Border, Side






def contar_archivos_en_directorio(folder):
    try:
        # Lista de elementos en el directorio
        elementos = os.listdir(folder)

        # Contador para archivos
        contador_archivos = 0

        # Itera sobre los elementos y cuenta los archivos
        for elemento in elementos:
            ruta_completa = os.path.join(folder, elemento)
            if os.path.isfile(ruta_completa):
                contador_archivos += 1

        return contador_archivos

    except FileNotFoundError:
        print(f'El directorio {folder} no existe.')
    except Exception as e:
        print(f'Ocurrió un error al contar los archivos: {e}')

# Search a file in a specific directory
folderUbication ='E:/Resources/NAM US/Python/QA ASF'

# Obtiene y muestra la cantidad de archivos en el directorio
cantidad_archivos = contar_archivos_en_directorio(folderUbication)
print(f'La cantidad de archivos en el directorio {folderUbication} es: {cantidad_archivos}')