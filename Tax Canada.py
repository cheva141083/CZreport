import itertools

# Lista de números
numeros = [
    135548.71,1350847.34,753.15,2050.78,280.99,67747.81,6388.49,5890.85,509.74,8928.34,80.36,29006.21,3448.10,43103.74,24906.16,1631.34,81592.06,154039.61,303.32,34136.71,272868.41
    ]


# Suma objetivo
suma_objetivo = 155843.32

# Variable para almacenar las combinaciones encontradas
combinaciones_encontradas = []

# Probar con combinaciones de diferentes tamaños
for r in range(2, len(numeros) + 1):  # Desde combinaciones de 2 hasta el tamaño total de la lista
    for comb in itertools.combinations(numeros, r):
        if sum(comb) == suma_objetivo:
            combinaciones_encontradas.append(comb)

# Imprimir las combinaciones encontradas
if combinaciones_encontradas:
    print("Combinaciones que suman", suma_objetivo, ":")
    for comb in combinaciones_encontradas:
        print(comb)
else:
    print("No se encontraron combinaciones que sumen", suma_objetivo)