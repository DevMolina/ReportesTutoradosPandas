import pandas as pd

# Cargar el archivo Excel
data = pd.read_excel('REPORTE.xlsx')

# Eliminar filas y columnas vacías
data.dropna(axis=0, how='all', inplace=True)
data.dropna(axis=1, how='all', inplace=True)

# Filtrar filas que contienen el código, nombre y situación académica
info_alumnos = data[data.iloc[:, 0].astype(str).str.isdigit() & (data.iloc[:, 0].astype(str).str.len() == 8)]

# Filtrar filas que contienen las notas de las asignaturas
notas_asignaturas = data[data.iloc[:, 1].astype(str).str.isdigit()]
# print(notas_asignaturas)
# Reiniciar los índices para cada DataFrame
info_alumnos.reset_index(drop=True, inplace=True)
notas_asignaturas.reset_index(drop=True, inplace=True)

# Establecer los nombres de las columnas para info_alumnos
info_alumnos.columns = ['Código', 'Nombre', 'Situación Académica', 'N.', 'COD', 'GRP', 'NOMBRE MATERIA', 'FALLAS', 'P20', 'S20']
print(info_alumnos)
# Eliminar columnas vacías de notas_asignaturas
notas_asignaturas.dropna(axis=1, how='all', inplace=True)

# Establecer los nombres de las columnas para notas_asignaturas
notas_asignaturas.columns = notas_asignaturas.iloc[0]
notas_asignaturas = notas_asignaturas[1:]

# Reiniciar los índices para notas_asignaturas
notas_asignaturas.reset_index(drop=True, inplace=True)

# Unir info_alumnos con notas_asignaturas
alumnos_con_notas = pd.concat([info_alumnos, notas_asignaturas], axis=1)

# Eliminar filas con NaN en la columna 'Código'
alumnos_con_notas = alumnos_con_notas.dropna(subset=['Código'])

# print(alumnos_con_notas)