import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Importar el archivo Excel
data = pd.read_excel('REPORTE.xlsx', header=None)



# Eliminar filas y columnas vacías
data.dropna(axis=0, how='all', inplace=True)
data.dropna(axis=1, how='all', inplace=True)
# print(data)
# Encontrar los índices de inicio de cada estudiante (donde el código está presente)
start_indices = data[data.iloc[:, 0].astype(str).str.isdigit() & (data.iloc[:, 0].astype(str).str.len() > 3)].index

# Crear una lista para almacenar los datos de los estudiantes y las notas de las asignaturas

# Iterar sobre los índices de inicio de cada estudiante
for start_index in start_indices:
    # Extraer información del estudiante
    codigo = data.iloc[start_index, 0]
    
    nombre = data.iloc[start_index, 1].split(':')[1].strip()
    situacion = data.iloc[start_index + 1, 4]
    promedio_acumulado = data.iloc[start_index + 1, 9].split(':')[1].strip()
    # print(codigo, nombre, situacion)
    # temp_df['Situación Académica'] = situacion
    
    # df_estudiantes = df_estudiantes._append({'Codigo': codigo, 'Nombre': nombre, 'Situacion': situacion}, ignore_index=True)
    # Agregar el DataFrame temporal a la lista de estudiantes
    # estudiantes.append(temp_df)
    # Agregar la información del estudiante al DataFrame original 'data'
    data.loc[start_index:start_index + 50, 'Codigo'] = codigo
    data.loc[start_index:start_index + 50, 'Nombre'] = nombre
    data.loc[start_index:start_index + 50, 'Situacion'] = situacion
    data.loc[start_index:start_index + 50, 'PromedioAcumulado'] = promedio_acumulado


# print(data.head(45))
# Concatenar todos los DataFrames de estudiantes en uno solo
# df = pd.concat(estudiantes)

# Restablecer los índices


data.columns = ['N.', 'CódigoAsignatura', 'Grupo', 'Asignatura', 'BORRAR', 'BORRAR', 'FALLAS', 'P20', 'S20', 'BORRAR',  'CódigoEstudiantil', 'Nombre', 'SituaciónAcadémica', 'PromedioAcumulado']
data = data.drop(columns='BORRAR')
data = data.dropna(subset=['CódigoAsignatura'])
data = data.dropna(subset=['P20'])
data.reset_index(drop=True, inplace=True)


# Cambio de tipo de dato colomnas
data['P20'] = data['P20'].astype(float)
data['S20'] = data['S20'].astype(float)

# Promedio

data['Promedio'] = (data['P20'] + data['S20'])/2


# Asignaturas perdidas

# Crear un nuevo DataFrame con los estudiantes que tienen un promedio por debajo de 3.0
df_promedio_bajo = data[data['Promedio'] < 3.0].copy()

data_organizada = df_promedio_bajo.sort_values(by=['Asignatura', 'Promedio'])

conteo_por_asignatura = df_promedio_bajo.groupby('Asignatura').size().reset_index(name='Conteo').sort_values(by=['Conteo'],ascending=False)
conteo_por_asignatura.reset_index(drop=True, inplace=True)
print(conteo_por_asignatura)




# print(data_organizada.head(45))

# Agrupar por estudiante y calcular el promedio de los promedios por asignatura
promedio_estudiante = data.groupby('Nombre')['Promedio'].mean().reset_index().sort_values(by=['Promedio'])

# Crear un nuevo conjunto de datos con el promedio de cada estudiante
# nuevo_dataset = pd.merge(promedio_estudiante, data[['Estudiante', 'OtrasColumnas']], on='Estudiante', how='left')

# Mostrar el nuevo conjunto de datos
print(promedio_estudiante)

# Definir los rangos y las categorías
rangos = [0, 3, 3.5, 4, 5.1]
categorias = ['No alcanza resultados de aprendizaje', 'Nivel Básico', 'Nivel intermedio', 'Nivel avanzado']

# Crear una nueva columna 'Categoria' que indique la categoría según el promedio
data['Categoria'] = pd.cut(data['Promedio'], bins=rangos, labels=categorias, right=False)
data_categorias = data.sort_values(by=['Categoria', 'Promedio']).copy()
conteo_por_categoria = data_categorias.groupby('Categoria').size().reset_index(name='Conteo').sort_values(by=['Conteo'])
conteo_por_categoria.reset_index(drop=True, inplace=True)
# Mostrar el DataFrame con la nueva columna
# print(conteo_por_categoria)

print(data.head(42))
#Tabla de tutorados
data_resumen = data.groupby('Nombre').agg({
    'CódigoEstudiantil': 'first',
    'SituaciónAcadémica': 'first',
    'PromedioAcumulado': 'first',
    'Promedio': 'mean'
}).reset_index()
columnas_resumen = ['CódigoEstudiantil', 'Nombre', 'SituaciónAcadémica', 'PromedioAcumulado', 'Promedio']
data_resumen = data_resumen[columnas_resumen]
data_resumen = data_resumen.sort_values(by=['Promedio']).copy().reset_index()

# Mostrar el DataFrame con la nueva columna
print(data_resumen.head(42))

ruta_archivo_excel = 'data_resumen.xlsx'

# Exporta el DataFrame a un archivo de Excel
data_resumen.to_excel(ruta_archivo_excel, index=False)

print("DataFrame exportado a Excel exitosamente.")



# Crear el gráfico de frecuencia usando Seaborn Estudiantes que Perdieron Asignaturas
plt.figure(figsize=(10, 6))
grafico = sns.barplot(x='Conteo', y='Asignatura', data=conteo_por_asignatura, palette='viridis')
# Ajustar el espacio para las etiquetas del eje Y
plt.subplots_adjust(left=0.3, right=0.9, top=0.9, bottom=0.1)

# Añadir la frecuencia a cada barra
for index, row in conteo_por_asignatura.iterrows():
    grafico.text(row['Conteo']+0.2, index, str(row['Conteo']), color='black', ha="center")

# Añadir etiquetas
plt.xlabel('Frecuencia')
plt.ylabel('Asignatura')
plt.title('Estudiantes que Perdieron Asignaturas')

# Mostrar el gráfico
plt.show()

# Crear el gráfico de Categorias Resultados de aprendizaje
plt.figure(figsize=(10, 6))
grafico = sns.barplot(x='Conteo', y='Categoria', data=conteo_por_categoria, palette='viridis')
# Ajustar el espacio para las etiquetas del eje Y
plt.subplots_adjust(left=0.3, right=0.9, top=0.9, bottom=0.1)

# Añadir la frecuencia a cada barra
for index, row in conteo_por_categoria.iterrows():
    grafico.text(row['Conteo']-10, index, str(row['Conteo']), color='white', ha="center")

# Añadir etiquetas
plt.xlabel('Frecuencia')
plt.ylabel('Categoria')
plt.title('Resultados de Aprendizaje')

# Mostrar el gráfico
plt.show()

MatematicaBasica = data.loc[data['Asignatura'] == 'MATEMÁTICA BÁSICA']
MatematicaBasica = MatematicaBasica.groupby('Grupo').agg({
    'CódigoAsignatura': 'first',
    'Promedio': 'mean'
}).reset_index()
MatematicaBasica = MatematicaBasica.sort_values(by=['Promedio'])
MatematicaBasica.reset_index(drop=True, inplace=True)
print(MatematicaBasica.head(50))
MatematicaBasica.to_excel('MatematicaBasica.xlsx', index=False)

MatematicaAvanzada= data.loc[data['Asignatura'] == 'MATEMÁTICA AVANZADA']
MatematicaAvanzada = MatematicaAvanzada.groupby('Grupo').agg({
    'CódigoAsignatura': 'first',
    'Promedio': 'mean'
}).reset_index()
MatematicaAvanzada = MatematicaAvanzada.sort_values(by=['Promedio'])
MatematicaAvanzada.reset_index(drop=True, inplace=True)
print(MatematicaAvanzada.head(50))
MatematicaAvanzada.to_excel('MatematicaAvanzada.xlsx', index=False)

Logica= data.loc[data['Asignatura'] == 'LÓGICA']
Logica = Logica.groupby('Grupo').agg({
    'CódigoAsignatura': 'first',
    'Promedio': 'mean'
}).reset_index()
Logica = Logica.sort_values(by=['Promedio'])
Logica.reset_index(drop=True, inplace=True)
print(Logica.head(50))
Logica.to_excel('Logica.xlsx', index=False)