import pandas as pd
import os
import re

# Función para limpiar caracteres no válidos
def limpiar_dato(dato):
    if isinstance(dato, str):
        # Eliminar caracteres de control no imprimibles (como \x00)
        dato = re.sub(r'[\x00-\x1F\x7F]', '', dato)
        # Eliminar espacios adicionales al inicio y al final
        dato = dato.strip()
    return dato

# Función para procesar cada archivo .asc
def procesar_archivo_asc(ruta_archivo):
    try:
        # Intentar abrir el archivo con UTF-8
        with open(ruta_archivo, 'r', encoding='utf-8') as file:
            lineas = file.readlines()
    except UnicodeDecodeError:
        # Si falla, intentar con latin-1 (ISO-8859-1)
        with open(ruta_archivo, 'r', encoding='latin-1') as file:
            lineas = file.readlines()
    
    # Extraer los títulos de las columnas
    titulos = [limpiar_dato(col) for col in lineas[0].strip().split('|')]
    num_columnas = len(titulos)  # Número esperado de columnas
    
    # Procesar las líneas de datos
    datos = []
    for linea in lineas[1:]:
        # Reemplazar los saltos de línea con '-|-' y luego dividir por '|'
        linea_procesada = [limpiar_dato(col) for col in linea.strip().replace('\n', '-|-').split('|')]
        
        # Ajustar el número de columnas si es necesario
        if len(linea_procesada) > num_columnas:
            # Si hay más columnas de las esperadas, truncar
            linea_procesada = linea_procesada[:num_columnas]
        elif len(linea_procesada) < num_columnas:
            # Si hay menos columnas, rellenar con valores vacíos
            linea_procesada.extend([''] * (num_columnas - len(linea_procesada)))
        
        datos.append(linea_procesada)
    
    # Crear un DataFrame con los datos
    df = pd.DataFrame(datos, columns=titulos)
    return df

# Directorio donde se encuentran los archivos .asc
directorio = '.'

# Crear una subcarpeta para guardar los archivos Excel
directorio_salida = os.path.join(directorio, 'excel_files')
os.makedirs(directorio_salida, exist_ok=True)  # Crear la carpeta si no existe

# Procesar cada archivo .asc en el directorio
for archivo in os.listdir(directorio):
    if archivo.endswith('.asc'):
        ruta_completa = os.path.join(directorio, archivo)
        
        # Procesar el archivo .asc
        df = procesar_archivo_asc(ruta_completa)
        
        # Crear el nombre del archivo Excel (mismo nombre que el .asc, pero con extensión .xlsx)
        nombre_excel = os.path.splitext(archivo)[0] + '.xlsx'
        ruta_excel = os.path.join(directorio_salida, nombre_excel)
        
        # Guardar el DataFrame en un archivo Excel
        df.to_excel(ruta_excel, index=False, engine='openpyxl')
        print(f"Archivo '{archivo}' convertido a '{nombre_excel}'.")

print(f"Proceso completado. Todos los archivos .asc han sido convertidos a .xlsx y guardados en '{directorio_salida}'.")