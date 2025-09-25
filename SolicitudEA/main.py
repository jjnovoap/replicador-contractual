import os
from openpyxl import load_workbook
from utils import cargar_plantilla, cargar_datos, procesar_solicitud, guardar_solicitud
from config import CAMPOS

def main():
    # Rutas de archivos
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    RUTA_DATOS = os.path.join(BASE_DIR, 'datos', 'BaseDatos.xlsx')
    RUTA_PLANTILLA = os.path.join(BASE_DIR, 'datos', 'plantillas', 'SolicitudOrden.xlsx')
    RUTA_OUTPUT = os.path.join(BASE_DIR, 'output')
    
    # Crear directorio output si no existe
    os.makedirs(RUTA_OUTPUT, exist_ok=True)
    
    # Cargar datos
    wb_datos = cargar_datos(RUTA_DATOS)
    hoja_datos = wb_datos.active
    
    # Obtener nombres de columnas (asumimos que la primera fila tiene los headers)
    headers = [cell.value for cell in hoja_datos[1]]
    
    # Procesar cada fila de datos
    for row in hoja_datos.iter_rows(min_row=2, values_only=True):
        datos_fila = dict(zip(headers, row))
        
        # Cargar plantilla limpia para cada solicitud
        wb_plantilla = cargar_plantilla(RUTA_PLANTILLA)
        
        # Procesar solicitud
        wb_procesado = procesar_solicitud(wb_plantilla, datos_fila)
        
        # Generar nombre de archivo de salida
        nombre_archivo = f"SolicitudOrden CE-{datos_fila.get('radi', '')}-2025.xlsx"
        
        # Guardar solicitud generada
        guardar_solicitud(wb_procesado, nombre_archivo, RUTA_OUTPUT)
        
        print(f"Solicitud generada: {nombre_archivo}")

if __name__ == "__main__":
    main()