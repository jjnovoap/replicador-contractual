import os

from config import CAMPOS
from utils import cargar_plantilla, cargar_datos, procesar_solicitud, guardar_solicitud, generar_word_obligaciones, ajustar_tamano_celdas, excel_a_pdf, word_a_pdf

def main():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    RUTA_DATOS = os.path.join(BASE_DIR, 'datos', '2. BaseDatos -enviar.xlsx')
    RUTA_PLANTILLA = os.path.join(BASE_DIR, 'datos', 'plantillas', 'SolicitudOrden.xlsx')
    RUTA_OUTPUT = os.path.join(BASE_DIR, 'output')
    os.makedirs(RUTA_OUTPUT, exist_ok=True)
    
    wb_datos = cargar_datos(RUTA_DATOS)
    hoja_datos = wb_datos.active
    headers = [cell.value for cell in hoja_datos[1]]
    
    for row in hoja_datos.iter_rows(min_row=2, values_only=True):
        datos_fila = dict(zip(headers, row))
        
        wb_plantilla = cargar_plantilla(RUTA_PLANTILLA)
        wb_procesado = procesar_solicitud(wb_plantilla, datos_fila)
        hoja_procesada = wb_procesado.active
        
        # Ajustar filas dinámicamente
        hoja_procesada = wb_procesado.active
        
        # 2. Define las celdas que necesitas ajustar
        # (Este es un ejemplo, debes cambiar 'A1', 'B2', etc. por las celdas reales que contienen texto largo)
        # Por ejemplo, si los datos se escriben en la columna C y en la fila 5, la celda es (5, 3).
        celdas_a_ajustar = [
            (21, 2),  
            (24, 2),
            #(27, 2),   
            (45, 3),  
            (46, 3),  
            (47, 3),
            (48, 3),   
            (49, 3)   
            # Agrega aquí todas las celdas que necesites ajustar
        ]
        
        # 3. Llama a la función con los parámetros correctos
        ajustar_tamano_celdas(
            hoja_procesada,
            celdas_a_ajustar
        )

        nombre_archivo = f"SolicitudOrden CE-{datos_fila.get('radi', '')}-2025.xlsx"
        ruta_excel = guardar_solicitud(wb_procesado, nombre_archivo, RUTA_OUTPUT)

        # Exportar Excel a PDF 
        excel_a_pdf(
            ruta_excel,
            ruta_excel.replace(".xlsx", ".pdf"),
            paginas_to_keep=[0]   # ✅ convierte solo la primera página
        )

        ruta_plantilla_word = os.path.join(BASE_DIR, "datos", "plantillas", "Anexo.docx")
        nombre_archivo_word = f"Anexo CE-{datos_fila.get('radi', '')}-2025 {datos_fila.get('nombre_contratista', '')}.docx"
        ruta_word = generar_word_obligaciones(datos_fila, RUTA_OUTPUT, nombre_archivo_word, ruta_plantilla_word)

        if ruta_word:
            word_a_pdf(ruta_word, ruta_word.replace(".docx", ".pdf"))

        print(f"Solicitud generada: {nombre_archivo}")

if __name__ == "__main__":
    main()