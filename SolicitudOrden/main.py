# =============================================================================
# MÓDULO MAIN: Orquestación del Flujo de Trabajo (Generación y Envío)
# =============================================================================

# -----------------------------------------------------------------------------
# LIBRERÍAS ESTÁNDAR
# -----------------------------------------------------------------------------
import os
import uuid # Necesario para generar un ID único para cada lote de emails.

# -----------------------------------------------------------------------------
# LIBRERÍAS DE TERCEROS
# -----------------------------------------------------------------------------
from dotenv import load_dotenv

# -----------------------------------------------------------------------------
# MÓDULOS LOCALES
# -----------------------------------------------------------------------------
from config import CAMPOS
# Importación de todas las funciones de utilidad.
from utils import cargar_plantilla, cargar_datos, procesar_solicitud, guardar_solicitud, generar_word_obligaciones, ajustar_tamano_celdas, excel_a_pdf, word_a_pdf, enviar_email_lote

# -----------------------------------------------------------------------------
# FUNCIÓN PRINCIPAL DE EJECUCIÓN
# -----------------------------------------------------------------------------

def main():
    
    # -------------------------------------------------------------------------
    ## 🚀 Configuración Inicial y Carga de Variables
    # -------------------------------------------------------------------------
    
    # 1. Carga de Variables de Entorno desde el archivo .env
    load_dotenv() 
    
    # 2. Extracción y Conversión de Variables de Configuración
    TAMANO_LOTE_EMAIL = int(os.environ.get("EMAIL_BATCH_SIZE", 10))
    REMITENTE_EMAIL = os.environ.get("EMAIL_REMITENTE")
    REMITENTE_PASSWORD = os.environ.get("REMITENTE_PASSWORD")
    CORREO_COPIA = os.environ.get("CORREO_COPIA") # Correo de copia opcional
    
    # 3. Definición de Rutas del Sistema
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    
    RUTA_DATOS = os.path.join(BASE_DIR, 'datos', '2. BaseDatos -enviar.xlsx')
    RUTA_PLANTILLA = os.path.join(BASE_DIR, 'datos', 'plantillas', 'SolicitudOrden.xlsx')
    RUTA_OUTPUT = os.path.join(BASE_DIR, 'output')
    RUTA_PLANTILLA_WORD = os.path.join(BASE_DIR, "datos", "plantillas", "Anexo.docx")
    
    os.makedirs(RUTA_OUTPUT, exist_ok=True) # Crea el directorio 'output' si no existe.

    # -------------------------------------------------------------------------
    ## 📚 Carga de Datos y Extracción del Correo Principal (Supervisor)
    # -------------------------------------------------------------------------
    
    # 4. Cargar la base de datos y extraer los encabezados
    wb_datos = cargar_datos(RUTA_DATOS)
    hoja_datos = wb_datos.active
    headers = [cell.value for cell in hoja_datos[1]] # Encabezados (primera fila)

    # 🚨 PASO CLAVE: EXTRAER EL CORREO PRINCIPAL ANTES DEL BUCLE 🚨
    CORREO_SUPERVISOR_PRINCIPAL = None
    try:
        # Obtener el primer registro de datos (Fila 2) para el correo del supervisor
        fila_principal = next(hoja_datos.iter_rows(min_row=2, max_row=2, values_only=True))
        datos_fila_principal = dict(zip(headers, fila_principal))
        
        # 5. Definir el CORREO_SUPERVISOR usando el dato extraído
        CORREO_SUPERVISOR_PRINCIPAL = datos_fila_principal.get('correo_director_proyecto')
        
    except StopIteration:
        print("❌ ERROR CRÍTICO: El archivo de datos no contiene registros de contratos.")
        return
    except Exception as e:
        print(f"❌ ERROR al intentar leer el correo del director del proyecto: {e}")
        return
        
    # -------------------------------------------------------------------------
    ## ⚙️ Definición de Destinatarios y Validación Final
    # -------------------------------------------------------------------------
    
    # 6. CREAR LA LISTA FINAL DE DESTINATARIOS
    CORREOS_DESTINATARIOS_RAW = [
        CORREO_SUPERVISOR_PRINCIPAL, # Principal, extraído del Excel
        CORREO_COPIA                 # Copia, extraído de .env
    ]

    # Filtra la lista (eliminando None o cadenas vacías)
    DESTINATARIOS_LISTA = [c.strip() for c in CORREOS_DESTINATARIOS_RAW if c and c.strip()]
    
    # 7. Validación Crítica Final
    if not REMITENTE_EMAIL or not REMITENTE_PASSWORD:
        print("❌ ERROR CRÍTICO: Las variables de entorno para las credenciales de email (EMAIL_REMITENTE, REMITENTE_PASSWORD) no están definidas.")
        return # Termina el programa de forma segura si faltan credenciales.

    if not DESTINATARIOS_LISTA:
        print("❌ ERROR CRÍTICO: No se pudo determinar un destinatario válido (correo_director_proyecto o CORREO_COPIA).")
        return

    # Lista de acumulación para la Fase 2 (Envío).
    ordenes_generadas = [] 
    
    # -------------------------------------------------------------------------
    ## --- FASE 1: Generación y Exportación de Archivos (Bucle Principal) ---
    # -------------------------------------------------------------------------
    print("\n--- INICIO DE GENERACIÓN DE ARCHIVOS ---")
    
    # Itera sobre cada registro (fila) de la base de datos
    for row in hoja_datos.iter_rows(min_row=2, values_only=True):
        datos_fila = dict(zip(headers, row)) # Mapear los datos de la fila a un diccionario (¡Ahora dentro del bucle!)
        
        # 1. Procesamiento de la Plantilla de Solicitud (Excel)
        wb_plantilla = cargar_plantilla(RUTA_PLANTILLA)
        wb_procesado = procesar_solicitud(wb_plantilla, datos_fila)
        hoja_procesada = wb_procesado.active
        
        # 2. Ajuste Dinámico de Celdas (Wrap Text)
        celdas_a_ajustar = [
            (21, 2), (24, 2), (27, 2), # Columna B
            (45, 3), (46, 3), (47, 3), (48, 3), (49, 3) # Columna C
        ]
        
        ajustar_tamano_celdas(
            hoja_procesada,
            celdas_a_ajustar
        )

        # 3. Guardar Solicitud en formato XLSX
        nombre_archivo = f"SolicitudOrden CE-{datos_fila.get('radi', '')}-2025.xlsx"
        ruta_excel = guardar_solicitud(wb_procesado, nombre_archivo, RUTA_OUTPUT)

        # 4. Conversión de Solicitud de Excel a PDF (solo la primera página)
        excel_a_pdf(
            ruta_excel,
            ruta_excel.replace(".xlsx", ".pdf"),
            paginas_to_keep=[0] 
        )

        # 5. Generar y Exportar Anexo de Obligaciones (Word y PDF)
        nombre_archivo_word = f"Anexo CE-{datos_fila.get('radi', '')}-2025 {datos_fila.get('nombre_contratista', '')}.docx"
        ruta_word = generar_word_obligaciones(datos_fila, RUTA_OUTPUT, nombre_archivo_word, RUTA_PLANTILLA_WORD)

        if ruta_word:
            # 5.1 Convierte el archivo DOCX generado a PDF
            word_a_pdf(ruta_word, ruta_word.replace(".docx", ".pdf"))

        print(f"✅ Documentos generados para: {nombre_archivo}")

        # 6. Almacenar metadatos para el envío de email
        ordenes_generadas.append(
            (nombre_archivo, ruta_excel, ruta_word, datos_fila)
        )
        # ⚠️ FIN DEL BUCLE DE GENERACIÓN POR REGISTRO

    print("\n--- GENERACIÓN DE ARCHIVOS COMPLETADA ---")

    # -------------------------------------------------------------------------
    ## --- FASE 2: Envío de Emails por Lote ---
    # -------------------------------------------------------------------------
    
    # Itera sobre la lista acumulada de órdenes, tomando bloques definidos por TAMANO_LOTE_EMAIL.
    for i in range(0, len(ordenes_generadas), TAMANO_LOTE_EMAIL):
        lote = ordenes_generadas[i:i + TAMANO_LOTE_EMAIL]
        
        # Generar un ID ÚNICO (UUID corto) para este lote de correos
        id_batch_unico = uuid.uuid4().hex[:6]

        # Llamada a la función de envío de correo electrónico.
        enviar_email_lote(
            lote, 
            REMITENTE_EMAIL, 
            REMITENTE_PASSWORD, 
            DESTINATARIOS_LISTA, 
            id_batch_unico 
        )
        
    print("\n--- Proceso completo de generación y envío ---")

if __name__ == "__main__":
    # La ejecución comienza aquí.
    main()