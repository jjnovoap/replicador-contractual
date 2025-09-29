# =============================================================================
# MÓDULO MAIN: Orquestación del Flujo de Trabajo (Generación y Envío)
# =============================================================================

# -----------------------------------------------------------------------------
# LIBRERÍAS ESTÁNDAR
# -----------------------------------------------------------------------------
import os
import uuid  # Generar un ID único para cada lote de emails.
from collections import defaultdict

# -----------------------------------------------------------------------------
# LIBRERÍAS DE TERCEROS
# -----------------------------------------------------------------------------
from dotenv import load_dotenv

# -----------------------------------------------------------------------------
# MÓDULOS LOCALES
# -----------------------------------------------------------------------------
from utils import (
    cargar_plantilla,
    cargar_datos,
    procesar_solicitud,
    guardar_solicitud,
    generar_word_obligaciones,
    ajustar_tamano_celdas,
    excel_a_pdf,
    word_a_pdf,
    enviar_email_lote,
)

# -----------------------------------------------------------------------------
# FUNCIÓN PRINCIPAL DE EJECUCIÓN
# -----------------------------------------------------------------------------

def main():
    # -------------------------------------------------------------------------
    ## 🚀 Configuración Inicial
    # -------------------------------------------------------------------------
    load_dotenv()
    # Diccionario para agrupar todas las órdenes por el correo del supervisor
    ordenes_por_supervisor = defaultdict(list)

    # Variables de entorno
    TAMANO_LOTE_EMAIL = int(os.environ.get("EMAIL_BATCH_SIZE", 10))
    REMITENTE_EMAIL = os.environ.get("EMAIL_REMITENTE")
    REMITENTE_PASSWORD = os.environ.get("REMITENTE_PASSWORD")
    CORREO_COPIA = os.environ.get("CORREO_COPIA")

    # Definición de Rutas del Sistema
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    RUTA_DATOS = os.path.join(BASE_DIR, "datos", "2. BaseDatos -enviar.xlsx")
    RUTA_PLANTILLA = os.path.join(BASE_DIR, "datos", "plantillas", "SolicitudOrden.xlsx")
    RUTA_OUTPUT = os.path.join(BASE_DIR, "output")
    RUTA_PLANTILLA_WORD = os.path.join(BASE_DIR, "datos", "plantillas", "Anexo.docx")

    os.makedirs(RUTA_OUTPUT, exist_ok=True)

    # -------------------------------------------------------------------------
    ## 📚 Carga de Datos
    # -------------------------------------------------------------------------
    wb_datos = cargar_datos(RUTA_DATOS)
    hoja_datos = wb_datos.active
    headers = [cell.value for cell in hoja_datos[1]]

    print("\n--- INICIO DE GENERACIÓN DE ARCHIVOS ---")

    # -------------------------------------------------------------------------
    ## --- FASE 1: Generación de Archivos ---
    # -------------------------------------------------------------------------
    for row in hoja_datos.iter_rows(min_row=2, values_only=True):
        # Sanitizar fila: convertir None a string vacío y limpiar espacios
        datos_fila_raw = dict(zip(headers, row))
        datos_fila = {}
        for k, v in datos_fila_raw.items():
            if v is None:
                datos_fila[k] = ""
            elif isinstance(v, str):
                datos_fila[k] = v.strip()
            else:
                datos_fila[k] = str(v)

        # 1. Procesar plantilla Excel
        wb_plantilla = cargar_plantilla(RUTA_PLANTILLA)
        wb_procesado = procesar_solicitud(wb_plantilla, datos_fila)
        hoja_procesada = wb_procesado.active

        # 2. Ajustar celdas dinámicamente
        celdas_a_ajustar = [
            (21, 2), (24, 2), (27, 2),
            (45, 3), (46, 3), (47, 3), (48, 3), (49, 3),
        ]
        ajustar_tamano_celdas(hoja_procesada, celdas_a_ajustar)

        # 3. Guardar Solicitud en formato XLSX
        nombre_archivo = f"SolicitudOrden CE-{datos_fila.get('radi', '')}-2025.xlsx"
        ruta_excel = guardar_solicitud(wb_procesado, nombre_archivo, RUTA_OUTPUT)

        # 4. Convertir Excel a PDF (solo la primera página)
        excel_a_pdf(
            ruta_excel, 
            ruta_excel.replace(".xlsx", ".pdf"), 
            paginas_to_keep=[0]
        )

        # 5. Generar Word (Anexo de Obligaciones) si aplica
        nombre_archivo_word = f"Anexo CE-{datos_fila.get('radi', '')}-2025 {datos_fila.get('nombre_contratista', '')}.docx"
        ruta_word = generar_word_obligaciones(datos_fila, RUTA_OUTPUT, nombre_archivo_word, RUTA_PLANTILLA_WORD)

        if ruta_word:
            # Convierte el archivo DOCX generado a PDF
            word_a_pdf(ruta_word, ruta_word.replace(".docx", ".pdf"))

        print(f"✅ Documentos generados para: {nombre_archivo}")

        # 6. Agrupar por supervisor para el envío por lotes
        supervisor = datos_fila.get("correo_director_proyecto")
        if supervisor:
            ordenes_por_supervisor[supervisor].append(
                (nombre_archivo, ruta_excel, ruta_word, datos_fila)
            )

    print("\n--- GENERACIÓN DE ARCHIVOS COMPLETADA ---")

    # -------------------------------------------------------------------------
    ## --- FASE 2: Envío de Emails por Supervisor ---
    # -------------------------------------------------------------------------
    if not REMITENTE_EMAIL or not REMITENTE_PASSWORD:
        print("❌ ERROR CRÍTICO: Credenciales de correo no configuradas.")
        return

    if not ordenes_por_supervisor:
        print("⚠️ No se encontraron órdenes para enviar.")
        return

    for supervisor, ordenes in ordenes_por_supervisor.items():
        # Construye la lista de destinatarios para este supervisor
        destinatarios = [supervisor]
        if CORREO_COPIA:
            destinatarios.append(CORREO_COPIA)

        print(f"\n📧 Preparando envío al supervisor: {supervisor} ({len(ordenes)} órdenes)")

        # Procesa las órdenes en lotes definidos por TAMANO_LOTE_EMAIL
        for i in range(0, len(ordenes), TAMANO_LOTE_EMAIL):
            lote = ordenes[i : i + TAMANO_LOTE_EMAIL]
            
            # Genera un ID ÚNICO (UUID corto) para este lote de correos
            id_batch_unico = uuid.uuid4().hex[:6]

            # Llamada a la función de envío de correo electrónico.
            enviar_email_lote(
                lote,
                REMITENTE_EMAIL,
                REMITENTE_PASSWORD,
                destinatarios,
                id_batch_unico,
            )

    print("\n--- Proceso completo de generación y envío ---")

if __name__ == "__main__":
    # La ejecución comienza aquí.
    main()