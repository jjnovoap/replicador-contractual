# =============================================================================
# MÓDULO MAIN: Orquestación del Flujo de Trabajo (Generación y Envío)
# =============================================================================

import os
import uuid
from collections import defaultdict
from dotenv import load_dotenv

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
    sanitize_filename,
)

def main():
    load_dotenv()
    ordenes_por_supervisor = defaultdict(list)

    TAMANO_LOTE_EMAIL = int(os.environ.get("EMAIL_BATCH_SIZE", 10))
    REMITENTE_EMAIL = os.environ.get("EMAIL_REMITENTE")
    REMITENTE_PASSWORD = os.environ.get("REMITENTE_PASSWORD")
    CORREO_COPIA = os.environ.get("CORREO_COPIA")

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    RUTA_DATOS = os.path.join(BASE_DIR, "datos", "2. BaseDatos -enviar.xlsx")
    RUTA_PLANTILLA = os.path.join(BASE_DIR, "datos", "plantillas", "SolicitudOrden.xlsx")
    RUTA_OUTPUT = os.path.join(BASE_DIR, "output")
    RUTA_PLANTILLA_WORD = os.path.join(BASE_DIR, "datos", "plantillas", "Anexo.docx")

    os.makedirs(RUTA_OUTPUT, exist_ok=True)

    wb_datos = cargar_datos(RUTA_DATOS)
    hoja_datos = wb_datos.active
    headers = [cell.value for cell in hoja_datos[1]]

    print("\n--- INICIO DE GENERACIÓN DE ARCHIVOS ---")

    for fila_idx, row in enumerate(hoja_datos.iter_rows(min_row=2, values_only=True), start=1):
        datos_fila_raw = dict(zip(headers, row))
        datos_fila = {}
        for k, v in datos_fila_raw.items():
            if v is None:
                datos_fila[k] = ""
            elif isinstance(v, str):
                datos_fila[k] = v.strip()
            else:
                datos_fila[k] = str(v)

        wb_plantilla = cargar_plantilla(RUTA_PLANTILLA)
        wb_procesado = procesar_solicitud(wb_plantilla, datos_fila)
        hoja_procesada = wb_procesado.active

        celdas_a_ajustar = [
            (21, 2), (24, 2), (27, 2),
            (45, 3), (46, 3), (47, 3), (48, 3), (49, 3),
        ]
        ajustar_tamano_celdas(hoja_procesada, celdas_a_ajustar)

        radi_val = datos_fila.get('radi') or f"no-radi-{fila_idx}"
        radi_val = sanitize_filename(radi_val)
        contractor_safe = sanitize_filename(datos_fila.get('nombre_contratista', ''))
        # Añadimos fila_idx al nombre para garantizar unicidad incluso si radi falta/duplicado
        nombre_archivo = f"SolicitudOrden_CE-{radi_val}-2025.xlsx"
        ruta_excel = guardar_solicitud(wb_procesado, nombre_archivo, RUTA_OUTPUT)

        excel_a_pdf(
            ruta_excel,
            ruta_excel.replace(".xlsx", ".pdf"),
            paginas_to_keep=[0]
        )

        nombre_archivo_word = f"Anexo_CE-{radi_val}-2025_{contractor_safe}.docx"
        ruta_word = generar_word_obligaciones(datos_fila, RUTA_OUTPUT, nombre_archivo_word, RUTA_PLANTILLA_WORD)

        if ruta_word:
            word_a_pdf(ruta_word, ruta_word.replace(".docx", ".pdf"))

        print(f"✅ Documentos generados para: {nombre_archivo}")


        supervisor = datos_fila.get("correo_director_proyecto")
        if supervisor:
            ordenes_por_supervisor[supervisor].append(
                (nombre_archivo, ruta_excel, ruta_word, datos_fila)
            )

    print("\n--- GENERACIÓN DE ARCHIVOS COMPLETADA ---")

    if not REMITENTE_EMAIL or not REMITENTE_PASSWORD:
        print("❌ ERROR CRÍTICO: Credenciales de correo no configuradas.")
        return

    if not ordenes_por_supervisor:
        print("⚠️ No se encontraron órdenes para enviar.")
        return

    for supervisor, ordenes in ordenes_por_supervisor.items():
        destinatarios = [supervisor]
        if CORREO_COPIA:
            destinatarios.append(CORREO_COPIA)
        print(f"\n📧 Preparando envío al supervisor: {supervisor} ({len(ordenes)} órdenes)")
        for i in range(0, len(ordenes), TAMANO_LOTE_EMAIL):
            lote = ordenes[i : i + TAMANO_LOTE_EMAIL]
            id_batch_unico = uuid.uuid4().hex[:6]
            enviar_email_lote(
                lote,
                REMITENTE_EMAIL,
                REMITENTE_PASSWORD,
                destinatarios,
                id_batch_unico,
            )

    print("\n--- Proceso completo de generación y envío ---")

if __name__ == "__main__":
    main()