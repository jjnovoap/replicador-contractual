# =============================================================================
# MÓDULO UTILS: Funciones de Procesamiento, Conversión y Comunicación
# =============================================================================

# -----------------------------------------------------------------------------
# LIBRERÍAS ESTÁNDAR
# -----------------------------------------------------------------------------
import io
import math
import os
import subprocess # Usado para ejecutar comandos externos (ej. LibreOffice)
import smtplib # Conexión al servidor de correo (SMTP)
import mimetypes
from email.mime.multipart import MIMEMultipart # Contenedor principal del email
from email.mime.text import MIMEText # Cuerpo de texto (plain)
from email.mime.base import MIMEBase # Contenedor base para adjuntos
from email import encoders # Utilidad para codificar adjuntos (Base64)

# -----------------------------------------------------------------------------
# LIBRERÍAS DE TERCEROS
# -----------------------------------------------------------------------------
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx2pdf import convert # Utilidad para conversión de DOCX a PDF

from PyPDF2 import PdfReader, PdfWriter # Manipulación de PDF (extracción/combinación de páginas)
import language_tool_python # Herramienta para corrección gramatical

# -----------------------------------------------------------------------------
# MÓDULOS LOCALES
# -----------------------------------------------------------------------------
# CAMPOS: Mapeo de columnas de BD a marcadores de plantilla.
# HOJA_SOLICITUD: Nombre de la hoja en la plantilla Excel a procesar.
# OBLIGACIONES: Lista de campos de la BD que contienen texto de obligaciones.
from config import CAMPOS, HOJA_SOLICITUD, OBLIGACIONES

# Inicialización: Se carga el corrector ortográfico y gramatical para español
tool = language_tool_python.LanguageTool('es')

# =============================================================================
# 1. FUNCIONES DE CORECCIÓN Y CARGA
# =============================================================================


def corregir_texto(texto: str) -> str:
    """
    Corrige texto con language_tool_python. Seguro para None/strings vacíos.
    """
    try:
        if texto is None:
            return ""
        texto = str(texto).strip()
        if not texto:
            return ""
        matches = tool.check(texto)
        return language_tool_python.utils.correct(texto, matches)
    except Exception as e:
        # Si el corrector falla, devolvemos el texto original (no None)
        print(f"⚠️ language_tool fallo: {e}")
        return texto


def cargar_plantilla(ruta_plantilla):
    """
    Carga el archivo Excel que sirve como plantilla de la solicitud.
    Retorna: Objeto Workbook de openpyxl.
    """
    return openpyxl.load_workbook(ruta_plantilla)

def cargar_datos(ruta_datos):
    """
    Carga el archivo Excel que contiene los datos fuente (BaseDatos -enviar.xlsx).
    Retorna: Objeto Workbook de openpyxl.
    """
    return openpyxl.load_workbook(ruta_datos)

def buscar_y_reemplazar(hoja, texto_buscar, texto_reemplazar):
    """
    Reemplaza marcadores en las celdas de una hoja. Trivializa None para evitar 'None' textual.
    """
    if texto_reemplazar is None:
        texto_reemplazar = ""
    texto_reemplazar = str(texto_reemplazar)

    for row in hoja.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            if isinstance(cell.value, str) and texto_buscar in cell.value:
                cell.value = cell.value.replace(texto_buscar, texto_reemplazar)

# =============================================================================
# 2. FUNCIONES DE AJUSTE Y PROCESAMIENTO DE EXCEL
# =============================================================================

def ajustar_tamano_celdas(hoja: Worksheet, celdas_a_ajustar: list,
                          ancho_max_columna: int = 197, factor_ajuste: float = 1.05):
    """
    Ajusta la altura de las filas para que el texto contenido en las celdas
    especificadas (con wrapText activado) quepa completamente.

    Parámetros Clave:
    - ancho_max_columna (int): Ancho de columna en unidades Excel.
      (197 es la conversión de 1383px. Mantener constante para el cálculo de líneas).
    - factor_ajuste (float): Factor de seguridad aplicado al cálculo de caracteres por línea.
    """
    filas_procesadas = set()

    for fila, col in celdas_a_ajustar:
        celda = hoja.cell(row=fila, column=col)
        valor_celda = str(celda.value) if celda.value else ""

        # Omite celdas ya procesadas o que están vacías
        if fila in filas_procesadas or not valor_celda:
            continue

        # Configuración obligatoria para que el ajuste de altura funcione
        celda.alignment = Alignment(wrapText=True, horizontal='left', vertical='center')

        # Cálculo de la capacidad de caracteres por línea
        caracteres_por_linea = int(ancho_max_columna * factor_ajuste) 

        if caracteres_por_linea > 0:
            # Calcular el número de líneas requeridas (redondeando hacia arriba)
            num_lineas = math.ceil(len(valor_celda) / caracteres_por_linea)
            # Altura en puntos (14 pt es la altura aproximada para una línea de Calibri 11)
            nueva_altura = num_lineas * 14 

            # Limitar alturas: Mínimo 15 pt (altura de una línea), Máximo 300 pt (límite práctico)
            nueva_altura = max(15, min(nueva_altura, 300))

            hoja.row_dimensions[fila].height = nueva_altura
            filas_procesadas.add(fila)


def procesar_solicitud(wb_plantilla, datos_fila):
    """
    Rellena la plantilla de Excel con los datos de una solicitud y gestiona
    la inserción de las obligaciones.

    Lógica de Obligaciones:
    - Si hay <= 5 obligaciones, se insertan en los marcadores del Excel ({obligacion_1}, etc.).
    - Si hay > 5 obligaciones, se limpian los marcadores del Excel y se indica
      que se generará un Word anexo.
    """
    hoja = wb_plantilla[HOJA_SOLICITUD]

    # 1. Reemplazo de campos generales (excluyendo obligaciones)
    for campo_db, campo_plantilla in CAMPOS.items():
        if campo_db not in OBLIGACIONES:
            valor = datos_fila.get(campo_db, '')
            valor = corregir_texto(str(valor)) # Aplica corrección antes de reemplazar
            buscar_y_reemplazar(hoja, campo_plantilla, valor)
    
    # 2. Gestión de Obligaciones
    obligaciones_reales = [
        corregir_texto(str(datos_fila.get(campo)).strip())
        for campo in OBLIGACIONES
        if datos_fila.get(campo) and str(datos_fila.get(campo)).strip() != ""
    ]

    num_obligaciones = len(obligaciones_reales)

    if num_obligaciones <= 5:
        # Caso A: Pocas obligaciones. Insertar en el Excel.
        print(f"Número de obligaciones: {num_obligaciones}. Agregando al Excel.")
        for i, obligacion in enumerate(obligaciones_reales, start=1):
            marcador_excel = f"{{obligacion_{i}}}"
            buscar_y_reemplazar(hoja, marcador_excel, obligacion)
        # Limpiar los marcadores restantes si hay menos de 5
        for i in range(num_obligaciones + 1, 6):
            marcador_excel = f"{{obligacion_{i}}}"
            buscar_y_reemplazar(hoja, marcador_excel, "")
    else: 
        # Caso B: Muchas obligaciones. Limpiar Excel, generar Word Anexo.
        print(f"Número de obligaciones: {num_obligaciones}. Se generará un documento de Word y se limpiarán los marcadores en Excel.")
        for i in range(1, 6):
            marcador_excel = f"{{obligacion_{i}}}"
            buscar_y_reemplazar(hoja, marcador_excel, " ") # Se reemplaza con espacio para no dejar vacío visualmente

    return wb_plantilla

def guardar_solicitud(wb, nombre_archivo, ruta_output):
    """Guarda el objeto Workbook de Excel procesado en la ruta de salida.
       Antes de guardar limpia celdas None o 'None' para que la planilla no muestre 'None'."""
    # Limpieza preventiva: reemplazar None o 'None' por cadena vacía
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                # ⚠️ Evitar tocar celdas combinadas (read-only)
                if isinstance(cell, MergedCell):
                    continue
                if cell.value is None:
                    cell.value = ""
                elif isinstance(cell.value, str) and cell.value.strip().lower() == "none":
                    cell.value = ""
    ruta_completa = os.path.join(ruta_output, nombre_archivo)
    wb.save(ruta_completa)
    return ruta_completa


# =============================================================================
# 3. FUNCIONES DE PROCESAMIENTO DE WORD
# =============================================================================

def generar_word_obligaciones(datos_fila, ruta_output, nombre_archivo, ruta_plantilla_word):
    """
    Genera el documento de Word 'Anexo' si la solicitud contiene 6 o más
    obligaciones. Rellena campos generales y crea una lista numerada de obligaciones.

    Retorna: La ruta al archivo DOCX generado, o None si no se genera.
    """
    # 1. Filtrar y corregir las obligaciones de la fila
    obligaciones = []
    for campo in OBLIGACIONES:
        valor = datos_fila.get(campo)
        if valor and str(valor).strip() != "":
            obligaciones.append(corregir_texto(str(valor).strip()))

    if len(obligaciones) <= 5:
        # No se cumple la condición de negocio para generar el Word
        print(f"No se genera Word: solo {len(obligaciones)} obligaciones (se necesitan 6 o más).")
        return None

    print(f"Generando documento de Word con {len(obligaciones)} obligaciones.")
    
    try:
        # Cargar la plantilla de Word
        doc = Document(ruta_plantilla_word)
    except Exception as e:
        print(f"❌ Error al cargar la plantilla de Word: {e}")
        return None

     # 2. Remplazo de marcadores generales en párrafos
    for p in doc.paragraphs:
        for campo, valor in datos_fila.items():
            if campo in OBLIGACIONES:
                continue
            marcador_word = f"{{{campo}}}"
            if marcador_word in p.text:
                # corregir_texto maneja None -> ""
                p.text = p.text.replace(marcador_word, corregir_texto(valor))
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(9)

    # 3. Inserción de la lista de obligaciones en el marcador {obligaciones}
    for p in doc.paragraphs:
        if "{obligaciones}" in p.text:
            p.clear() # Limpiar el párrafo marcador
            # Insertar las obligaciones como una lista numerada antes del marcador
            for i, obligacion in enumerate(obligaciones, start=1):
                new_p = p.insert_paragraph_before(f"{i}. {obligacion}")
                # ⚠️ Aplicación de formato específico para la lista
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                new_p.paragraph_format.space_before = Pt(0)
                new_p.paragraph_format.space_after = Pt(0)
                new_p.paragraph_format.first_line_indent = Pt(0)
                for run in new_p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(9)
            # Eliminar el párrafo marcador original
            p._element.getparent().remove(p._element)
            break

    # 4. Guardar el documento final
    nombre_doc = f"{nombre_archivo.replace('.xlsx', '.docx')}" # Asegura extensión .docx
    ruta_doc = os.path.join(ruta_output, nombre_doc)
    doc.save(ruta_doc)
    print(f"Documento Word generado: {ruta_doc}")
    return ruta_doc

# =============================================================================
# 4. FUNCIONES DE CONVERSIÓN DE DOCUMENTOS (PDF)
# =============================================================================

def excel_a_pdf(ruta_excel, ruta_pdf, paginas_to_keep=None, timeout=30, wait_interval=0.5):
    """
    Convierte un Excel a PDF usando LibreOffice (requerimiento externo)
    y luego filtra el PDF resultante para mantener solo las páginas necesarias
    (por defecto, solo la primera página [0] para la solicitud principal).

    Dependencia Crítica: Requiere que LibreOffice (soffice) esté instalado
    y accesible desde la línea de comandos.
    """
    if paginas_to_keep is None:
        paginas_to_keep = [0]  # Por defecto, solo la primera página

    try:
        # 1) Ejecutar comando externo para generar el PDF (LibreOffice)
        print("DEBUG: Iniciando conversión Excel a PDF vía LibreOffice...")
        subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf", "--outdir",
            os.path.dirname(ruta_pdf), ruta_excel
        ], check=True, timeout=timeout) # Uso de timeout para evitar cuelgues

        # 2) Determinar la ruta del PDF temporal generado por LibreOffice
        generated_pdf = os.path.join(
            os.path.dirname(ruta_excel),
            os.path.splitext(os.path.basename(ruta_excel))[0] + ".pdf"
        )

        if not os.path.exists(generated_pdf):
            print(f"❌ No se encontró el PDF generado por LibreOffice en: {generated_pdf}")
            return

        # 3) Leer el PDF generado y filtrar páginas
        with open(generated_pdf, "rb") as f:
            pdf_bytes = f.read()
        reader = PdfReader(io.BytesIO(pdf_bytes))
        total_pages = len(reader.pages)

        # 4) Crear el PDF final con solo las páginas deseadas
        writer = PdfWriter()
        for idx in paginas_to_keep:
            if 0 <= idx < total_pages:
                writer.add_page(reader.pages[idx])
            else:
                print(f"⚠️ Índice fuera de rango al filtrar: {idx} (total={total_pages})")

        if not writer.pages:  # Fallback: si falla el filtrado, añadir la primera página
            print("⚠️ Fallback: No se pudo filtrar. Añadiendo la primera página (índice 0).")
            writer.add_page(reader.pages[0])

        # 5) Guardar el PDF final con la ruta y nombre correctos
        with open(ruta_pdf, "wb") as f_out:
            writer.write(f_out)

        # 6) Limpieza: Borrar el PDF completo que generó LibreOffice
        if os.path.abspath(generated_pdf) != os.path.abspath(ruta_pdf):
            try:
                os.remove(generated_pdf)
            except Exception as e:
                print(f"⚠️ No se pudo borrar el PDF temporal de LibreOffice: {e}")

        print(f"✅ PDF de Solicitud (filtrado) generado: {ruta_pdf}")

    except subprocess.CalledProcessError as cpe:
        print(f"❌ LibreOffice falló al convertir el archivo (Error de Subproceso): {cpe}")
    except subprocess.TimeoutExpired:
        print("❌ La conversión de Excel a PDF superó el tiempo límite (Timeout).")
    except Exception as e:
        print(f"❌ Error genérico en excel_a_pdf: {e}")

def word_a_pdf(ruta_word, ruta_pdf):
    """
    Convierte un DOCX a PDF usando la librería docx2pdf.
    Dependencia Crítica: Requiere que Microsoft Word (Windows) o LibreOffice
    (Linux/macOS) estén instalados.
    """
    try:
        convert(ruta_word, ruta_pdf)
        print(f"✅ PDF de Anexo generado: {ruta_pdf}")
    except Exception as e:
        print(f"❌ Error al convertir Word a PDF. Asegúrate de tener Word o LibreOffice instalado: {e}")

# =============================================================================
# 5. FUNCIONES DE COMUNICACIÓN (EMAIL)
# =============================================================================

def enviar_email_lote(lote_archivos_y_datos: list, remitente_email: str, remitente_password: str, destinatarios_lista: list, ID_BATCH_UNICO: str):
    ASUNTO_BASE = "Solicitud(es) para aval"
    asunto_final = f"{ASUNTO_BASE} Ref: {ID_BATCH_UNICO}"
    num_ordenes = len(lote_archivos_y_datos)
    print(f"\n--- INICIO ENVÍO SMTP (Lote de {num_ordenes} órdenes) ---")

    lista_solicitudes = ""
    archivos_adjuntos = []
    destinatarios_header = ', '.join(destinatarios_lista)

    for i, (_, ruta_excel, ruta_word, datos_fila) in enumerate(lote_archivos_y_datos, start=1):
        radi = datos_fila.get('radi', 'N/A')
        nombre_contratista = datos_fila.get('nombre_contratista', 'Contratista')
        lista_solicitudes += f"{i}. CE - {radi} - 2025 - {nombre_contratista}\n"

        ruta_pdf_excel = ruta_excel.replace(".xlsx", ".pdf")
        if os.path.exists(ruta_pdf_excel):
            archivos_adjuntos.append(os.path.abspath(ruta_pdf_excel))

        ruta_pdf_word = None
        if ruta_word:
            posible_pdf = ruta_word.replace(".docx", ".pdf")
            if os.path.exists(posible_pdf):
                archivos_adjuntos.append(os.path.abspath(posible_pdf))

    # deduplicar manteniendo orden
    archivos_adjuntos = list(dict.fromkeys(archivos_adjuntos))

    cuerpo = f"""Buen día profesor.

De manera atenta remito la(s) siguiente(s) solicitud(es) para su Vo.Bo.

{lista_solicitudes}
Gracias de antemano. 

Cordialmente,

--
Profesional de apoyo 
Cursos de Extensión de Lenguas Extranjeras
Departamento de Lenguas Extranjeras
Facultad de Ciencias Humanas - Sede Bogotá
Universidad Nacional de Colombia
"""

    if not archivos_adjuntos:
        print("❌ Error: No se encontraron archivos PDF válidos para adjuntar. Cancelando envío.")
        return

    # DEBUG: muestra existencia y tamaños
    print("DEBUG: Archivos candidatos (ruta - existe - tamaño bytes):")
    for p in archivos_adjuntos:
        try:
            size = os.path.getsize(p) if os.path.exists(p) else -1
        except Exception:
            size = -1
        print(f" - {p} | exists={os.path.exists(p)} | size={size}")

    server = None
    try:
        print("DEBUG: Intentando conexión a SMTP (smtp.gmail.com:587)...")
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        print("DEBUG: Conexión TLS establecida.")
        server.login(remitente_email, remitente_password)
        print("✅ AUTENTICACIÓN EXITOSA.")

        msg = MIMEMultipart()
        msg['From'] = remitente_email
        msg['To'] = destinatarios_header
        msg['Subject'] = asunto_final
        msg.attach(MIMEText(cuerpo, 'plain'))

        def adjuntar_archivo(msg, ruta_archivo):
            """Adjunta un archivo solo si existe, tiene tamaño y se puede leer."""
            if not ruta_archivo or not os.path.isfile(ruta_archivo):
                print(f"⚠️ No se adjunta (archivo no válido): {ruta_archivo}")
                return False
            tamaño = os.path.getsize(ruta_archivo)
            if tamaño == 0:
                print(f"⚠️ No se adjunta (archivo 0 bytes): {ruta_archivo}")
                return False

            ctype, encoding = mimetypes.guess_type(ruta_archivo)
            if ctype:
                maintype, subtype = ctype.split('/', 1)
            else:
                maintype, subtype = 'application', 'octet-stream'

            with open(ruta_archivo, "rb") as f:
                part = MIMEBase(maintype, subtype)
                part.set_payload(f.read())

            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(ruta_archivo)}"')
            msg.attach(part)
            print(f"✅ Archivo adjuntado: {ruta_archivo}")
            return True

        # Adjuntar
        attached_count = 0
        for ruta in archivos_adjuntos:
            if adjuntar_archivo(msg, ruta):
                attached_count += 1

        if attached_count == 0:
            print("❌ Ningún archivo pudo ser adjuntado (tamaño inválido o errores). Cancelando envío.")
            return

        server.sendmail(remitente_email, destinatarios_lista, msg.as_string())

        print(f"✅ Lote enviado con éxito a destinatarios '{destinatarios_header}'. Asunto: '{asunto_final}'")
        print(f"✅ Total de adjuntos físicamente adjuntados: {attached_count}")

    except smtplib.SMTPAuthenticationError as e:
        print(f"❌ ERROR CRÍTICO [SMTP AUTH]: {e}")
    except smtplib.SMTPServerDisconnected as e:
        print(f"❌ ERROR CRÍTICO [SMTP DISCONNECT]: {e}")
    except Exception as e:
        print(f"❌ ERROR GENÉRICO: {e}")
    finally:
        if server:
            try:
                server.quit()
                print("DEBUG: Conexión SMTP cerrada.")
            except Exception as e:
                print(f"⚠️ Error al cerrar conexión SMTP: {e}")

    print("--- FIN ENVÍO SMTP ---")