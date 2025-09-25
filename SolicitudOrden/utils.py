# --- Librerías estándar ---
import io
import math
import os
import subprocess

# --- Librerías de terceros ---
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx2pdf import convert

from PyPDF2 import PdfReader, PdfWriter
import language_tool_python

# --- Módulos locales ---
from config import CAMPOS, HOJA_SOLICITUD, OBLIGACIONES

# Inicializa el corrector ortográfico
tool = language_tool_python.LanguageTool('es')

def corregir_texto(texto: str) -> str:
    """Corrige ortografía y gramática en un texto en español"""
    if not texto:
        return texto
    matches = tool.check(texto)
    return language_tool_python.utils.correct(texto, matches)

def cargar_plantilla(ruta_plantilla):
    """Carga el archivo plantilla de solicitud"""
    return openpyxl.load_workbook(ruta_plantilla)

def cargar_datos(ruta_datos):
    """Carga el archivo de datos fuente"""
    return openpyxl.load_workbook(ruta_datos)

def buscar_y_reemplazar(hoja, texto_buscar, texto_reemplazar):
    """Busca y reemplaza texto en una hoja de Excel"""
    for row in hoja.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and texto_buscar in cell.value:
                cell.value = cell.value.replace(texto_buscar, texto_reemplazar)

def ajustar_tamano_celdas(hoja: Worksheet, celdas_a_ajustar: list,
                         ancho_max_columna: int = 31, factor_ajuste: float = 1.05):
    """
    Ajusta la altura de las filas y el ancho de las columnas para que el texto quepa
    de forma más realista según la plantilla base (Calibri 11, columnas medianas).
    """
    filas_procesadas = set()

    for fila, col in celdas_a_ajustar:
        celda = hoja.cell(row=fila, column=col)
        valor_celda = str(celda.value) if celda.value else ""

        if fila in filas_procesadas or not valor_celda:
            continue

        celda.alignment = Alignment(wrapText=True, horizontal='left', vertical='center')

        # Ancho útil estimado
        caracteres_por_linea = int(ancho_max_columna * factor_ajuste)

        if caracteres_por_linea > 0:
            num_lineas = math.ceil(len(valor_celda) / caracteres_por_linea)
            nueva_altura = num_lineas * 2  # 14 pt por línea (Calibri 11)

            # Limitar alturas
            nueva_altura = max(15, min(nueva_altura, 300))

            hoja.row_dimensions[fila].height = nueva_altura

            #print(f"Fila {fila}: {len(valor_celda)} chars, "
            #      f"{num_lineas} líneas → altura={nueva_altura}")
            filas_procesadas.add(fila)

def procesar_solicitud(wb_plantilla, datos_fila):
    """Procesa la plantilla con los datos de una fila"""
    hoja = wb_plantilla[HOJA_SOLICITUD]

    # Reemplazar todos los campos marcados, excluyendo las obligaciones
    for campo_db, campo_plantilla in CAMPOS.items():
        if campo_db not in OBLIGACIONES:
            valor = datos_fila.get(campo_db, '')
            valor = corregir_texto(str(valor))
            buscar_y_reemplazar(hoja, campo_plantilla, valor)
    
    # Construir lista de obligaciones corregidas
    obligaciones_reales = [
        corregir_texto(str(datos_fila.get(campo)).strip())
        for campo in OBLIGACIONES
        if datos_fila.get(campo) and str(datos_fila.get(campo)).strip() != ""
    ]

    num_obligaciones = len(obligaciones_reales)

    if num_obligaciones <= 5:
        print(f"Número de obligaciones: {num_obligaciones}. Agregando al Excel.")
        for i, obligacion in enumerate(obligaciones_reales, start=1):
            marcador_excel = f"{{obligacion_{i}}}"
            buscar_y_reemplazar(hoja, marcador_excel, obligacion)
        for i in range(num_obligaciones + 1, 6):
            marcador_excel = f"{{obligacion_{i}}}"
            buscar_y_reemplazar(hoja, marcador_excel, "")
    else: 
        print(f"Número de obligaciones: {num_obligaciones}. Se generará un documento de Word y se limpiarán los marcadores en Excel.")
        for i in range(1, 6):
            marcador_excel = f"{{obligacion_{i}}}"
            buscar_y_reemplazar(hoja, marcador_excel, " ")

    return wb_plantilla

def guardar_solicitud(wb, nombre_archivo, ruta_output):
    """Guarda la solicitud generada"""
    ruta_completa = os.path.join(ruta_output, nombre_archivo)
    wb.save(ruta_completa)
    return ruta_completa

def generar_word_obligaciones(datos_fila, ruta_output, nombre_archivo, ruta_plantilla_word):
    """
    Genera un documento de Word con todas las obligaciones si hay 6 o más.
    Si hay 5 o menos, no hace nada.
    """
    obligaciones = []
    for campo in OBLIGACIONES:
        valor = datos_fila.get(campo)
        if valor and str(valor).strip() != "":
            obligaciones.append(corregir_texto(str(valor).strip()))

    if len(obligaciones) <= 5:
        print(f"No se genera Word: solo {len(obligaciones)} obligaciones (se necesitan 6 o más).")
        return None

    print(f"Generando documento de Word con {len(obligaciones)} obligaciones.")
    
    try:
        doc = Document(ruta_plantilla_word)
    except Exception as e:
        print(f"Error al cargar la plantilla de Word: {e}")
        return None

    for p in doc.paragraphs:
        for campo, valor in datos_fila.items():
            if campo in OBLIGACIONES or valor is None:
                continue
            marcador_word = f"{{{campo}}}"
            if marcador_word in p.text:
                p.text = p.text.replace(marcador_word, corregir_texto(str(valor)))
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(9)

    for p in doc.paragraphs:
        if "{obligaciones}" in p.text:
            p.clear()
            for i, obligacion in enumerate(obligaciones, start=1):
                new_p = p.insert_paragraph_before(f"{i}. {obligacion}")
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                new_p.paragraph_format.space_before = Pt(0)
                new_p.paragraph_format.space_after = Pt(0)
                new_p.paragraph_format.first_line_indent = Pt(0)
                for run in new_p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(9)
            p._element.getparent().remove(p._element)
            break

    nombre_doc = f"{nombre_archivo.replace('.xlsx', '.docx')}"
    ruta_doc = os.path.join(ruta_output, nombre_doc)
    doc.save(ruta_doc)
    print(f"Documento Word generado: {ruta_doc}")
    return ruta_doc


def excel_a_pdf(ruta_excel, ruta_pdf, paginas_to_keep=None, timeout=30, wait_interval=0.5):
    """
    Convierte un Excel a PDF usando LibreOffice (requiere instalado)
    y conserva solo las páginas indicadas (por defecto, la primera página).
    """
    if paginas_to_keep is None:
        paginas_to_keep = [0]  # por defecto solo la primera página

    try:
        # 1) Ejecutar LibreOffice → genera el PDF con el mismo nombre base que el Excel
        subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf", "--outdir",
            os.path.dirname(ruta_pdf), ruta_excel
        ], check=True)

        # 2) Determinar el nombre real generado por LibreOffice
        generated_pdf = os.path.join(
            os.path.dirname(ruta_excel),
            os.path.splitext(os.path.basename(ruta_excel))[0] + ".pdf"
        )

        # ⚠️ Puede que ruta_pdf sea distinto → verificamos y ajustamos
        if not os.path.exists(generated_pdf):
            print(f"❌ No se encontró el PDF generado por LibreOffice en: {generated_pdf}")
            return

        # 3) Leer PDF generado
        with open(generated_pdf, "rb") as f:
            pdf_bytes = f.read()
        reader = PdfReader(io.BytesIO(pdf_bytes))

        total_pages = len(reader.pages)
        if total_pages == 0:
            print(f"❌ El PDF generado no contiene páginas: {generated_pdf}")
            return

        # 4) Crear nuevo PDF solo con páginas deseadas
        writer = PdfWriter()
        for idx in paginas_to_keep:
            if 0 <= idx < total_pages:
                writer.add_page(reader.pages[idx])
            else:
                print(f"⚠️ Índice fuera de rango: {idx} (total={total_pages})")

        if not writer.pages:  # fallback
            writer.add_page(reader.pages[0])

        # 5) Guardar al destino final
        with open(ruta_pdf, "wb") as f_out:
            writer.write(f_out)

        # 6) Borrar el PDF completo que generó LibreOffice (para no duplicar)
        if os.path.abspath(generated_pdf) != os.path.abspath(ruta_pdf):
            try:
                os.remove(generated_pdf)
            except Exception as e:
                print(f"⚠️ No se pudo borrar el PDF original de LibreOffice: {e}")

        print(f"✅ PDF final generado: {ruta_pdf}")

    except subprocess.CalledProcessError as cpe:
        print(f"❌ LibreOffice falló al convertir el archivo: {cpe}")
    except Exception as e:
        print(f"❌ Error en excel_a_pdf: {e}")

def word_a_pdf(ruta_word, ruta_pdf):
    """Convierte un Word a PDF"""
    convert(ruta_word, ruta_pdf)