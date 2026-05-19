# Replicador de Órdenes Contractuales

## Descripción

Este proyecto es un replicador de órdenes contractuales diseñado para generar solicitudes y anexos a partir de datos de Excel y plantillas.

Incluye tres entornos:

- `SolicitudOrden`: Genera solicitudes en Excel a partir de `datos/2. BaseDatos -enviar.xlsx` y una plantilla de orden.
- `SolicitudOCD_OSU`: Genera solicitudes en Excel a partir de `datos/BaseDatos.xlsx` y la misma plantilla de orden.
- `SolicitudEA`: Orquesta generación, ajuste de texto, creación de anexos en Word, conversión a PDF y envío de correos por lotes.

## Estructura del proyecto

- `main.py`: Menú principal para seleccionar el entorno a ejecutar.
- `SolicitudOrden/`: Módulo básico de generación de solicitudes Excel.
- `SolicitudOCD_OSU/`: Módulo alternativo de generación de solicitudes Excel.
- `SolicitudEA/`: Módulo avanzado con generación de Excel, Word, PDF y envío de correos.

Cada módulo cuenta con:

- `main.py`: Flujo principal de ejecución.
- `config.py`: Mapeo de campos y configuración de hoja Excel.
- `utils.py`: Funciones reutilizables para carga, reemplazo y guardado.

## Requisitos

- Python 3.8+.
- Paquetes Python:
  - `openpyxl`
  - `python-docx`
  - `docx2pdf`
  - `PyPDF2`
  - `language_tool_python`
  - `python-dotenv`

Instalación sugerida:

```bash
pip install openpyxl python-docx docx2pdf PyPDF2 language_tool_python python-dotenv
```

### Requisito adicional para `SolicitudEA`

- LibreOffice (`soffice`) instalado en el sistema para la conversión de Excel a PDF.
- Java instalado si `language_tool_python` lo requiere para la corrección de texto.

## Uso general

Ejecuta el menú principal desde la raíz del proyecto:

```bash
python main.py
```

Selecciona el entorno deseado:

1. `SolicitudOrden`
2. `SolicitudOCD_OSU`
3. `SolicitudEA`

## Uso de cada módulo

### `SolicitudOrden`

- Lee datos de `SolicitudOrden/datos/2. BaseDatos -enviar.xlsx`.
- Carga la plantilla `SolicitudOrden/datos/plantillas/SolicitudOrden.xlsx`.
- Reemplaza las etiquetas definidas en `SolicitudOrden/config.py`.
- Guarda resultados en `SolicitudOrden/output/`.

### `SolicitudOCD_OSU`

- Lee datos de `SolicitudOCD_OSU/datos/BaseDatos.xlsx`.
- Usa la misma plantilla `SolicitudOCD_OSU/datos/plantillas/SolicitudOrden.xlsx`.
- Guarda resultados en `SolicitudOCD_OSU/output/`.

### `SolicitudEA`

- Lee datos de `SolicitudEA/datos/2. BaseDatos -enviar.xlsx`.
- Usa plantilla Excel `SolicitudEA/datos/plantillas/SolicitudOrden.xlsx`.
- Ajusta el tamaño de celdas para campos largos.
- Genera un anexo en Word usando `SolicitudEA/datos/plantillas/Anexo.docx` cuando hay más de 5 obligaciones.
- Convierte archivos Excel y Word a PDF.
- Agrupa órdenes por supervisor y envía correos por lotes.

## Variables de entorno para `SolicitudEA`

Crea un archivo `.env` dentro de `SolicitudEA/` con al menos estas variables:

```env
EMAIL_REMITENTE=tu_correo@example.com
REMITENTE_PASSWORD=tu_contraseña
CORREO_COPIA=copia@example.com
EMAIL_BATCH_SIZE=10
```

- `EMAIL_REMITENTE`: Correo que enviará los mensajes.
- `REMITENTE_PASSWORD`: Contraseña o token del remitente.
- `CORREO_COPIA`: Correo adicional en copia (opcional).
- `EMAIL_BATCH_SIZE`: Tamaño de cada lote de envío.

## Notas importantes

- `SolicitudEA` valida y corrige texto en español usando `language_tool_python`.
- `SolicitudEA` gestiona nombres de archivo seguros y evita sobrescribir archivos existentes.
- Si no se configuran las credenciales de correo, `SolicitudEA` generará los documentos pero no realizará el envío.

## Mejoras posibles

- Añadir validación de datos más estricta antes de reemplazar en plantillas.
- Permitir ejecutar generación sin envío en `SolicitudEA` mediante una opción de configuración.
- Añadir un `requirements.txt` o `pyproject.toml` para facilitar la instalación.
