from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, RedirectResponse
from pydantic import BaseModel
from typing import Optional, List
from scalar_fastapi import scalar_interface
import os
import json
import qrcode
from datetime import datetime
from num2words import num2words

# Librerías para generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Librería para generación de Word (.docx)
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Instancia principal de la aplicación
app = FastAPI(
    title="API Integral de Gestión Documental y Búsqueda para Guinea Ecuatorial",
    version="4.0",
    description="API completa con búsqueda de servicios, plantilla de modelos, generador de PDF y Word con Código QR y formato administrativo local."
)

# Redirección automática desde la raíz para evitar "Not Found"
@app.get("/", include_in_schema=False)
async def root():
    return RedirectResponse(url="/gui")

# Interfaz gráfica limpia y moderna en /gui
@app.get("/gui", include_in_schema=False)
async def scalar_html():
    return scalar_interface(
        openapi_url=app.openapi_url,
        title=app.title,
    )

# ==========================================
# DIRECTORIOS Y CONFIGURACIÓN INICIAL
# ==========================================
DIR_TEMP = "temp_files"
DIR_MODELOS = "modelos_guardados"
DIR_LOGOS = "logos_guardados"

os.makedirs(DIR_TEMP, exist_ok=True)
os.makedirs(DIR_MODELOS, exist_ok=True)
os.makedirs(DIR_LOGOS, exist_ok=True)

RUTA_REGISTRO_MODELOS = os.path.join(DIR_MODELOS, "registro.json")

def cargar_registro():
    if os.path.exists(RUTA_REGISTRO_MODELOS):
        with open(RUTA_REGISTRO_MODELOS, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def guardar_registro(data):
    with open(RUTA_REGISTRO_MODELOS, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# ==========================================
# MODELOS DE DATOS (PYDANTIC)
# ==========================================
class SolicitudInformacion(BaseModel):
    query: str
    categoria: Optional[str] = None

# ==========================================
# RUTAS DE BÚSQUEDA Y SERVICIOS
# ==========================================
@app.get("/buscar/comercial", summary="Buscar Productos Servicios")
async def buscar_comercial(query: str):
    return {
        "query": query,
        "resultados": [
            {"id": 1, "titulo": f"Resultado 1 para '{query}'", "descripcion": "Servicios administrativos y comerciales en Malabo y Bata."},
            {"id": 2, "titulo": f"Resultado 2 para '{query}'", "descripcion": "Gestión documental oficial y tramitación."}
        ]
    }

@app.post("/buscar/informacion", summary="Buscar Información")
async def buscar_informacion(solicitud: SolicitudInformacion):
    return {
        "busqueda": solicitud.query,
        "categoria": solicitud.categoria or "General",
        "estado": "Completado",
        "detalles": "Información procesada correctamente para Guinea Ecuatorial."
    }

# ==========================================
# RUTAS DE PLANTILLAS Y MODELOS
# ==========================================
@app.post("/modelos/guardar", summary="Guardar Modelo")
async def guardar_modelo(nombre: str = Form(...), descripcion: str = Form(...), archivo: UploadFile = File(...)):
    extension = os.path.splitext(archivo.filename)[1]
    nombre_archivo = f"{nombre}{extension}"
    ruta_guardado = os.path.join(DIR_MODELOS, nombre_archivo)
    
    with open(ruta_guardado, "wb") as f:
        f.write(await archivo.read())
        
    registro = cargar_registro()
    registro[nombre] = {
        "descripcion": descripcion,
        "archivo": nombre_archivo,
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    guardar_registro(registro)
    
    return {"mensaje": "Modelo guardado con éxito", "modelo": nombre}

@app.get("/modelos/listar", summary="Listar Modelos")
async def listar_modelos():
    return cargar_registro()

# ==========================================
# RUTAS DE UTILIDADES
# ==========================================
@app.get("/utilidades/monto-letras", summary="Convertir Monto")
async def monto_a_letras(monto: float):
    try:
        letras = num2words(monto, lang='es').capitalize()
        return {"monto": monto, "letras": f"{letras} francos CFA"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# ==========================================
# GENERADOR DE DOCUMENTOS (PDF Y WORD)
# ==========================================
@app.post("/generar-documento", summary="Generar Documento")
async def generar_documento(
    titulo: str = Form(...),
    subtitulo: Optional[str] = Form(None),
    contenido: str = Form(...),
    remitente: str = Form(...),
    cargo_remitente: Optional[str] = Form(None),
    destinatario: str = Form(...),
    formato: str = Form("pdf"),  # "pdf" o "word"
    incluir_qr: bool = Form(True)
):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_base = f"doc_{timestamp}"
    
    # 1. Generar Código QR opcional
    ruta_qr = None
    if incluir_qr:
        ruta_qr = os.path.join(DIR_TEMP, f"qr_{timestamp}.png")
        info_qr = f"Doc: {titulo}\nRemitente: {remitente}\nFecha: {datetime.now().strftime('%Y-%m-%d')}"
        img_qr = qrcode.make(info_qr)
        img_qr.save(ruta_qr)

    # 2. Generación en PDF
    if formato.lower() == "pdf":
        ruta_salida = os.path.join(DIR_TEMP, f"{nombre_base}.pdf")
        doc = SimpleDocTemplate(ruta_salida, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
        styles = getSampleStyleSheet()
        story = []

        style_title = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor("#1A365D"), alignment=1)
        style_sub = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=12, leading=14, textColor=colors.HexColor("#4A5568"), alignment=1)
        style_body = ParagraphStyle('Body', parent=styles['Normal'], fontSize=11, leading=16, textColor=colors.HexColor("#2D3748"))

        story.append(Paragraph(titulo.upper(), style_title))
        if subtitulo:
            story.append(Paragraph(subtitulo, style_sub))
        story.append(Spacer(1, 20))

        story.append(Paragraph(f"<b>Para:</b> {destinatario}", style_body))
        story.append(Paragraph(f"<b>De:</b> {remitente}" + (f" ({cargo_remitente})" if cargo_remitente else ""), style_body))
        story.append(Paragraph(f"<b>Fecha:</b> {datetime.now().strftime('%d/%m/%Y')}", style_body))
        story.append(Spacer(1, 15))

        for parrafo in contenido.split('\n'):
            if parrafo.strip():
                story.append(Paragraph(parrafo, style_body))
                story.append(Spacer(1, 10))

        if ruta_qr and os.path.exists(ruta_qr):
            story.append(Spacer(1, 15))
            story.append(RLImage(ruta_qr, width=80, height=80))

        doc.build(story)
        return FileResponse(ruta_salida, filename=f"{titulo}.pdf", media_type="application/pdf")

    # 3. Generación en Word (.docx)
    elif formato.lower() == "word":
        ruta_salida = os.path.join(DIR_TEMP, f"{nombre_base}.docx")
        doc = Document()

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_title = p_title.add_run(titulo.upper())
        run_title.bold = True
        run_title.font.size = Pt(18)
        run_title.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)

        if subtitulo:
            p_sub = doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_sub = p_sub.add_run(subtitulo)
            run_sub.font.size = Pt(12)
            run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

        doc.add_paragraph(f"Para: {destinatario}")
        doc.add_paragraph(f"De: {remitente}" + (f" ({cargo_remitente})" if cargo_remitente else ""))
        doc.add_paragraph(f"Fecha: {datetime.now().strftime('%d/%m/%Y')}")
        doc.add_paragraph()

        for parrafo in contenido.split('\n'):
            if parrafo.strip():
                doc.add_paragraph(parrafo)

        if ruta_qr and os.path.exists(ruta_qr):
            doc.add_picture(ruta_qr, width=Inches(1.2))

        doc.save(ruta_salida)
        return FileResponse(ruta_salida, filename=f"{titulo}.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    else:
        raise HTTPException(status_code=400, detail="Formato no soportado. Usa 'pdf' o 'word'.")
