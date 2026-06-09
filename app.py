import os
import json
import re
import uuid
from datetime import date
from io import BytesIO

from flask import Flask, render_template, request, jsonify, send_file, session
import anthropic

# Server-side store to avoid Flask cookie 4 KB limit
_doc_store: dict = {}

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Table, TableStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "memorandos-sushi-break-2024")

EMPRESA = "SUSHI BREAK SAS"
NIT = "901359641-1"
REPRESENTANTE = "Juan Manuel Mercado"
CARGO_REP = "Representante Legal"

SEDES = ["NOVENA", "NORTE", "OESTE", "V LILI", "JAMUNDI", "PALMIRA", "CP", "MARBELLA"]

TIPOS_DOCUMENTO = {
    "llamado":     "MEMORANDO DE LLAMADO DE ATENCIÓN",
    "descargos":   "CITACIÓN A DILIGENCIA DE DESCARGOS",
    "sancion":     "MEMORANDO DE SANCIÓN DISCIPLINARIA",
    "terminacion": "CARTA DE TERMINACIÓN DE CONTRATO",
}

SYSTEM_PROMPT = """Eres un experto en derecho laboral colombiano especializado en el Código Sustantivo del Trabajo (CST) y reglamentos internos de trabajo. Tu tarea es redactar documentos disciplinarios laborales formales según el tipo que te indique el usuario.

ARTÍCULOS DEL CST SEGÚN TIPO DE DOCUMENTO:

LLAMADO DE ATENCIÓN:
- CST Art. 58 — Obligaciones del trabajador
- CST Art. 60 — Prohibiciones al trabajador
- CST Art. 111 — Reglamento interno de trabajo

CITACIÓN A DESCARGOS:
- CST Art. 115 — Procedimiento disciplinario (derecho a ser escuchado)
- CST Art. 62 — Causas justas de terminación
- CST Art. 29 — Debido proceso

SANCIÓN DISCIPLINARIA (suspensión sin sueldo):
- CST Art. 112 — Sanciones disciplinarias
- CST Art. 113 — Límites de las sanciones
- CST Art. 115 — Procedimiento para imponer sanciones
- CST Art. 58 — Obligaciones del trabajador

TERMINACIÓN DE CONTRATO:
- CST Art. 62 — Causas justas de terminación con justa causa
- CST Art. 64 — Indemnización por terminación sin justa causa (si aplica)
- CST Art. 65 — Indemnización moratoria

INSTRUCCIONES:
1. El usuario ya eligió el tipo de documento — NO lo cambies ni lo cuestiones
2. Redacta el documento completo en tono formal, legal y profesional
3. Cita los artículos del CST que correspondan al tipo de documento
4. El documento debe ser completo, bien redactado y listo para imprimir
5. Para terminación de contrato, menciona la justa causa según el Art. 62 CST

RESPONDE ÚNICAMENTE con un JSON válido con esta estructura exacta:
{
  "tipo_falta": "LEVE|MEDIA|GRAVE",
  "clasificacion": "descripción breve de la situación",
  "articulos": ["Art. X CST - descripción", ...],
  "recomendacion": "observación legal relevante",
  "tipo_documento": "nombre del documento tal como lo indicó el usuario",
  "documento": {
    "asunto": "texto del asunto",
    "cuerpo": "texto completo del cuerpo del documento con saltos de línea \\n"
  }
}"""


@app.route("/")
def index():
    today = date.today().strftime("%Y-%m-%d")
    return render_template("index.html", sedes=SEDES, today=today)


@app.route("/analizar", methods=["POST"])
def analizar():
    data = request.get_json()
    nombre = data.get("nombre", "").strip()
    cargo = data.get("cargo", "").strip()
    sede = data.get("sede", "").strip()
    fecha = data.get("fecha", "").strip()
    reincidente = data.get("reincidente", "no")
    descripcion = data.get("descripcion", "").strip()
    tipo_key = data.get("tipo_documento", "llamado")

    if not all([nombre, cargo, sede, fecha, descripcion]):
        return jsonify({"error": "Todos los campos son obligatorios"}), 400

    tipo_doc_label = TIPOS_DOCUMENTO.get(tipo_key, "MEMORANDO DE LLAMADO DE ATENCIÓN")

    user_message = f"""Genera el siguiente documento disciplinario:

TIPO DE DOCUMENTO SOLICITADO: {tipo_doc_label}

DATOS DEL EMPLEADO:
- Nombre: {nombre}
- Cargo: {cargo}
- Sede: {sede}
- Fecha del documento: {fecha}
- ¿Es reincidente?: {reincidente.upper()}

DESCRIPCIÓN DE LA SITUACIÓN:
{descripcion}

Redacta el documento completo del tipo indicado ({tipo_doc_label}). Usa los artículos del CST correspondientes a ese tipo de documento."""

    try:
        client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
        message = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=2048,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_message}]
        )
        raw = message.content[0].text.strip()

        # Extract JSON even if wrapped in markdown code block
        json_match = re.search(r'\{[\s\S]*\}', raw)
        if not json_match:
            return jsonify({"error": "Respuesta inválida del modelo"}), 500
        resultado = json.loads(json_match.group())

        # Store in server-side dict (avoids Flask cookie 4 KB limit)
        doc_id = str(uuid.uuid4())
        _doc_store[doc_id] = {
            "resultado": resultado,
            "empleado": {
                "nombre": nombre,
                "cargo": cargo,
                "sede": sede,
                "fecha": fecha,
                "reincidente": reincidente,
            },
        }
        # Keep only the last 50 docs to avoid unbounded memory growth
        if len(_doc_store) > 50:
            oldest = next(iter(_doc_store))
            del _doc_store[oldest]

        session["doc_id"] = doc_id
        return jsonify(resultado)

    except json.JSONDecodeError:
        return jsonify({"error": "No se pudo parsear la respuesta del modelo"}), 500
    except anthropic.APIError as e:
        return jsonify({"error": f"Error de API: {str(e)}"}), 500
    except Exception as e:
        return jsonify({"error": f"Error inesperado: {str(e)}"}), 500


def _format_date_spanish(fecha_str):
    months = {
        "01": "enero", "02": "febrero", "03": "marzo", "04": "abril",
        "05": "mayo", "06": "junio", "07": "julio", "08": "agosto",
        "09": "septiembre", "10": "octubre", "11": "noviembre", "12": "diciembre"
    }
    try:
        parts = fecha_str.split("-")
        return f"{int(parts[2])} de {months[parts[1]]} de {parts[0]}"
    except Exception:
        return fecha_str


@app.route("/generar_pdf")
def generar_pdf():
    doc_id = session.get("doc_id")
    entry = _doc_store.get(doc_id) if doc_id else None
    if not entry:
        return "No hay documento generado. Vuelve a analizar la situación.", 400
    resultado = entry["resultado"]
    empleado = entry["empleado"]

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=2.5 * cm,
        leftMargin=2.5 * cm,
        topMargin=2 * cm,
        bottomMargin=2.5 * cm
    )

    styles = getSampleStyleSheet()
    story = []

    # Header style
    header_style = ParagraphStyle(
        "Header",
        parent=styles["Normal"],
        fontSize=14,
        fontName="Helvetica-Bold",
        textColor=colors.white,
        alignment=TA_CENTER,
        spaceAfter=2,
    )
    sub_header_style = ParagraphStyle(
        "SubHeader",
        parent=styles["Normal"],
        fontSize=10,
        fontName="Helvetica",
        textColor=colors.white,
        alignment=TA_CENTER,
        spaceAfter=2,
    )
    title_style = ParagraphStyle(
        "Title",
        parent=styles["Normal"],
        fontSize=12,
        fontName="Helvetica-Bold",
        alignment=TA_CENTER,
        spaceBefore=14,
        spaceAfter=14,
        textColor=colors.HexColor("#1a1a2e"),
    )
    label_style = ParagraphStyle(
        "Label",
        parent=styles["Normal"],
        fontSize=9,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#444444"),
        spaceAfter=2,
    )
    value_style = ParagraphStyle(
        "Value",
        parent=styles["Normal"],
        fontSize=10,
        fontName="Helvetica",
        textColor=colors.black,
        spaceAfter=6,
    )
    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontSize=10,
        fontName="Helvetica",
        alignment=TA_JUSTIFY,
        spaceAfter=10,
        leading=16,
    )
    sign_style = ParagraphStyle(
        "Sign",
        parent=styles["Normal"],
        fontSize=10,
        fontName="Helvetica-Bold",
        alignment=TA_CENTER,
    )

    # Dark header block
    header_data = [
        [Paragraph(EMPRESA, header_style)],
        [Paragraph(f"NIT: {NIT}", sub_header_style)],
        [Paragraph(resultado["tipo_documento"], sub_header_style)],
    ]
    header_table = Table(header_data, colWidths=[16 * cm])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1a1a2e")),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 0.4 * cm))

    # Employee info table
    fecha_es = _format_date_spanish(empleado["fecha"])
    info_data = [
        [Paragraph("PARA:", label_style), Paragraph(f"{empleado['nombre']} — {empleado['cargo']}", value_style)],
        [Paragraph("SEDE:", label_style), Paragraph(empleado["sede"], value_style)],
        [Paragraph("DE:", label_style), Paragraph(f"{REPRESENTANTE} — {CARGO_REP}", value_style)],
        [Paragraph("ASUNTO:", label_style), Paragraph(resultado["documento"]["asunto"], value_style)],
        [Paragraph("FECHA:", label_style), Paragraph(fecha_es, value_style)],
    ]
    info_table = Table(info_data, colWidths=[2.5 * cm, 13.5 * cm])
    info_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, -1), (-1, -1), 0.5, colors.HexColor("#cccccc")),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 0.4 * cm))

    # Body paragraphs
    for line in resultado["documento"]["cuerpo"].split("\n"):
        line = line.strip()
        if line:
            story.append(Paragraph(line, body_style))

    story.append(Spacer(1, 1.5 * cm))
    story.append(HRFlowable(width="40%", thickness=1, color=colors.HexColor("#1a1a2e"), hAlign="CENTER"))
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph(REPRESENTANTE, sign_style))
    story.append(Paragraph(CARGO_REP, ParagraphStyle("sc", parent=sign_style, fontName="Helvetica", fontSize=9)))
    story.append(Paragraph(EMPRESA, ParagraphStyle("sc2", parent=sign_style, fontName="Helvetica", fontSize=9)))

    doc.build(story)
    buffer.seek(0)

    filename = f"memorando_{empleado['nombre'].replace(' ', '_')}.pdf"
    return send_file(buffer, mimetype="application/pdf",
                     as_attachment=True, download_name=filename)


@app.route("/generar_word")
def generar_word():
    doc_id = session.get("doc_id")
    entry = _doc_store.get(doc_id) if doc_id else None
    if not entry:
        return "No hay documento generado. Vuelve a analizar la situación.", 400
    resultado = entry["resultado"]
    empleado = entry["empleado"]

    document = Document()

    # Page margins
    for section in document.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    def set_cell_bg(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color)
        tcPr.append(shd)

    # Header table
    header_table = document.add_table(rows=3, cols=1)
    header_table.style = "Table Grid"
    rows_data = [EMPRESA, f"NIT: {NIT}", resultado["tipo_documento"]]
    font_sizes = [14, 10, 10]
    for i, (text, fsize) in enumerate(zip(rows_data, font_sizes)):
        cell = header_table.rows[i].cells[0]
        set_cell_bg(cell, "1A1A2E")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = (i == 0)
        run.font.size = Pt(fsize)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)

    document.add_paragraph()

    # Employee info
    fecha_es = _format_date_spanish(empleado["fecha"])
    fields = [
        ("PARA:", f"{empleado['nombre']} — {empleado['cargo']}"),
        ("SEDE:", empleado["sede"]),
        ("DE:", f"{REPRESENTANTE} — {CARGO_REP}"),
        ("ASUNTO:", resultado["documento"]["asunto"]),
        ("FECHA:", fecha_es),
    ]
    info_table = document.add_table(rows=len(fields), cols=2)
    info_table.style = "Table Grid"
    col_widths = [Cm(2.8), Cm(13.2)]
    for row_idx, (label, value) in enumerate(fields):
        row = info_table.rows[row_idx]
        row.cells[0].width = col_widths[0]
        row.cells[1].width = col_widths[1]

        lp = row.cells[0].paragraphs[0]
        lr = lp.add_run(label)
        lr.bold = True
        lr.font.size = Pt(9)
        lr.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

        vp = row.cells[1].paragraphs[0]
        vr = vp.add_run(value)
        vr.font.size = Pt(10)

    document.add_paragraph()

    # Body
    for line in resultado["documento"]["cuerpo"].split("\n"):
        line = line.strip()
        if line:
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.add_run(line)
            run.font.size = Pt(10)
            p.paragraph_format.space_after = Pt(8)

    # Signature
    document.add_paragraph()
    document.add_paragraph()
    sig_p = document.add_paragraph()
    sig_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sig_run = sig_p.add_run("_" * 35)
    sig_run.font.size = Pt(10)

    for line in [REPRESENTANTE, CARGO_REP, EMPRESA]:
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(line)
        run.bold = (line == REPRESENTANTE)
        run.font.size = Pt(10)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)

    filename = f"memorando_{empleado['nombre'].replace(' ', '_')}.docx"
    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=filename
    )


if __name__ == "__main__":
    app.run(debug=True)
