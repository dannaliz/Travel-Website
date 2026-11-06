"""Genera los PDFs, documentos Word, Excel y TXT del sitio Viajes Horizonte.

Uso (desde la carpeta del proyecto):
    pip install reportlab python-docx openpyxl
    python herramientas/generar_archivos.py

Los PDFs incluyen enlaces a las páginas exclusivas (carpeta exclusivo/), que solo
muestran su contenido cuando se abren desde esos enlaces.
"""
from pathlib import Path
from urllib.parse import quote

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (Flowable, KeepTogether, ListFlowable, ListItem,
                                Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

# Dirección pública del sitio (GitHub Pages). Los enlaces del PDF deben ser absolutos.
SITIO = "https://dannaliz.github.io/Travel-Website/"

RAIZ = Path(__file__).resolve().parent.parent
AZUL = colors.HexColor("#0b4f6c")
CORAL = colors.HexColor("#f26b4f")
GRIS = colors.HexColor("#5a6b78")
ARENA = colors.HexColor("#fbf5ec")

DESTINOS = {
    "cancun": {
        "titulo": "Cancún, México",
        "duracion": "7 días / 6 noches",
        "paquete": "Vuelo + hotel todo incluido",
        "clave": "cun-2026",
        "consulta": "Cancún",
        "sobre": [
            "7 días / 6 noches con vuelo y hotel todo incluido en el Caribe mexicano. "
            "Cancún combina playas de arena blanca y mar turquesa con visitas a Isla Mujeres, "
            "Chichén Itzá y Tulum.",
        ],
        "datos": [
            ("Moneda", "Peso mexicano (MXN)"),
            ("Idioma", "Español; inglés en zonas turísticas"),
            ("Clima", "26–33 °C; lluvias de junio a octubre"),
            ("Zona horaria", "UTC−5 todo el año"),
            ("Documentos", "Pasaporte o identificación oficial"),
        ],
        "dias": [
            ("Llegada", "Traslado al hotel, tarde de playa y cena de bienvenida."),
            ("Descanso", "Día libre para disfrutar del hotel todo incluido."),
            ("Isla Mujeres", "Catamarán, snorkel y tarde en Playa Norte."),
            ("Chichén Itzá", "Zona arqueológica, cenote Ik Kil y el pueblo mágico de Valladolid."),
            ("Tulum y Playa del Carmen", "Ruinas frente al mar y paseo por la Quinta Avenida."),
            ("Día libre", "Actividades opcionales y cena de despedida."),
            ("Regreso", "Desayuno y traslado al aeropuerto."),
        ],
        "incluye": [
            "Vuelo redondo en clase turista.",
            "6 noches en hotel todo incluido.",
            "Traslados aeropuerto – hotel – aeropuerto.",
            "Excursiones con guía a Isla Mujeres, Chichén Itzá y Tulum.",
            "Seguro de viajero.",
        ],
        "no_incluye": [
            "Impuesto ambiental (Visitax) y de saneamiento.",
            "Actividades opcionales.",
            "Propinas y gastos personales.",
        ],
        "recomendaciones": [
            "Usa bloqueador biodegradable en cenotes y arrecifes.",
            "Lleva ropa ligera, sombrero y calzado cómodo.",
            "Ten efectivo en pesos para propinas y mercados.",
            "Hidrátate bien: el calor y la humedad son intensos.",
        ],
    },
    "paris": {
        "titulo": "París, Francia",
        "duracion": "10 días / 9 noches",
        "paquete": "Vuelo + hotel + tours",
        "clave": "par-2026",
        "consulta": "París",
        "sobre": [
            "París, la Ciudad de la Luz, es una de las capitales más visitadas del mundo. Famosa por "
            "su arte, su arquitectura, su gastronomía y su moda, ofrece monumentos icónicos como la "
            "Torre Eiffel, la catedral de Notre-Dame, el Arco del Triunfo y el Museo del Louvre.",
            "Recorrer sus barrios, como Montmartre o Le Marais, caminar junto al río Sena y disfrutar "
            "de un café en una terraza son parte esencial de la experiencia parisina.",
        ],
        "datos": [
            ("Moneda", "Euro (EUR)"),
            ("Idioma", "Francés; inglés en zonas turísticas"),
            ("Clima", "Verano 20–26 °C; invierno 3–8 °C"),
            ("Zona horaria", "UTC+1 (UTC+2 en verano)"),
            ("Documentos", "Pasaporte vigente; requisitos del espacio Schengen"),
        ],
        "dias": [
            ("Llegada a París", "Llegada al aeropuerto Charles de Gaulle, traslado al hotel y tiempo libre."),
            ("Torre Eiffel y crucero por el Sena", "Subida al segundo piso de la torre, paseo por los Campos de Marte y crucero al atardecer."),
            ("Museo del Louvre", "Visita guiada para conocer la Mona Lisa y la Venus de Milo. Tarde libre en el Jardín de las Tullerías."),
            ("Île de la Cité y Barrio Latino", "Notre-Dame, la Sainte-Chapelle y librerías y cafés históricos."),
            ("Palacio de Versalles", "Excursión de día completo: Galería de los Espejos, apartamentos reales y jardines."),
            ("Montmartre", "Basílica del Sagrado Corazón y la Plaza du Tertre. Noche opcional de cabaret."),
            ("Campos Elíseos y Arco del Triunfo", "Caminata, subida al Arco del Triunfo y tarde de compras."),
            ("Museo de Orsay y Le Marais", "La mayor colección de arte impresionista del mundo y paseo por Le Marais."),
            ("Día libre", "Actividades opcionales como Disneyland París o el Valle del Loira. Cena de despedida."),
            ("Regreso", "Desayuno, check-out y traslado al aeropuerto."),
        ],
        "incluye": [
            "Vuelo redondo en clase turista.",
            "9 noches de hotel con desayuno incluido.",
            "Traslados aeropuerto – hotel – aeropuerto.",
            "Entradas y visitas guiadas a la Torre Eiffel, el Louvre, Versalles y Orsay.",
            "Crucero por el río Sena.",
            "Seguro de viajero con cobertura en Europa.",
        ],
        "no_incluye": [
            "Comidas y cenas no especificadas.",
            "Tasa turística municipal (se paga en el hotel).",
            "Actividades marcadas como opcionales.",
        ],
        "recomendaciones": [
            "Compra un pase de transporte para moverte en metro de forma económica.",
            "Cuida tus pertenencias en el metro y en lugares concurridos.",
            "Lleva calzado cómodo: París se disfruta caminando.",
            "Saluda con un “Bonjour” al entrar a tiendas y restaurantes.",
        ],
    },
}


def url_exclusiva(clave_destino, ancla):
    d = DESTINOS[clave_destino]
    return f"{SITIO}exclusivo/{clave_destino}.html?desde=pdf&clave={d['clave']}#{ancla}"


# ---------------------------------------------------------------- PDF ----------

class BotonEnlace(Flowable):
    """Botón redondeado y clicable dentro del PDF."""

    def __init__(self, texto, url, ancho=10.5 * cm, alto=1.0 * cm, color=CORAL):
        super().__init__()
        self.texto, self.url, self.width, self.height, self.color = texto, url, ancho, alto, color

    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.roundRect(0, 0, self.width, self.height, self.height / 2, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 10.5)
        c.drawCentredString(self.width / 2, self.height / 2 - 3.6, self.texto)
        c.linkURL(self.url, (0, 0, self.width, self.height), relative=1, thickness=0)


def generar_pdf(clave):
    d = DESTINOS[clave]
    salida = RAIZ / "pdfs" / f"itinerario-{clave}.pdf"
    base = getSampleStyleSheet()
    est = {
        "marca": ParagraphStyle("marca", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=9, textColor=CORAL, spaceAfter=4),
        "titulo": ParagraphStyle("titulo", parent=base["Title"], fontSize=26, leading=30, textColor=AZUL, alignment=0, spaceAfter=4),
        "sub": ParagraphStyle("sub", parent=base["Normal"], fontSize=12, textColor=GRIS, spaceAfter=14),
        "h2": ParagraphStyle("h2", parent=base["Heading2"], fontSize=15, textColor=AZUL, spaceBefore=14, spaceAfter=6),
        "p": ParagraphStyle("p", parent=base["Normal"], fontSize=10.5, leading=15),
        "dia": ParagraphStyle("dia", parent=base["Normal"], fontSize=10.5, leading=15, spaceAfter=6),
        "nota": ParagraphStyle("nota", parent=base["Normal"], fontSize=9, leading=12, textColor=GRIS, alignment=TA_CENTER),
        "caja": ParagraphStyle("caja", parent=base["Normal"], fontSize=11, leading=15, textColor=AZUL, alignment=TA_CENTER),
    }
    enlace = lambda texto, url: f'<a href="{url}" color="#f26b4f"><u>{texto}</u></a>'

    hist = []
    hist += [Paragraph("VIAJES HORIZONTE · ITINERARIO DE VIAJE", est["marca"]),
             Paragraph(d["titulo"], est["titulo"]),
             Paragraph(f'{d["duracion"]} · {d["paquete"]} · '
                       + enlace("Ver precios por temporada", url_exclusiva(clave, "precios")), est["sub"])]

    hist.append(Paragraph("Sobre el destino", est["h2"]))
    for p in d["sobre"]:
        hist.append(Paragraph(p, est["p"]))
        hist.append(Spacer(1, 4))
    hist.append(Paragraph(
        "¿Sabías que…? Preparamos una sección con "
        + enlace("datos curiosos de este destino", url_exclusiva(clave, "curiosidades"))
        + " que solo puedes abrir desde este PDF.", est["p"]))

    hist.append(Paragraph("Datos útiles", est["h2"]))
    tabla = Table([[Paragraph(f"<b>{k}</b>", est["p"]), Paragraph(v, est["p"])] for k, v in d["datos"]],
                  colWidths=[4 * cm, 12.5 * cm])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), ARENA),
        ("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.HexColor("#e4ddd2")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    hist.append(tabla)

    hist.append(Paragraph("Itinerario día a día", est["h2"]))
    for i, (titulo, texto) in enumerate(d["dias"], 1):
        hist.append(Paragraph(f'<font color="#0b4f6c"><b>Día {i} – {titulo}.</b></font> {texto}', est["dia"]))

    def lista(items):
        return ListFlowable([ListItem(Paragraph(t, est["p"]), leftIndent=12) for t in items],
                            bulletType="bullet", start="•", leftIndent=12)

    hist.append(Paragraph("El paquete incluye", est["h2"]))
    hist.append(lista(d["incluye"]))
    hist.append(Paragraph("No incluye", est["h2"]))
    hist.append(lista(d["no_incluye"]))
    hist.append(Spacer(1, 4))
    hist.append(Paragraph("Consulta los " + enlace("precios de las actividades opcionales", url_exclusiva(clave, "extras"))
                          + " en nuestra sección exclusiva.", est["p"]))
    hist.append(Paragraph("Recomendaciones", est["h2"]))
    hist.append(lista(d["recomendaciones"]))

    # Caja final con los accesos exclusivos
    botones = Table([
        [Paragraph("<b>Contenido exclusivo para quien tiene este itinerario</b><br/>"
                   "Estas secciones del sitio solo se abren desde los enlaces de este PDF.", est["caja"])],
        [BotonEnlace("Ver precios por temporada", url_exclusiva(clave, "precios"))],
        [BotonEnlace("Actividades opcionales y código de descuento", url_exclusiva(clave, "extras"))],
        [BotonEnlace("Datos curiosos del destino", url_exclusiva(clave, "curiosidades"), color=AZUL)],
        [Paragraph("Al abrirlas te pediremos tu nombre, tu correo y por qué quieres viajar (solo la primera vez).<br/>"
                   "¿Listo para viajar? " + enlace("Pide tu cotización",
                   "mailto:contacto@viajeshorizonte.com?subject=" + quote(f"Cotización {d['consulta']}"))
                   + ".", est["caja"])],
    ], colWidths=[16.5 * cm])
    botones.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), ARENA),
        ("BOX", (0, 0), (-1, -1), 1.2, CORAL),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, 0), 14), ("BOTTOMPADDING", (0, -1), (-1, -1), 14),
    ]))
    hist += [Spacer(1, 18), KeepTogether(botones)]

    def pie(canv, doc):
        canv.saveState()
        canv.setFont("Helvetica", 8.5)
        canv.setFillColor(GRIS)
        canv.drawString(2 * cm, 1.2 * cm, f"Viajes Horizonte  |  {d['titulo']}  |  contacto@viajeshorizonte.com")
        canv.drawRightString(letter[0] - 2 * cm, 1.2 * cm, f"Página {doc.page}")
        canv.restoreState()

    doc = SimpleDocTemplate(str(salida), pagesize=letter, leftMargin=2 * cm, rightMargin=2 * cm,
                            topMargin=1.8 * cm, bottomMargin=2 * cm,
                            title=f"Itinerario – {d['titulo']}", author="Viajes Horizonte")
    doc.build(hist, onFirstPage=pie, onLaterPages=pie)
    return salida


# ---------------------------------------------------------------- Word ---------

def documento_base(titulo, subtitulo):
    doc = Document()
    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(11)
    marca = doc.add_paragraph()
    r = marca.add_run("VIAJES HORIZONTE")
    r.bold, r.font.size, r.font.color.rgb = True, Pt(9), RGBColor(0xF2, 0x6B, 0x4F)
    h = doc.add_heading(titulo, level=0)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x0B, 0x4F, 0x6C)
    p = doc.add_paragraph(subtitulo)
    p.runs[0].font.color.rgb = RGBColor(0x5A, 0x6B, 0x78)
    return doc


def generar_itinerario_docx(clave):
    d = DESTINOS[clave]
    doc = documento_base(f"Itinerario – {d['titulo']}", f"{d['duracion']} · {d['paquete']} · Versión editable")
    doc.add_heading("Sobre el destino", level=1)
    for p in d["sobre"]:
        doc.add_paragraph(p)
    doc.add_heading("Datos útiles", level=1)
    t = doc.add_table(rows=0, cols=2)
    t.style = "Light Grid Accent 1"
    for k, v in d["datos"]:
        fila = t.add_row().cells
        fila[0].text, fila[1].text = k, v
        fila[0].paragraphs[0].runs[0].bold = True
    doc.add_heading("Itinerario día a día", level=1)
    for i, (titulo, texto) in enumerate(d["dias"], 1):
        p = doc.add_paragraph()
        p.add_run(f"Día {i} – {titulo}. ").bold = True
        p.add_run(texto)
        doc.add_paragraph("Mis notas: ______________________________________________")
    for titulo, items in (("El paquete incluye", d["incluye"]), ("No incluye", d["no_incluye"]),
                          ("Recomendaciones", d["recomendaciones"])):
        doc.add_heading(titulo, level=1)
        for it in items:
            doc.add_paragraph(it, style="List Bullet")
    nota = doc.add_paragraph()
    nota.add_run("Precios, promociones y códigos de descuento: consulta los enlaces exclusivos del itinerario en PDF.").italic = True
    salida = RAIZ / "documentos" / f"itinerario-{clave}.docx"
    doc.save(salida)
    return salida


def generar_lista_equipaje():
    doc = documento_base("Lista de equipaje", "Marca cada casilla conforme vayas empacando.")
    secciones = {
        "Documentos": ["Pasaporte / identificación oficial", "Boletos de avión (impresos o en el celular)",
                       "Confirmación del hotel", "Póliza del seguro de viajero", "Tarjetas y algo de efectivo",
                       "Copias digitales de tus documentos"],
        "Ropa": ["Ropa interior y calcetines (1 por día + 1)", "Playeras / blusas", "Pantalones o shorts",
                 "Una muda elegante para cenar", "Pijama", "Chamarra o suéter ligero", "Calzado cómodo para caminar"],
        "Solo para playa (Cancún)": ["Traje de baño (2)", "Bloqueador biodegradable", "Sombrero y lentes de sol",
                                     "Sandalias", "Bolsa seca para el celular", "Repelente de insectos"],
        "Solo para ciudad (París)": ["Paraguas compacto", "Bufanda (otoño/invierno)", "Mochila antirrobo",
                                     "Adaptador de corriente tipo C/E"],
        "Salud e higiene": ["Medicamentos personales con receta", "Analgésico y antidiarreico", "Curitas",
                            "Cepillo y pasta de dientes", "Desodorante y artículos personales (envases de 100 ml)"],
        "Tecnología": ["Celular y cargador", "Batería externa", "Audífonos", "Cámara y memoria"],
    }
    for titulo, items in secciones.items():
        doc.add_heading(titulo, level=1)
        for it in items:
            doc.add_paragraph(f"☐  {it}")
    salida = RAIZ / "documentos" / "lista-de-equipaje.docx"
    doc.save(salida)
    return salida


def generar_requisitos():
    doc = documento_base("Requisitos y documentos de viaje",
                         "Guía general. Verifica siempre la información oficial antes de viajar.")
    doc.add_heading("Viajes nacionales (Cancún, desde México)", level=1)
    for t in ["Identificación oficial vigente (INE o pasaporte).",
              "Menores: acta de nacimiento o pasaporte; si viajan sin ambos padres, carta de autorización.",
              "Pago del Visitax de Quintana Roo (impuesto al visitante) antes de salir del estado."]:
        doc.add_paragraph(t, style="List Bullet")
    doc.add_heading("Viajes a Europa (París)", level=1)
    for t in ["Pasaporte con vigencia mínima de 3 meses después de la fecha de regreso (recomendado: 6 meses).",
              "Los mexicanos no necesitan visa para estancias turísticas de hasta 90 días en el espacio Schengen.",
              "Autorización ETIAS: consulta en el sitio oficial de la Unión Europea si ya es obligatoria en tus fechas.",
              "Seguro médico de viaje con cobertura mínima de 30,000 euros.",
              "Comprobante de hospedaje, boleto de regreso y solvencia económica (pueden pedirlos en migración)."]:
        doc.add_paragraph(t, style="List Bullet")
    doc.add_heading("Antes de salir", level=1)
    t = doc.add_table(rows=1, cols=3)
    t.style = "Light Grid Accent 1"
    for c, txt in zip(t.rows[0].cells, ["Trámite", "Cuándo", "Listo"]):
        c.text = txt
    for fila in [("Revisar vigencia del pasaporte", "3 meses antes"), ("Contratar seguro de viajero", "1 mes antes"),
                 ("Avisar al banco que viajarás", "1 semana antes"), ("Check-in en línea", "24–48 h antes")]:
        celdas = t.add_row().cells
        celdas[0].text, celdas[1].text, celdas[2].text = fila[0], fila[1], "☐"
    p = doc.add_paragraph()
    p.add_run("\nFuentes oficiales: Secretaría de Relaciones Exteriores (sre.gob.mx) y "
              "travel-europe.europa.eu.").italic = True
    salida = RAIZ / "documentos" / "requisitos-de-viaje.docx"
    doc.save(salida)
    return salida


# ---------------------------------------------------------------- Excel --------

def generar_presupuesto():
    wb = Workbook()
    ws = wb.active
    ws.title = "Presupuesto"
    azul = PatternFill("solid", fgColor="0B4F6C")
    arena = PatternFill("solid", fgColor="FBF5EC")
    borde = Border(bottom=Side(style="thin", color="E4DDD2"))
    moneda = '"$"#,##0.00'

    ws["A1"] = "Planificador de presupuesto de viaje – Viajes Horizonte"
    ws["A1"].font = Font(size=14, bold=True, color="0B4F6C")
    ws["A3"], ws["B3"] = "Destino:", "Escribe tu destino"
    ws["A4"], ws["B4"] = "Número de viajeros:", 2
    ws["A5"], ws["B5"] = "Días de viaje:", 7
    for c in ("A3", "A4", "A5"):
        ws[c].font = Font(bold=True)
    for c in ("B3", "B4", "B5"):
        ws[c].fill = arena

    encabezados = ["Categoría", "Concepto", "Costo unitario", "Cantidad", "Subtotal", "Pagado", "Pendiente"]
    fila_enc = 7
    for i, h in enumerate(encabezados, 1):
        c = ws.cell(row=fila_enc, column=i, value=h)
        c.font, c.fill, c.alignment = Font(bold=True, color="FFFFFF"), azul, Alignment(horizontal="center")

    conceptos = [
        ("Transporte", "Vuelo redondo (por persona)", 0, "=B4"),
        ("Transporte", "Traslados / taxis", 0, 1),
        ("Hospedaje", "Hotel (por noche)", 0, "=B5-1"),
        ("Comida", "Comidas por día (todo el grupo)", 0, "=B5"),
        ("Actividades", "Excursiones y entradas", 0, 1),
        ("Seguro", "Seguro de viajero (por persona)", 0, "=B4"),
        ("Compras", "Recuerdos y compras", 0, 1),
        ("Otros", "Propinas e imprevistos", 0, 1),
    ]
    primera = fila_enc + 1
    for i, (cat, con, costo, cant) in enumerate(conceptos):
        r = primera + i
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2, value=con)
        ws.cell(row=r, column=3, value=costo).number_format = moneda
        ws.cell(row=r, column=4, value=cant)
        ws.cell(row=r, column=5, value=f"=C{r}*D{r}").number_format = moneda
        ws.cell(row=r, column=6, value=0).number_format = moneda
        ws.cell(row=r, column=7, value=f"=E{r}-F{r}").number_format = moneda
        for col in (3, 4, 6):
            ws.cell(row=r, column=col).fill = arena
        for col in range(1, 8):
            ws.cell(row=r, column=col).border = borde
    ultima = primera + len(conceptos) - 1

    t = ultima + 2
    resumen = [("Total del viaje", f"=SUM(E{primera}:E{ultima})"),
               ("Total pagado", f"=SUM(F{primera}:F{ultima})"),
               ("Pendiente por pagar", f"=SUM(G{primera}:G{ultima})"),
               ("Costo por persona", f"=IFERROR(E{t}/B4,0)"),
               ("Costo por día", f"=IFERROR(E{t}/B5,0)")]
    for i, (k, f) in enumerate(resumen):
        ws.cell(row=t + i, column=4, value=k).font = Font(bold=True)
        c = ws.cell(row=t + i, column=5, value=f)
        c.number_format, c.font = moneda, Font(bold=True, color="F26B4F")

    ws.cell(row=t + 6, column=1, value="Escribe solo en las celdas color arena; los totales se calculan solos.").font = Font(italic=True, color="5A6B78")
    for col, ancho in zip("ABCDEFG", (16, 34, 16, 20, 16, 14, 14)):
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = ws.cell(row=primera, column=1)

    salida = RAIZ / "documentos" / "planificador-de-presupuesto.xlsx"
    wb.save(salida)
    return salida


# ---------------------------------------------------------------- TXT ----------

def generar_frases():
    texto = """FRASES ÚTILES EN FRANCÉS – VIAJES HORIZONTE
==========================================

Saludos
  Bonjour (bon-YUR) ............ Buenos días / Hola
  Bonsoir (bon-SUAR) ........... Buenas noches (al llegar)
  Au revoir (o-re-VUAR) ........ Adiós
  Merci (mer-SÍ) ............... Gracias
  S'il vous plaît (sil-vu-PLÉ) . Por favor
  Excusez-moi (ex-kü-SÉ-mua) ... Disculpe

En el restaurante
  Une table pour deux, s'il vous plaît ... Una mesa para dos, por favor
  L'addition, s'il vous plaît ............ La cuenta, por favor
  Une carafe d'eau ....................... Una jarra de agua (gratis)
  Je suis allergique à ... ............... Soy alérgico/a a ...

Para moverte
  Où est le métro ? ...................... ¿Dónde está el metro?
  C'est combien ? ........................ ¿Cuánto cuesta?
  Je ne comprends pas .................... No entiendo
  Parlez-vous espagnol / anglais ? ....... ¿Habla español / inglés?

Emergencias
  Au secours ! ........................... ¡Auxilio!
  Appelez un médecin ..................... Llame a un médico
  Número de emergencias en Europa: 112

Consejo: saluda siempre con "Bonjour" al entrar a una tienda o restaurante.
"""
    salida = RAIZ / "documentos" / "frases-utiles-frances.txt"
    salida.write_text(texto, encoding="utf-8-sig")
    return salida


if __name__ == "__main__":
    (RAIZ / "pdfs").mkdir(exist_ok=True)
    (RAIZ / "documentos").mkdir(exist_ok=True)
    hechos = [generar_pdf("cancun"), generar_pdf("paris"),
              generar_itinerario_docx("cancun"), generar_itinerario_docx("paris"),
              generar_lista_equipaje(), generar_requisitos(), generar_presupuesto(), generar_frases()]
    for f in hechos:
        print("OK", f.relative_to(RAIZ))
