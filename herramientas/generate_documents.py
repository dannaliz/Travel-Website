"""Generates the English documents that hold the site's valuable content.

The HTML pages no longer contain prices, discount codes or detailed
itineraries: that information lives ONLY in these files.

    Step 3 (destinations.html):
        files/itinerary-cancun.pdf, files/itinerary-paris.pdf   (day-by-day)
        files/itinerary-cancun.docx, files/itinerary-paris.docx (editable)
    Step 4 (resources.html):
        files/prices-cancun.pdf, files/prices-paris.pdf         (prices + code)
        files/price-list-2026.xlsx                              (all prices)
        files/insider-guide.docx                                (codes + fun facts)

Usage (from the project folder):
    pip install reportlab python-docx openpyxl
    python herramientas/generate_documents.py
"""
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (ListFlowable, ListItem, Paragraph,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "files"
BLUE = colors.HexColor("#0b4f6c")
CORAL = colors.HexColor("#f26b4f")
GREY = colors.HexColor("#5a6b78")
SAND = colors.HexColor("#fbf5ec")

DESTINATIONS = {
    "cancun": {
        "title": "Cancún, Mexico",
        "length": "7 days / 6 nights",
        "package": "Flight + all-inclusive hotel",
        "about": "White-sand beaches and turquoise water in the Mexican Caribbean, combined with visits "
                 "to Isla Mujeres, the Mayan city of Chichén Itzá and the seaside ruins of Tulum.",
        "facts": [("Currency", "Mexican peso (MXN)"), ("Language", "Spanish; English in tourist areas"),
                  ("Weather", "26–33 °C; rainy season June to October"), ("Time zone", "UTC−5 all year"),
                  ("Documents", "Passport or official ID")],
        "days": [("Arrival", "Transfer to the hotel, beach afternoon and welcome dinner."),
                 ("Rest day", "Free day to enjoy the all-inclusive resort."),
                 ("Isla Mujeres", "Catamaran, snorkeling and an afternoon at Playa Norte."),
                 ("Chichén Itzá", "Archaeological site, Ik Kil cenote and the town of Valladolid."),
                 ("Tulum & Playa del Carmen", "Ruins by the sea and a walk along Fifth Avenue."),
                 ("Free day", "Optional activities and farewell dinner."),
                 ("Return", "Breakfast and transfer to the airport.")],
        "included": ["Round-trip economy flight", "6 nights all-inclusive", "Airport transfers",
                     "Guided tours to Isla Mujeres, Chichén Itzá and Tulum", "Travel insurance"],
        "not_included": ["Environmental tax (Visitax)", "Optional activities", "Tips and personal expenses"],
        "tips": ["Use biodegradable sunscreen in cenotes and reefs.", "Bring light clothes, a hat and comfortable shoes.",
                 "Carry cash in pesos for tips and markets.", "Stay hydrated: heat and humidity are intense."],
        "price_note": "Price per person in MXN, departing from Mexico City.",
        "seasons": [("Low", "May – June, September – November", 18900, 26500, 9900),
                    ("Mid", "January – April (except Holy Week)", 22400, 31200, 11500),
                    ("High", "July – August, Holy Week, December", 27800, 38900, 13900)],
        "activities": [("Swim with whale sharks (June – September)", "Full day", 3200),
                       ("Xcaret Plus with transportation", "Full day", 3650),
                       ("Snorkeling at the Underwater Museum (MUSA)", "3 hours", 1350),
                       ("Romantic dinner on the beach", "2 hours", 2400),
                       ("Cenote route by ATV", "5 hours", 1890)],
        "code": "HORIZONTE-CUN10",
        "code_text": "10% off the package.",
        "fun": [("A young city", "Cancún was officially founded in 1970 as a planned tourism project."),
                ("The descent of Kukulkán", "On the equinoxes, sunlight forms a serpent's shadow on the pyramid at Chichén Itzá."),
                ("A museum under the sea", "MUSA has more than 500 submerged sculptures that work as an artificial reef."),
                ("Underground rivers", "Yucatán has no surface rivers: water flows underground and emerges in cenotes."),
                ("Turtle beaches", "From May to October sea turtles nest on the beaches."),
                ("Its own time zone", "Quintana Roo uses UTC−5 all year, one hour ahead of Mexico City.")],
    },
    "paris": {
        "title": "Paris, France",
        "length": "10 days / 9 nights",
        "package": "Flight + hotel + tours",
        "about": "The City of Light: art, architecture, food and fashion. Walk through Montmartre and "
                 "Le Marais, enjoy a café by the Seine and visit the Eiffel Tower, the Louvre and Versailles.",
        "facts": [("Currency", "Euro (EUR)"), ("Language", "French; English in tourist areas"),
                  ("Weather", "Summer 20–26 °C; winter 3–8 °C"), ("Time zone", "UTC+1 (UTC+2 in summer)"),
                  ("Documents", "Valid passport; Schengen area requirements")],
        "days": [("Arrival in Paris", "Charles de Gaulle airport, transfer to the hotel and free time."),
                 ("Eiffel Tower & Seine cruise", "Second floor of the tower, Champ de Mars and a sunset cruise."),
                 ("The Louvre", "Guided visit: Mona Lisa and Venus de Milo; afternoon in the Tuileries."),
                 ("Île de la Cité & Latin Quarter", "Notre-Dame, Sainte-Chapelle and historic cafés."),
                 ("Palace of Versailles", "Full-day trip: Hall of Mirrors, royal apartments and gardens."),
                 ("Montmartre", "Sacré-Cœur and Place du Tertre; optional cabaret night."),
                 ("Champs-Élysées & Arc de Triomphe", "Walk, climb the Arc and shopping."),
                 ("Musée d'Orsay & Le Marais", "The world's largest Impressionist collection."),
                 ("Free day", "Optional Disneyland Paris or Loire Valley; farewell dinner."),
                 ("Return", "Breakfast, check-out and airport transfer.")],
        "included": ["Round-trip economy flight", "9 nights with breakfast", "Airport transfers",
                     "Tickets and guided visits: Eiffel Tower, Louvre, Versailles, Orsay", "Seine cruise",
                     "Travel insurance valid in Europe"],
        "not_included": ["Meals not specified", "City tourist tax (paid at the hotel)", "Optional activities"],
        "tips": ["Buy a transport pass to use the metro cheaply.", "Watch your belongings in crowded places.",
                 "Wear comfortable shoes: Paris is best on foot.", "Say \"Bonjour\" when entering shops and restaurants."],
        "price_note": "Price per person in MXN, departing from Mexico City.",
        "seasons": [("Low", "November – March (except holidays)", 46500, 61900, 31800),
                    ("Mid", "April – May, September – October", 52900, 70400, 35600),
                    ("High", "June – August, December", 61300, 82700, 41200)],
        "activities": [("Disneyland Paris (1 day, 2 parks)", "Full day", 3900),
                       ("Loire Valley castles", "Full day", 4600),
                       ("Moulin Rouge show with champagne", "2 hours", 3300),
                       ("French cooking class (croissants & macarons)", "3 hours", 2450),
                       ("Dinner cruise on the Seine", "2.5 hours", 2900)],
        "code": "HORIZONTE-PAR15",
        "code_text": "15% off optional activities.",
        "fun": [("A temporary tower", "The Eiffel Tower was built for the 1889 World's Fair and was meant to last 20 years."),
                ("It grows in summer", "In the heat, the iron expands and the tower can grow about 15 cm."),
                ("Impossible to see it all", "The Louvre displays around 35,000 works."),
                ("A city underground", "The Catacombs hold the remains of about six million people."),
                ("Lady Liberty's sister", "A smaller replica of the Statue of Liberty stands on the Île aux Cygnes."),
                ("The baguette is heritage", "In 2022 UNESCO recognized baguette-making as Intangible Cultural Heritage.")],
    },
}

money = lambda n: f"${n:,} MXN"


# ------------------------------------------------------------------ PDF helpers

def styles():
    b = getSampleStyleSheet()
    return {
        "brand": ParagraphStyle("brand", parent=b["Normal"], fontName="Helvetica-Bold", fontSize=9, textColor=CORAL),
        "title": ParagraphStyle("title", parent=b["Title"], fontSize=24, leading=28, textColor=BLUE, alignment=0, spaceAfter=4),
        "sub": ParagraphStyle("sub", parent=b["Normal"], fontSize=12, textColor=GREY, spaceAfter=12),
        "h2": ParagraphStyle("h2", parent=b["Heading2"], fontSize=15, textColor=BLUE, spaceBefore=14, spaceAfter=6),
        "p": ParagraphStyle("p", parent=b["Normal"], fontSize=10.5, leading=15),
        "code": ParagraphStyle("code", parent=b["Normal"], fontName="Helvetica-Bold", fontSize=20, leading=26,
                               textColor=CORAL, alignment=1),
        "center": ParagraphStyle("center", parent=b["Normal"], fontSize=10.5, leading=15, alignment=1),
    }


def table(rows, widths, st, header=True):
    t = Table([[Paragraph(str(c), st["p"]) for c in r] for r in rows], colWidths=widths)
    style = [("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.HexColor("#e4ddd2")),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]
    if header:
        style += [("BACKGROUND", (0, 0), (-1, 0), BLUE), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white)]
        t = Table([[Paragraph(f'<font color="white"><b>{c}</b></font>', st["p"]) for c in rows[0]]]
                  + [[Paragraph(str(c), st["p"]) for c in r] for r in rows[1:]], colWidths=widths)
    else:
        style += [("BACKGROUND", (0, 0), (0, -1), SAND)]
    t.setStyle(TableStyle(style))
    return t


def bullets(items, st):
    return ListFlowable([ListItem(Paragraph(i, st["p"]), leftIndent=12) for i in items],
                        bulletType="bullet", start="•", leftIndent=12)


def build_pdf(path, title, story, footer):
    def foot(canv, doc):
        canv.saveState()
        canv.setFont("Helvetica", 8.5)
        canv.setFillColor(GREY)
        canv.drawString(2 * cm, 1.2 * cm, f"Horizonte Travel  |  {footer}  |  Confidential – registered travelers only")
        canv.drawRightString(letter[0] - 2 * cm, 1.2 * cm, f"Page {doc.page}")
        canv.restoreState()
    SimpleDocTemplate(str(path), pagesize=letter, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=1.8 * cm,
                      bottomMargin=2 * cm, title=title, author="Horizonte Travel").build(story, onFirstPage=foot, onLaterPages=foot)


# ------------------------------------------------------------------ documents

def itinerary_pdf(key):
    d, st = DESTINATIONS[key], styles()
    s = [Paragraph("HORIZONTE TRAVEL · ITINERARY", st["brand"]), Paragraph(d["title"], st["title"]),
         Paragraph(f'{d["length"]} · {d["package"]}', st["sub"]), Paragraph(d["about"], st["p"]),
         Paragraph("Quick facts", st["h2"]), table(d["facts"], [4 * cm, 12.5 * cm], st, header=False),
         Paragraph("Day by day", st["h2"])]
    for i, (t, txt) in enumerate(d["days"], 1):
        s += [Paragraph(f'<font color="#0b4f6c"><b>Day {i} – {t}.</b></font> {txt}', st["p"]), Spacer(1, 4)]
    s += [Paragraph("Included", st["h2"]), bullets(d["included"], st),
          Paragraph("Not included", st["h2"]), bullets(d["not_included"], st),
          Paragraph("Recommendations", st["h2"]), bullets(d["tips"], st), Spacer(1, 14),
          Paragraph("Prices and discount codes are in the price documents unlocked in Step 4.", st["center"])]
    build_pdf(OUT / f"itinerary-{key}.pdf", f"Itinerary – {d['title']}", s, d["title"])


def prices_pdf(key):
    d, st = DESTINATIONS[key], styles()
    s = [Paragraph("HORIZONTE TRAVEL · PRICES 2026–2027", st["brand"]), Paragraph(f'{d["title"]}: prices', st["title"]),
         Paragraph(f'{d["length"]} · {d["package"]}', st["sub"]),
         Paragraph("Package prices", st["h2"]), Paragraph(d["price_note"], st["p"]), Spacer(1, 6),
         table([("Season", "Dates", "Double room", "Single room", "Child (2–11)")]
               + [(a, b, money(c), money(e), money(f)) for a, b, c, e, f in d["seasons"]],
               [2 * cm, 5.5 * cm, 3 * cm, 3 * cm, 3 * cm], st),
         Paragraph("Optional activities", st["h2"]), Paragraph("Price per person. Book them with your advisor.", st["p"]),
         Spacer(1, 6),
         table([("Activity", "Duration", "Price")] + [(a, b, money(c)) for a, b, c in d["activities"]],
               [9.5 * cm, 3 * cm, 4 * cm], st),
         Spacer(1, 20)]
    box = Table([[Paragraph("Your discount code", st["center"])], [Paragraph(d["code"], st["code"])],
                 [Paragraph(d["code_text"] + " Mention it when you request your quote at contacto@viajeshorizonte.com.", st["center"])]],
                colWidths=[14 * cm])
    box.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 1.5, CORAL), ("BACKGROUND", (0, 0), (-1, -1), SAND),
                             ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    s += [box, Spacer(1, 10), Paragraph("Prices subject to availability and change without notice.", st["center"])]
    build_pdf(OUT / f"prices-{key}.pdf", f"Prices – {d['title']}", s, d["title"])


def heading_doc(title, subtitle):
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    r = doc.add_paragraph().add_run("HORIZONTE TRAVEL")
    r.bold, r.font.size, r.font.color.rgb = True, Pt(9), RGBColor(0xF2, 0x6B, 0x4F)
    for run in doc.add_heading(title, level=0).runs:
        run.font.color.rgb = RGBColor(0x0B, 0x4F, 0x6C)
    doc.add_paragraph(subtitle).runs[0].font.color.rgb = RGBColor(0x5A, 0x6B, 0x78)
    return doc


def itinerary_docx(key):
    d = DESTINATIONS[key]
    doc = heading_doc(f"Itinerary – {d['title']}", f"{d['length']} · {d['package']} · editable version")
    doc.add_paragraph(d["about"])
    for i, (t, txt) in enumerate(d["days"], 1):
        doc.add_heading(f"Day {i} – {t}", level=2)
        doc.add_paragraph(txt)
        doc.add_paragraph("My notes: ____________________________________________")
    doc.add_heading("Included", level=1)
    for x in d["included"]:
        doc.add_paragraph(x, style="List Bullet")
    doc.save(OUT / f"itinerary-{key}.docx")


def insider_docx():
    doc = heading_doc("Insider guide", "Discount codes and fun facts for registered travelers")
    for d in DESTINATIONS.values():
        doc.add_heading(d["title"], level=1)
        p = doc.add_paragraph("Discount code: ")
        r = p.add_run(d["code"])
        r.bold, r.font.size, r.font.color.rgb = True, Pt(14), RGBColor(0xF2, 0x6B, 0x4F)
        doc.add_paragraph(d["code_text"])
        doc.add_heading("Fun facts", level=2)
        for t, txt in d["fun"]:
            para = doc.add_paragraph(style="List Bullet")
            para.add_run(t + ": ").bold = True
            para.add_run(txt)
    doc.save(OUT / "insider-guide.docx")


def price_list_xlsx():
    wb = Workbook()
    head = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="0B4F6C")
    sheets = [("Packages", ["Destination", "Season", "Dates", "Double room", "Single room", "Child (2–11)"],
               [(d["title"], *row) for d in DESTINATIONS.values() for row in d["seasons"]]),
              ("Optional activities", ["Destination", "Activity", "Duration", "Price"],
               [(d["title"], *row) for d in DESTINATIONS.values() for row in d["activities"]])]
    for i, (name, cols, rows) in enumerate(sheets):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = name
        ws.append(cols)
        for c in ws[1]:
            c.font, c.fill, c.alignment = head, fill, Alignment(horizontal="center")
        for r in rows:
            ws.append(list(r))
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, int):
                    c.number_format = '"$"#,##0" MXN"'
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(len(str(c.value)) for c in col) + 4
        ws.freeze_panes = "A2"
    wb.save(OUT / "price-list-2026.xlsx")


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    for k in DESTINATIONS:
        itinerary_pdf(k)
        itinerary_docx(k)
        prices_pdf(k)
    insider_docx()
    price_list_xlsx()
    for f in sorted(OUT.iterdir()):
        print(f.name, f.stat().st_size)
