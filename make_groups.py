"""Make Class X IT project-file printing groups (section-wise, then gender-wise,
3 or 4 students per group), give the 4 projects to the groups in turn
(1 2 3 4 1 2 3 4 ...) and write a PDF with a link and QR code for each project.

Usage:  python make_groups.py students.xlsx
"""
import sys
import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, KeepTogether)
from reportlab.graphics.barcode.qr import QrCodeWidget
from reportlab.graphics.shapes import Drawing


def qr_code(url, size=32 * mm):
    """Return a QR code drawing for url, size x size."""
    w = QrCodeWidget(url, barLevel="M")
    x1, y1, x2, y2 = w.getBounds()
    d = Drawing(size, size, transform=[size / (x2 - x1), 0, 0,
                                       size / (y2 - y1), 0, 0])
    d.add(w)
    return d

XLSX = sys.argv[1] if len(sys.argv) > 1 else "students_class_X.xlsx"
OUT = "Class_X_IT_Project_Groups.pdf"
BASE = "https://github.com/pbhil95/X-IT-Project-File/raw/main/"
PROJECTS = [  # (name, link, colour)
    ("Library Management System", BASE + "IT-Project-File.pdf", "#2E7D32"),
    ("Student Result Management System", BASE + "More-Projects/"
     "Student-Result-Management/IT-Project-Student-Result-Management.pdf",
     "#5B3FA0"),
    ("Hospital Management System", BASE + "More-Projects/Hospital-Management/"
     "IT-Project-Hospital-Management.pdf", "#C62828"),
    ("Inventory and Sales Management System", BASE + "More-Projects/"
     "Inventory-Sales-Management/IT-Project-Inventory-Sales-Management.pdf",
     "#D35400"),
]


def split_sizes(n):
    """Split n students into groups of 4 and 3 only (fewest groups)."""
    if n <= 4:
        return [n]
    groups = -(-n // 4)                     # ceil(n / 4)
    base, extra = divmod(n, groups)
    return [base + 1] * extra + [base] * (groups - extra)


# ---- read IT students ----------------------------------------------------
ws = openpyxl.load_workbook(XLSX).active
students = []
for row in ws.iter_rows(min_row=5, values_only=True):
    sno, name, gender, section, subject = row[:5]
    if not name or subject != "Information Technology":
        continue
    gender = "Boys" if str(gender).strip().upper().startswith("M") else "Girls"
    students.append((name.strip(), gender, section.strip()))

# ---- make groups ---------------------------------------------------------
groups = []                                # (group no, section, gender, names)
for section in sorted({s[2] for s in students}):
    for gender in ("Boys", "Girls"):
        names = sorted(n for n, g, s in students if s == section and g == gender)
        i = 0
        for size in split_sizes(len(names)):
            groups.append((len(groups) + 1, section, gender, names[i:i + size]))
            i += size


def project_of(group_no):
    """Projects go round in turn: G-01 -> 1, G-02 -> 2, ... G-05 -> 1."""
    return (group_no - 1) % len(PROJECTS)

# ---- build PDF -----------------------------------------------------------
title = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=17, leading=22,
                       alignment=1, textColor=colors.HexColor("#1F4E79"),
                       spaceAfter=4)
sub = ParagraphStyle("s", fontName="Helvetica", fontSize=11, alignment=1,
                     leading=15)
head = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=13,
                      textColor=colors.white)
cell = ParagraphStyle("c", fontName="Helvetica", fontSize=10.5, leading=14)
link_style = ParagraphStyle("l", parent=sub, fontSize=11.5)

story = [
    Paragraph("PM SHRI JNV TARIKHET, ALMORA", title),
    Paragraph("<b>Class X - Information Technology (402) - Session 2026-27</b>", sub),
    Paragraph("Project File Printing Groups", sub),
    Spacer(1, 3 * mm),
]
card_t = ParagraphStyle("ct", parent=cell, fontName="Helvetica-Bold",
                        fontSize=10.5, leading=13, textColor=colors.white)
card_s = ParagraphStyle("cs", parent=cell, fontSize=9.5, leading=12)
cards = []
def tint(hexcol, f=0.88):
    r, g, b = (int(hexcol[i:i + 2], 16) for i in (1, 3, 5))
    return colors.Color(*(v / 255 + (1 - v / 255) * f for v in (r, g, b)))


for k, (name, link, col) in enumerate(PROJECTS):
    card_t = ParagraphStyle("ct", parent=cell, fontName="Helvetica-Bold",
                            fontSize=10.5, leading=13,
                            textColor=colors.HexColor(col))
    gl = ", ".join(f"G-{g[0]:02d}" for g in groups if project_of(g[0]) == k)
    text = [Paragraph(f"Project {k + 1}", card_t),
            Paragraph(name, card_t), Spacer(1, 2 * mm),
            Paragraph(f'<link href="{link}" color="blue"><u>Click here to '
                      'open / download</u></link>', card_s),
            Paragraph("or scan the QR code", card_s), Spacer(1, 1.5 * mm),
            Paragraph(f"<b>Groups:</b> {gl}", card_s)]
    c = Table([[text, qr_code(link, 26 * mm)]], colWidths=[56 * mm, 28 * mm])
    c.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 1, colors.HexColor(col)),
        ("BACKGROUND", (0, 0), (-1, -1), tint(col)),
        ("LINEBEFORE", (0, 0), (0, 0), 5, colors.HexColor(col)),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]))
    cards.append(c)
grid = Table([[cards[0], cards[1]], [cards[2], cards[3]]],
             colWidths=[88.5 * mm, 88.5 * mm])
grid.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                          ("LEFTPADDING", (0, 0), (-1, -1), 2),
                          ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                          ("TOPPADDING", (0, 0), (-1, -1), 2),
                          ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
story.append(grid)
story += [
    Spacer(1, 2 * mm),
    Paragraph(f"Total IT students: <b>{len(students)}</b> &nbsp;|&nbsp; "
              f"Total groups: <b>{len(groups)}</b>", sub),
    Spacer(1, 3 * mm),
]

# ---- printing instructions -----------------------------------------------
note = ParagraphStyle("n", parent=cell, fontSize=10.5, leading=15)
steps = [
    "The 4 projects are given to the groups in turn (1, 2, 3, 4, 1, 2, ...). "
    "Find your group below and print <b>your group's project</b> only.",
    "Each group prints <b>one</b> copy of its project file and shares it.",
    '<font color="#C00000"><b>The front page (cover page) must be printed in '
    'COLOUR.</b></font> All other pages can be printed in black &amp; white.',
    "Write the names and CBSE Roll Nos. of <b>all group members</b> and the "
    "Section on the front page, acknowledgement and certificate.",
]
rows = [[Paragraph("<b>Printing Instructions</b>", head)]]
rows += [[Paragraph(f"{i}. {t}", note)] for i, t in enumerate(steps, 1)]
ins = Table(rows, colWidths=[177 * mm])
ins.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#C55A11")),
    ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FFF4E5")),
    ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#C55A11")),
    ("TOPPADDING", (0, 0), (-1, -1), 3),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
]))
story += [ins, Spacer(1, 5 * mm)]

BLUE, PINK = colors.HexColor("#2E75B6"), colors.HexColor("#C0507A")
for section in sorted({g[1] for g in groups}):
    for gender in ("Boys", "Girls"):
        sel = [g for g in groups if g[1] == section and g[2] == gender]
        if not sel:
            continue
        colour = BLUE if gender == "Boys" else PINK
        n = sum(len(g[3]) for g in sel)
        data = [[Paragraph(f"Section {section} - {gender} ({n} students)", head),
                 "", ""],
                ["Group", "Students", "Project"]]
        pcols = []
        for no, _, _, names in sel:
            members = "<br/>".join(f"{k}. {nm}" for k, nm in enumerate(names, 1))
            k = project_of(no)
            pname, plink, pcol = PROJECTS[k]
            data.append([f"G-{no:02d}", Paragraph(members, cell),
                         Paragraph(f'<font color="{pcol}"><b>Project {k + 1}</b>'
                                   f'</font><br/><link href="{plink}" '
                                   f'color="{pcol}"><u>{pname}</u></link>', cell)])
            pcols.append(pcol)
        t = Table(data, colWidths=[22 * mm, 85 * mm, 70 * mm])
        t.setStyle(TableStyle([
            ("SPAN", (0, 0), (-1, 0)),
            ("BACKGROUND", (0, 0), (-1, 0), colour),
            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#E7E6E6")),
            ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
            ("FONTNAME", (0, 2), (0, -1), "Helvetica-Bold"),
            ("ALIGN", (0, 1), (0, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.6, colors.grey),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ] + [("LINEBEFORE", (2, r + 2), (2, r + 2), 3, colors.HexColor(c))
             for r, c in enumerate(pcols)]))
        story += [KeepTogether(t), Spacer(1, 6 * mm)]

story.append(Paragraph("Pankaj Bhilingwal<br/>Internal Examiner (IT)",
                       ParagraphStyle("sig", parent=cell, alignment=2)))

SimpleDocTemplate(OUT, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm,
                  topMargin=14 * mm, bottomMargin=14 * mm,
                  title="Class X IT Project Groups 2026-27").build(story)
print(f"{len(students)} IT students -> {len(groups)} groups -> {OUT}")
for no, sec, gen, names in groups:
    print(f"G-{no:02d} {sec} {gen:5} P{project_of(no) + 1}: {', '.join(names)}")
