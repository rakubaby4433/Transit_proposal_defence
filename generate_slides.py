"""
Proposal Defense Slide Generator
=================================
Intelligent Public Transit Route Recommendation System for Kathmandu Valley
Group: THA080BCT006 / 012 / 022 / 030  |  Supervisor: Er. Anup Shrestha

Builds a .pptx that follows the IOE Thapathali "Slide Design Guidelines"
(Dinesh Baniya Kshatri, 2020) and the 12-minute Proposal Defense order.

Design rules enforced:
  * White background, black text on every slide
  * No animations
  * Cover: Title Arial Bold 44 | Members Arial Bold 22 | Dept/Campus/Date Arial 20 | NO footer
  * Titles: Arial Bold, centered, 36-40 pt
  * Body: Arial, left-justified, 28 pt main bullet / 24 pt sub-bullet
  * <= 6 main bullets, <= 8 words per bullet, <= 2-3 sub-bullets
  * Footer (all non-cover slides): date bottom-left, slide number bottom-right
  * One figure / table per slide

USAGE:
  pip install -r requirements.txt
  python generate_slides.py
  -> produces Proposal_Defense.pptx

OPTIONAL IMAGES (drop in same folder, auto-embedded if present):
  block_diagram.png   flowchart.png   gantt.png
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------------
# CONSTANTS
# ----------------------------------------------------------------------------
FONT = "Arial"
BLACK = RGBColor(0x00, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY_LINE = RGBColor(0x00, 0x00, 0x00)
HEADER_FILL = RGBColor(0xD9, 0xD9, 0xD9)

PRESENTATION_DATE = "June 16, 2026"
OUTPUT_FILE = "Proposal_Defense.pptx"

# 4:3 canvas (10 x 7.5 in) keeps Arial sizing identical to the guideline deck.
SLIDE_W = Inches(10)
SLIDE_H = Inches(7.5)

# ----------------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------------

def _set_white_background(slide):
    """Force a solid white fill on the slide background."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = WHITE


def _blank_slide(prs):
    """Return a brand-new blank slide with white background."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 6 = truly blank
    _set_white_background(slide)
    return slide


def _add_title(slide, text, size=38):
    """Centered Arial Bold title, 36-40 pt."""
    size = max(36, min(40, size))
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(9.0), Inches(1.0))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    f = run.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = True
    f.color.rgb = BLACK
    return box


def _add_footer(slide, page_no):
    """Date bottom-left, slide number bottom-right (non-cover slides only)."""
    # Date (left)
    left = slide.shapes.add_textbox(Inches(0.3), Inches(7.05), Inches(4.0), Inches(0.35))
    p = left.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = PRESENTATION_DATE
    r.font.name = FONT
    r.font.size = Pt(12)
    r.font.color.rgb = BLACK
    # Slide number (right)
    right = slide.shapes.add_textbox(Inches(8.7), Inches(7.05), Inches(1.0), Inches(0.35))
    p = right.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = str(page_no)
    r.font.name = FONT
    r.font.size = Pt(12)
    r.font.color.rgb = BLACK


def _add_bullets(slide, items, top=Inches(1.6), height=Inches(5.0),
                 left=Inches(0.8), width=Inches(8.6)):
    """
    Add a left-justified bullet list.
    `items` is a list of (text, level) tuples.
      level 0 -> main bullet, Arial 28, solid disc
      level 1 -> sub-bullet,  Arial 24, dash
    """
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for text, level in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        p.level = level
        p.space_after = Pt(10)
        run = p.add_run()
        bullet_char = "\u2022  " if level == 0 else "\u2013  "
        run.text = bullet_char + text
        f = run.font
        f.name = FONT
        f.size = Pt(28) if level == 0 else Pt(24)
        f.bold = False
        f.color.rgb = BLACK
        if level == 1:
            # indent sub-bullets
            pPr = p._pPr if p._pPr is not None else p.get_or_add_pPr()
            pPr.set('marL', str(Emu(Inches(0.5)).emu if hasattr(Emu(Inches(0.5)),'emu') else 457200))
    return box


def _style_cell(cell, text, bold=False, size=24, align=PP_ALIGN.LEFT, fill=None):
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = WHITE
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = BLACK


def _maybe_image_or_placeholder(slide, filename, caption,
                                left=Inches(1.3), top=Inches(1.7),
                                width=Inches(7.4), height=Inches(4.6)):
    """Embed image if the file exists, else draw a labeled placeholder box."""
    if os.path.exists(filename):
        slide.shapes.add_picture(filename, left, top, width=width)
    else:
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = WHITE
        shape.line.color.rgb = BLACK
        shape.line.width = Pt(1.25)
        tf = shape.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = caption
        r.font.name = FONT
        r.font.size = Pt(24)
        r.font.bold = False
        r.font.color.rgb = BLACK


# ----------------------------------------------------------------------------
# BUILD PRESENTATION
# ----------------------------------------------------------------------------

def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    page = 0  # footer counter (starts after cover)

    # ---------------- SLIDE 1 : COVER (no footer) ----------------
    s = _blank_slide(prs)

    # Title (Arial Bold 44)
    box = s.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(9.0), Inches(1.8))
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "An Intelligent Public Transit Route Recommendation System for Kathmandu Valley"
    r.font.name = FONT; r.font.size = Pt(44); r.font.bold = True; r.font.color.rgb = BLACK

    # Members (Arial Bold 22)
    members = [
        "Arun Raj Neupane (THA080BCT006)",
        "Bipin Tharu (THA080BCT012)",
        "Naresh Rasaili (THA080BCT022)",
        "Rakesh Kumar Yadav (THA080BCT030)",
    ]
    mbox = s.shapes.add_textbox(Inches(0.5), Inches(3.0), Inches(9.0), Inches(2.0))
    mtf = mbox.text_frame; mtf.word_wrap = True
    for i, m in enumerate(members):
        p = mtf.paragraphs[0] if i == 0 else mtf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = m
        r.font.name = FONT; r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = BLACK

    # Supervisor + Dept + Campus + Date (Arial 20)
    info = [
        "Project Supervisor: Er. Anup Shrestha",
        "Department of Electronics and Computer Engineering",
        "Institute of Engineering, Thapathali Campus",
        "Kathmandu, Nepal",
        PRESENTATION_DATE,
    ]
    ibox = s.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(9.0), Inches(2.0))
    itf = ibox.text_frame; itf.word_wrap = True
    for i, line in enumerate(info):
        p = itf.paragraphs[0] if i == 0 else itf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = line
        r.font.name = FONT; r.font.size = Pt(20); r.font.bold = False; r.font.color.rgb = BLACK

    # ---------------- SLIDE 2 : PRESENTATION OUTLINE ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Presentation Outline")
    _add_bullets(s, [
        ("Motivation", 0),
        ("Objectives", 0),
        ("Scope of Project", 0),
        ("Proposed Methodology", 0),
        ("Expected Results and Applications", 0),
        ("Timeline, Budget and References", 0),
    ])
    _add_footer(s, page)

    # ---------------- SLIDE 3 : MOTIVATION ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Motivation")
    _add_bullets(s, [
        ("High dependency on public transport", 0),
        ("Bus routes informal and undocumented", 0),
        ("Newcomers struggle to find buses", 0),
        ("Confusion over boarding and transfers", 0),
        ("Fare and travel time unknown", 0),
        ("Need a structured digital solution", 0),
    ])
    _add_footer(s, page)

    # ---------------- SLIDE 4 : OBJECTIVES ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Objectives")
    _add_bullets(s, [
        ("Develop web-based route recommendation system", 0),
        ("Target Kathmandu Valley public transit", 0),
        ("Recommend most suitable bus route", 0),
        ("Apply spatial indexing for stops", 1),
        ("Use Haversine formula for distance", 1),
        ("Use Dijkstra and scoring system", 1),
    ])
    _add_footer(s, page)

    # ---------------- SLIDE 5 : SCOPE OF PROJECT ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Scope of Project")
    _add_bullets(s, [
        ("Covers Kathmandu, Lalitpur, Bhaktapur", 0),
        ("Focuses on selected bus routes", 0),
        ("Uses structured route dataset", 0),
        ("No real-time traffic monitoring", 0),
        ("Serves commuters, students, tourists", 0),
        ("Useful for urban mobility research", 0),
    ])
    _add_footer(s, page)

    # ---------------- SLIDE 6 : METHODOLOGY - SYSTEM ARCHITECTURE ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Methodology: System Architecture")
    _maybe_image_or_placeholder(
        s, "block_diagram.png",
        "[ Insert Block Diagram here ]\n\nUser Input  ->  Spatial Search  ->  Route Dataset\n->  Route Computation  ->  Output Visualization")
    _add_footer(s, page)

    # ---------------- SLIDE 7 : METHODOLOGY - WORKING PRINCIPLE ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Methodology: Working Principle")
    _add_bullets(s, [
        ("User enters source and destination", 0),
        ("Spatial index finds nearby stops", 0),
        ("Dataset returns stop and routes", 0),
        ("Stops modeled as graph nodes", 0),
        ("Dijkstra computes least-cost route", 0),
        ("Scoring ranks the final route", 0),
    ])
    _add_footer(s, page)

    # ---------------- SLIDE 8 : METHODOLOGY - ALGORITHM & FLOWCHART ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Methodology: Algorithm & Flowchart")
    _maybe_image_or_placeholder(
        s, "flowchart.png",
        "[ Insert Flowchart here ]\n\nInput -> Validate -> Spatial Search -> Haversine\n-> Build Graph -> Dijkstra -> Score -> Display")
    _add_footer(s, page)

    # ---------------- SLIDE 9 : METHODOLOGY - INSTRUMENTATION ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Methodology: Instrumentation")
    _add_bullets(s, [
        ("Frontend: React.js and Leaflet.js", 0),
        ("Backend: FastAPI and Python", 0),
        ("Maps: OpenStreetMap tile service", 0),
        ("Database: SQLite / PostgreSQL / JSON", 0),
        ("Editor: Visual Studio Code", 0),
        ("Hardware: standard laptop, internet", 0),
    ])
    _add_footer(s, page)

    # ---------------- SLIDE 10 : EXPECTED RESULTS (1) ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Expected Results")
    _add_bullets(s, [
        ("Recommendation of direct bus routes", 0),
        ("Recommendation of transfer-based routes", 0),
        ("Display of boarding points", 0),
        ("Display of transfer points", 0),
        ("Clear destination stop guidance", 0),
        ("Route shown on interactive map", 0),
    ])
    _add_footer(s, page)

    # ---------------- SLIDE 11 : EXPECTED RESULTS (2) ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Expected Results (Contd.)")
    _add_bullets(s, [
        ("Approximate fare calculation", 0),
        ("Approximate travel time estimation", 0),
        ("Simple user-friendly web interface", 0),
        ("Structured dataset for selected routes", 0),
        ("Reduced confusion for newcomers", 0),
        ("Better access to transit information", 0),
    ])
    _add_footer(s, page)

    # ---------------- SLIDE 12 : PROJECT APPLICATIONS ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Project Applications")
    _add_bullets(s, [
        ("Daily commuters within the valley", 0),
        ("Students travelling to colleges", 0),
        ("Tourists exploring Kathmandu Valley", 0),
        ("Newcomers unfamiliar with routes", 0),
        ("Researchers studying urban mobility", 0),
        ("Basis for future transit apps", 0),
    ])
    _add_footer(s, page)

    # ---------------- SLIDE 13 : TENTATIVE TIMELINE (GANTT) ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Tentative Timeline (Gantt Chart)")
    _maybe_image_or_placeholder(
        s, "gantt.png",
        "[ Insert Gantt Chart here ]\n\nProposal -> Data Collection -> Design\n-> Implementation -> Testing -> Documentation")
    _add_footer(s, page)

    # ---------------- SLIDE 14 : ESTIMATED BUDGET (table) ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "Estimated Project Budget")
    rows, cols = 5, 2
    tbl_w = Inches(7.0); tbl_h = Inches(3.2)
    left = Inches(1.5); top = Inches(2.0)
    table = s.shapes.add_table(rows, cols, left, top, tbl_w, tbl_h).table
    table.columns[0].width = Inches(4.6)
    table.columns[1].width = Inches(2.4)
    # header
    _style_cell(table.cell(0, 0), "Item", bold=True, size=24, align=PP_ALIGN.LEFT, fill=HEADER_FILL)
    _style_cell(table.cell(0, 1), "Estimated Cost (NPR)", bold=True, size=24, align=PP_ALIGN.CENTER, fill=HEADER_FILL)
    data = [
        ("Internet / data collection", "500"),
        ("Software tools", "0"),
        ("Printing / documentation", "700"),
        ("Miscellaneous", "500"),
    ]
    for i, (item, cost) in enumerate(data, start=1):
        _style_cell(table.cell(i, 0), item, bold=False, size=24, align=PP_ALIGN.LEFT)
        _style_cell(table.cell(i, 1), cost, bold=False, size=24, align=PP_ALIGN.CENTER)
    # total row styled bold via a 6th row -> instead append note
    total_box = s.shapes.add_textbox(Inches(1.5), Inches(5.4), Inches(7.0), Inches(0.6))
    p = total_box.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = "Total:  NPR 1700"
    r.font.name = FONT; r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = BLACK
    _add_footer(s, page)

    # ---------------- SLIDE 15 : REFERENCES (1) ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "References")
    refs1 = [
        "[1] E. W. Dijkstra, A note on two problems in connexion with graphs, Numerische Mathematik, vol. 1, pp. 269-271, 1959.",
        "[2] T. H. Cormen et al., Introduction to Algorithms, 3rd ed. MIT Press, 2009.",
        "[3] M. A. Rahman et al., Shortest path algorithms in transportation networks: A survey, IJCA, vol. 181, no. 14, 2018.",
        "[4] S. Zhang et al., A survey of route recommendations, J. Transp. Tech., vol. 12, no. 3, 2022.",
        "[5] Q. Shi et al., Evaluation model of bus route optimization, IEEE ITSC, 2020.",
    ]
    box = s.shapes.add_textbox(Inches(0.6), Inches(1.6), Inches(8.8), Inches(5.0))
    tf = box.text_frame; tf.word_wrap = True
    for i, ref in enumerate(refs1):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT; p.space_after = Pt(10)
        r = p.add_run(); r.text = ref
        r.font.name = FONT; r.font.size = Pt(20); r.font.color.rgb = BLACK
    _add_footer(s, page)

    # ---------------- SLIDE 16 : REFERENCES (2) ----------------
    page += 1
    s = _blank_slide(prs)
    _add_title(s, "References (Contd.)")
    refs2 = [
        "[6] OpenStreetMap. [Online]. Available: https://www.openstreetmap.org/",
        "[7] Leaflet: An open-source JavaScript library. [Online]. Available: https://leafletjs.com/",
        "[8] Google Maps. [Online]. Available: https://www.google.com/maps",
        "[9] Moovit transit app. [Online]. Available: https://moovitapp.com/",
        "[10] MobilityData, GTFS Reference. [Online]. Available: https://gtfs.org/",
    ]
    box = s.shapes.add_textbox(Inches(0.6), Inches(1.6), Inches(8.8), Inches(5.0))
    tf = box.text_frame; tf.word_wrap = True
    for i, ref in enumerate(refs2):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT; p.space_after = Pt(10)
        r = p.add_run(); r.text = ref
        r.font.name = FONT; r.font.size = Pt(20); r.font.color.rgb = BLACK
    _add_footer(s, page)

    # ---------------- SLIDE 17 : THANK YOU ----------------
    page += 1
    s = _blank_slide(prs)
    box = s.shapes.add_textbox(Inches(0.5), Inches(2.8), Inches(9.0), Inches(1.5))
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = "Thank You!"
    r.font.name = FONT; r.font.size = Pt(40); r.font.bold = True; r.font.color.rgb = BLACK
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run(); r2.text = "Questions & Discussion"
    r2.font.name = FONT; r2.font.size = Pt(28); r2.font.bold = False; r2.font.color.rgb = BLACK
    _add_footer(s, page)

    prs.save(OUTPUT_FILE)
    print(f"Saved {OUTPUT_FILE} with {len(prs.slides._sldIdLst)} slides.")


if __name__ == "__main__":
    build()
