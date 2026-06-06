# -*- coding: utf-8 -*-
"""
JAL Strategies — Queenstown Harbor Background Research (Word / .docx)

Generates JAL_Queenstown_Harbor_Research.docx — a cited background memo on
Queenstown Harbor (what it is, how it came about, how it was funded), formatted
on-brand to match the proposal deck (navy / aubergine / plum; Montserrat
headings, DM Sans body). Companion to generate.js (the pitch deck).

Facts are corroborated across multiple public sources; sponsor projections are
labeled as such. Run:  python3 generate_research_doc.py
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT

OUT = "JAL_Queenstown_Harbor_Research.docx"

# ----- Brand palette (matches generate.js) ---------------------------------
NAVY      = RGBColor(0x0B, 0x16, 0x3C)
AUBERGINE = RGBColor(0x3A, 0x24, 0x3A)
PLUM      = RGBColor(0x6B, 0x5A, 0x6B)
MAUVE     = RGBColor(0xC4, 0xB8, 0xC4)
SLATE     = RGBColor(0x8C, 0x9B, 0xB5)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
CREAM_HEX = "F4F2ED"
NAVY_HEX  = "0B163C"
MAUVE_HEX = "C4B8C4"

HEAD = "Montserrat"   # falls back gracefully if not installed on the reader's machine
BODY = "DM Sans"

# ----- Low-level helpers ----------------------------------------------------
def set_run(run, font=BODY, size=10, bold=False, italic=False, color=None,
            spacing=None, allcaps=False):
    run.font.name = font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts"); rPr.append(rFonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(a), font)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    if spacing is not None:                       # letter-spacing in twips
        sp = OxmlElement("w:spacing"); sp.set(qn("w:val"), str(spacing)); rPr.append(sp)
    if allcaps:
        caps = OxmlElement("w:caps"); caps.set(qn("w:val"), "true"); rPr.append(caps)
    return run

def para(doc, space_before=0, space_after=6, align=None, line=1.12):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if align is not None:
        p.alignment = align
    return p

def runp(doc, text, **kw):
    """One-run paragraph."""
    layout = {k: kw.pop(k) for k in ("space_before", "space_after", "align", "line") if k in kw}
    p = para(doc, **layout)
    set_run(p.add_run(text), **kw)
    return p

def bottom_border(paragraph, color=MAUVE_HEX, size=6):
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single"); b.set(qn("w:sz"), str(size))
    b.set(qn("w:space"), "6"); b.set(qn("w:color"), color)
    pbdr.append(b); pPr.append(pbdr)

def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    tcPr.append(shd)

def add_hyperlink(paragraph, url, text, size=8.5, color="3A243A"):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink"); link.set(qn("r:id"), r_id)
    r = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(a), BODY)
    rPr.append(rFonts)
    c = OxmlElement("w:color"); c.set(qn("w:val"), color); rPr.append(c)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size * 2))); rPr.append(sz)
    r.append(rPr)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text
    r.append(t); link.append(r)
    paragraph._p.append(link)
    return link

def eyebrow(doc, text, space_before=14):
    p = runp(doc, text, font=HEAD, size=8.5, bold=True, color=AUBERGINE,
             spacing=40, space_before=space_before, space_after=2)
    return p

def h2(doc, text, space_before=16):
    eyebrow_done = None
    p = runp(doc, text, font=HEAD, size=15, bold=True, color=NAVY,
             space_before=space_before, space_after=4, line=1.05)
    bottom_border(p, color=MAUVE_HEX, size=6)
    return p

def h3(doc, text):
    return runp(doc, text, font=HEAD, size=11.5, bold=True, color=AUBERGINE,
                space_before=10, space_after=3)

def body(doc, runs, space_before=0, space_after=7, line=1.15, align=None):
    """runs: list of (text, kw) tuples for mixed formatting."""
    p = para(doc, space_before=space_before, space_after=space_after, line=line, align=align)
    for text, kw in runs:
        set_run(p.add_run(text), **{**dict(font=BODY, size=10, color=NAVY), **kw})
    return p

def bullet(doc, runs, space_after=4):
    p = para(doc, space_after=space_after, line=1.13)
    p.paragraph_format.left_indent = Inches(0.28)
    p.paragraph_format.first_line_indent = Inches(-0.16)
    set_run(p.add_run("•  "), font=BODY, size=10, color=AUBERGINE, bold=True)
    for text, kw in runs:
        set_run(p.add_run(text), **{**dict(font=BODY, size=10, color=NAVY), **kw})
    return p

def keyval_table(doc, rows, label_w=2.5, val_w=4.0):
    t = doc.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.allow_autofit = False
    for k, v in rows:
        cells = t.add_row().cells
        cells[0].width = Inches(label_w); cells[1].width = Inches(val_w)
        shade(cells[0], CREAM_HEX)
        pk = cells[0].paragraphs[0]; pk.paragraph_format.space_after = Pt(2); pk.paragraph_format.space_before = Pt(2)
        set_run(pk.add_run(k), font=HEAD, size=9, bold=True, color=NAVY)
        pv = cells[1].paragraphs[0]; pv.paragraph_format.space_after = Pt(2); pv.paragraph_format.space_before = Pt(2)
        set_run(pv.add_run(v), font=BODY, size=9.5, color=NAVY)
    _set_table_borders(t)
    return t

def grid_table(doc, headers, rows, widths):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.allow_autofit = False
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].width = Inches(widths[i])
        shade(hdr[i], NAVY_HEX)
        p = hdr[i].paragraphs[0]; p.paragraph_format.space_after = Pt(2); p.paragraph_format.space_before = Pt(2)
        set_run(p.add_run(h), font=HEAD, size=9, bold=True, color=WHITE)
    for r in rows:
        cells = t.add_row().cells
        for i, val in enumerate(r):
            cells[i].width = Inches(widths[i])
            p = cells[i].paragraphs[0]; p.paragraph_format.space_after = Pt(2); p.paragraph_format.space_before = Pt(2)
            kw = dict(font=BODY, size=9, color=NAVY)
            if i == 0:
                kw.update(font=HEAD, bold=True, color=AUBERGINE, size=9.5)
            elif i == 1:
                kw.update(font=HEAD, bold=True, color=NAVY, size=9)
            else:
                kw.update(color=PLUM)
            set_run(p.add_run(val), **kw)
    _set_table_borders(t)
    return t

def _set_table_borders(t, color="E6E1DB", size=4):
    tbl = t._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(size))
        e.set(qn("w:space"), "0"); e.set(qn("w:color"), color)
        borders.append(e)
    tblPr.append(borders)

def callout(doc, lines):
    """Single-cell shaded box with a left accent bar (cream fill, navy text)."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = t.rows[0].cells[0]
    cell.width = Inches(6.5)
    shade(cell, CREAM_HEX)
    # left accent border
    tcPr = cell._tc.get_or_add_tcPr()
    bdr = OxmlElement("w:tcBorders")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single"); left.set(qn("w:sz"), "24")
    left.set(qn("w:space"), "0"); left.set(qn("w:color"), "3A243A")
    bdr.append(left); tcPr.append(bdr)
    first = True
    for label, text in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_after = Pt(2); p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.line_spacing = 1.12
        if label:
            set_run(p.add_run(label + " "), font=HEAD, size=9, bold=True, color=AUBERGINE)
        set_run(p.add_run(text), font=BODY, size=9, color=NAVY)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

# ============================================================ BUILD ==========
doc = Document()
doc.core_properties.title = "Queenstown Harbor — Background Research"
doc.core_properties.author = "JAL Strategies"

# Base style
normal = doc.styles["Normal"]
normal.font.name = BODY
normal.font.size = Pt(10)
normal.font.color.rgb = NAVY
sect = doc.sections[0]
sect.top_margin = Inches(0.8); sect.bottom_margin = Inches(0.8)
sect.left_margin = Inches(0.9); sect.right_margin = Inches(0.9)

# ----- Title block ----------------------------------------------------------
runp(doc, "JAL STRATEGIES  ·  BACKGROUND RESEARCH", font=HEAD, size=9, bold=True,
     color=PLUM, spacing=50, space_after=4)
runp(doc, "Queenstown Harbor", font=HEAD, size=26, bold=True, color=NAVY,
     space_after=2, line=1.0)
p = runp(doc, "History, Developer & Funding — a cited background memo", font=HEAD,
         size=13, bold=True, color=AUBERGINE, space_after=6)
meta = body(doc, [
    ("Prepared for ", dict(size=9.5, color=PLUM)),
    ("JAL Strategies (Justin A. Levine)", dict(size=9.5, color=NAVY, bold=True)),
    ("    ·    June 2026    ·    Compiled from public sources; each claim cited.", dict(size=9.5, color=PLUM)),
], space_after=8)
bottom_border(meta, color=NAVY_HEX, size=10)

# ----- Confidence note ------------------------------------------------------
callout(doc, [
    ("Confidence note.",
     "Facts about the property, its courses, its developer, and the 2026 transaction are corroborated "
     "across multiple independent sources and are high-confidence. Forward-looking figures (EBITDA, "
     "refinancing timing, the “perpetual ownership” mechanics) come from the acquirer’s own "
     "investor-marketing materials and are labeled [sponsor projection] — treat them as promotional, "
     "not independently verified. One minor opening-date discrepancy is flagged inline."),
])

# ----- 1. At a glance -------------------------------------------------------
eyebrow(doc, "OVERVIEW")
h2(doc, "1.  At a glance", space_before=2)
keyval_table(doc, [
    ("Name", "Queenstown Harbor (a.k.a. Queenstown Harbor Golf Links / Golf Resort)"),
    ("Location", "310 Links Lane, Queenstown, MD 21658 — Maryland’s Eastern Shore, where the Chester River meets the Chesapeake Bay"),
    ("Type", "Public / daily-fee waterfront golf resort"),
    ("Golf", "36 holes — The River (par 72, ~7,096 yds) and The Lakes (par 71, ~6,569 yds)"),
    ("Architect", "Lindsay B. Ervin (LBE Golf Course Design)"),
    ("Opened", "River Course 1991; built out to 36 holes by 1996"),
    ("Original developer", "The Brick Companies (Edgewater, MD; founded 1892)"),
    ("Current owner (May 2026)", "Accountable Equity, LLC via the Capital H6 fund; operated by VIVÂMEE Hospitality"),
    ("Acquisition price", "~$25 million — combined with The Golf Club at South River"),
    ("Land", "~870+ acres across the two acquired properties, incl. ~198 acres under permanent conservation"),
    ("Amenities", "River House & Tavern event venue, on-site cottages/bungalows, Toptracer driving range, short-game areas"),
])

# ----- 2. The property today -----------------------------------------------
eyebrow(doc, "THE ASSET")
h2(doc, "2.  The property today", space_before=2)
body(doc, [("Queenstown Harbor markets itself as a waterfront golf ", {}),
           ("resort", dict(italic=True)),
           (" roughly an hour from Washington, D.C. and Baltimore, and minutes over the Bay Bridge from "
            "Annapolis. It pairs two championship courses with lodging and an events business.", {})])
body(doc, [("The River Course", dict(bold=True)),
           (" — the original layout — runs along the mouth of the Chester River, with water in play on "
            "14 of 18 holes and panoramic Bay/river views on more than half the course. It plays to par 72 at "
            "roughly 7,096 yards (course rating ~74.2), the sterner of the two tests.", {})])
body(doc, [("The Lakes Course", dict(bold=True)),
           (" is the more forgiving, broadly playable layout (par 71, ~6,569 yards). It has repeatedly been named "
            "one of Golf Digest’s Top 50 Public Courses for Women. In 2024 it underwent a major bunker "
            "renovation — 42 traditional bunkers removed and replaced with 20 grass-faced, flat-bottom bunkers.", {})])
body(doc, [("Beyond golf:", dict(bold=True)),
           (" a public Toptracer driving range, two putting greens, a chipping green and a practice bunker; the "
            "River House & Tavern event venue (indoor space plus an outdoor patio/lawn over the Chester River); and "
            "on-site cottages and bungalows for overnight stays.", {})])

# ----- 3. How it came about -------------------------------------------------
eyebrow(doc, "PROVENANCE")
h2(doc, "3.  How it came about", space_before=2)
h3(doc, "The land: from a colonial “thumb grant” to a dairy farm")
body(doc, [("The site carries a deep colonial history the club itself foregrounds. By its own account, the land was "
            "first granted in the 1600s to ", {}),
           ("Henry deCoursey", dict(bold=True)),
           (", a surveyor, by ", {}),
           ("Cecil Calvert, the second Lord Baltimore", dict(bold=True)),
           (". As the story goes, deCoursey pressed his thumb onto a map and was given all the land beneath the "
            "thumbprint — known locally for centuries as the “thumb grant” (also recorded as "
            "“My Lord’s Gift”).", {})])
body(doc, [("Treat the thumb-grant story as the property’s own narrative / local lore. ", dict(italic=True, color=PLUM)),
           ("The broader fact — that all early Maryland land flowed from grants issued by the Calvert "
            "proprietors — is well documented; the colorful thumbprint detail is the club’s telling.",
            dict(italic=True, color=PLUM))])
body(doc, [("By the time ", {}), ("The Brick Companies", dict(bold=True)),
           (" found the property in the 1970s, it was a working dairy farm and crop fields. Company executives "
            "decided to build a golf course there specifically to open the land’s beauty and wildlife to the "
            "public.", {})])
body(doc, [("During construction, a large ", {}), ("archaeological dig", dict(bold=True)),
           (" under the original deCoursey home site recovered and preserved millions of artifacts — many now "
            "displayed at The Brick Companies’ headquarters in Edgewater, MD, and many donated to the "
            "Maryland Historical Trust.", {})])

h3(doc, "Design and buildout timeline")
grid_table(doc,
    ["Era", "Milestone", "What happened"],
    [
        ["1600s", "“My Lord’s Gift”", "Surveyor Henry deCoursey is granted the tract by Cecil Calvert, 2nd Lord Baltimore — the “thumb grant” (club lore)."],
        ["1970s", "Dairy farm acquired", "The Brick Companies (a private Maryland firm founded 1892) buys the working farm to open its land to the public as golf."],
        ["1991", "River Course opens", "Lindsay Ervin’s design debuts on the Chester River; a construction-era archaeological dig preserves millions of artifacts."],
        ["mid-1990s", "Expands to 27 holes", "A third nine is added as the destination grows."],
        ["1996", "36 holes complete", "A fourth nine completes the build; the 3rd and 4th nines combine into the Lakes Course — today’s 36 holes."],
        ["May 2026", "Capital H6 acquires", "Accountable Equity / VIVÂMEE buy Queenstown Harbor (with South River) for ~$25M from The Brick Companies."],
    ],
    widths=[1.0, 1.7, 3.8])
body(doc, [("Date discrepancy flagged: ", dict(bold=True, color=AUBERGINE)),
           ("golf-course databases put the River Course opening at 1991; the current owner’s investor "
            "materials instead say the resort has been “operating since 1994.” The 1991 date is the more "
            "consistently cited across independent golf sources; 1994 may reflect when the multi-course "
            "destination was fully established.", dict(color=PLUM, size=9.5))],
     space_before=4)

# ----- 4. The developer -----------------------------------------------------
eyebrow(doc, "THE DEVELOPER")
h2(doc, "4.  The developer: The Brick Companies", space_before=2)
body(doc, [("Queenstown Harbor was conceived, built, and owned for ~30 years by ", {}),
           ("The Brick Companies (TBC)", dict(bold=True)),
           (" — a privately held, family-owned firm headquartered in Edgewater, Maryland.", {})])
bullet(doc, [("Founded 1892 ", dict(bold=True)),
             ("as the Washington Brick and Terra Cotta Company, along the Washington City Canal in Southwest D.C.; "
              "for decades its sole business was brick and terra-cotta pipe.", {})])
bullet(doc, [("A 1939 fire ", dict(bold=True)),
             ("destroyed the plant; the company never made another brick, and instead diversified.", {})])
bullet(doc, [("Relocated ", dict(bold=True)),
             ("its headquarters from D.C. to Edgewater, MD in 2002.", {})])
bullet(doc, [("Today ", dict(bold=True)),
             ("TBC owns, develops and manages commercial, residential and recreational property across D.C., "
              "Maryland and Virginia — including golf (Queenstown Harbor, The Golf Club at South River) and "
              "Atlantic Marinas.", {})])
body(doc, [("Why this matters for funding: ", dict(bold=True, color=AUBERGINE)),
           ("the course was a diversification play by a cash-generating private real-estate company — not a "
            "venture-financed or publicly funded development.", {})], space_before=4)

# ----- 5. How it was funded -------------------------------------------------
eyebrow(doc, "FUNDING")
h2(doc, "5.  How it was funded", space_before=2)
h3(doc, "(a)  The original development (1970s–1996)")
body(doc, [("The honest answer: ", dict(bold=True)),
           ("no public construction-cost figure exists — The Brick Companies is privately held and has never "
            "disclosed what it spent to build Queenstown Harbor. What the record does support:", {})])
bullet(doc, [("The land was acquired by The Brick Companies in the 1970s (former dairy farm) and developed in-house.", {})])
bullet(doc, [("TBC describes earlier business partnerships as having provided the “financial platform” that "
              "let it diversify into golf courses, marinas and real estate — i.e., the course was funded out of "
              "the company’s own balance sheet: ", {}),
             ("private, self-funded equity", dict(bold=True)), (", not outside project finance.", {})])
body(doc, [("Bottom line: ", dict(bold=True, color=AUBERGINE)),
           ("a privately financed amenity/real-estate development by an established family company. A hard "
            "development cost would have to come from TBC directly or Queen Anne’s County records.", {})])

h3(doc, "(b)  The May 2026 acquisition (~$25M)")
body(doc, [("In May 2026, ", {}), ("VIVÂMEE Hospitality, LLC", dict(bold=True)),
           (" and its affiliate ", {}), ("Accountable Equity, LLC", dict(bold=True)),
           (" completed a combined ~$25 million acquisition of two Mid-Atlantic assets from The Brick Companies: "
            "Queenstown Harbor Golf Resort (Eastern Shore) and The Golf Club at South River (near Annapolis). "
            "Together the portfolio is 54 holes across 870+ acres of waterfront and conservation-protected land.", {})])
h3(doc, "The financing / ownership structure")
bullet(doc, [("Seller: ", dict(bold=True)), ("The Brick Companies (confirmed via John Anderes, TBC’s EVP of Hospitality & Golf, in the transaction coverage).", {})])
bullet(doc, [("Acquisition vehicle: ", dict(bold=True)), ("the deal ran through Capital H6, LLC, a single-purpose fund sponsored by Accountable Equity, LLC.", {})])
bullet(doc, [("Equity source: ", dict(bold=True)), ("capital raised from accredited investors under a Rule 506(c) private placement (a Capital H6 Confidential Private Placement Memorandum dated November 2025).", {})])
bullet(doc, [("Operator: ", dict(bold=True)), ("VIVÂMEE Hospitality runs the properties. Accountable Equity and VIVÂMEE were co-founded by Josh and Melanie McCallen; Josh McCallen is CEO of both.", {})])
bullet(doc, [("Portfolio effect: ", dict(bold=True)), ("the closing expanded the VIVÂMEE platform to six resort properties and five championship golf courses across Maryland and New Jersey (incl. the adjacent Kent Island Resort, Renault Winery Resort, LBI National Golf & Resort, Bohemia Manor Farm).", {})])
body(doc, [("Sponsor’s investment thesis — [sponsor projection]: ", dict(bold=True, color=AUBERGINE)),
           ("Accountable Equity’s materials pitch a “perpetual ownership” model (no forced 3–5-year "
            "exit), with a refinancing event expected around Years 4–5 returning a portion of capital plus "
            "accrued preferred returns; they advertise >$2.5M projected Year-1 EBITDA and a growth plan leaning "
            "into lodging, events, weddings, corporate retreats and wellness. ", {}),
           ("These are the sponsor’s own projections and should be independently diligenced.",
            dict(italic=True, color=PLUM))], space_before=4)

# ----- 6. Sister asset ------------------------------------------------------
eyebrow(doc, "PORTFOLIO CONTEXT")
h2(doc, "6.  The sister asset: The Golf Club at South River", space_before=2)
body(doc, [("Acquired in the same $25M transaction:", {})])
bullet(doc, [("Location: ", dict(bold=True)), ("Edgewater, MD (Annapolis / South River side).", {})])
bullet(doc, [("Course: ", dict(bold=True)), ("private 18-hole layout designed by Brian Ault; opened as public in 1996, converted to private in 2007.", {})])
bullet(doc, [("Membership: ", dict(bold=True)), ("600+ member families — recurring dues-based revenue.", {})])
bullet(doc, [("Footprint: ", dict(bold=True)), ("~165 acres, with 8 lakes and multiple protected areas.", {})])

# ----- 7. Open items --------------------------------------------------------
eyebrow(doc, "DILIGENCE")
h2(doc, "7.  Open items / things to confirm directly", space_before=2)
bullet(doc, [("Original development cost ", dict(bold=True)), ("of Queenstown Harbor — not public (TBC is private). Source: TBC or Queen Anne’s County records.", {})])
bullet(doc, [("Exact opening year ", dict(bold=True)), ("— 1991 (golf databases) vs. “since 1994” (owner materials).", {})])
bullet(doc, [("The $25M allocation ", dict(bold=True)), ("between Queenstown Harbor and South River — reported only as a combined price.", {})])
bullet(doc, [("Acreage split ", dict(bold=True)), ("— the 870+ acres and 198-acre conservation figures are combined across both properties; confirm the Queenstown-only developable acreage against survey/easement data.", {})])
bullet(doc, [("“Perpetual ownership,” EBITDA, refinancing timeline ", dict(bold=True)), ("— sponsor projections; verify against the PPM and financials.", {})])

# ----- 8. Sources -----------------------------------------------------------
eyebrow(doc, "CITATIONS")
h2(doc, "8.  Sources", space_before=2)

SOURCE_GROUPS = [
    ("Property / club (official & golf databases)", [
        ("Queenstown Harbor — Who We Are (history, deCoursey grant, dairy farm, dig)", "https://qhgolf.com/who-we-are/"),
        ("Queenstown Harbor — Golf (courses, amenities, 2024 bunker renovation)", "https://qhgolf.com/golf/"),
        ("VisitMaryland — Queenstown Harbor Golf Links", "https://www.visitmaryland.org/listing/sports/queenstown-harbor-golf-links"),
        ("Visit Queen Anne’s County — Queenstown Harbor Golf", "https://visitqueenannes.com/business/queenstown-harbor-golf/"),
        ("Maryland State Golf Association — Queenstown Harbor", "https://msga.org/queenstown-harbor/"),
        ("Where2Golf — River Course (opening / buildout history)", "https://www.where2golf.com/usa-mid-atlantic/queenstown-harbor-river-course/"),
        ("Golf Digest — Queenstown Harbor: The River", "https://www.golfdigest.com/courses/md/queenstown-harbor-the-river"),
        ("Golf Digest — Queenstown Harbor: The Lakes", "https://www.golfdigest.com/courses/md/queenstown-harbor-the-lakes"),
        ("foreTee — River Course specs (yardage, rating)", "https://foretee.com/courses/maryland/queenstown/usa/atlantic-golf-at-queenstown-harbor--river/6614"),
        ("LBE Golf Course Design (Lindsay Ervin) — Queenstown", "https://www.lbegolf.com/queenstown"),
        ("Golf the Mid-Atlantic — Lakes “Tale of Two Nines”", "http://www.golfthemidatlantic.com/story/161"),
    ]),
    ("Developer — The Brick Companies", [
        ("The Brick Companies — Golf Courses", "https://www.thebrickcompanies.com/golf-courses"),
        ("The Brick Companies — Our Legacy (founded 1892; diversification)", "https://www.thebrickcompanies.com/legacy"),
    ]),
    ("Land history", [
        ("MyEasternShoreMD — history of names (My Lord’s Gift / thumb grant / Blakeford)", "https://www.myeasternshoremd.com/news/queen_annes_county/book-tells-interesting-history-of-names/article_a269685c-2a1c-565b-b1a9-17716da29511.html"),
    ]),
    ("2026 acquisition & financing", [
        ("Connect CRE — VIVÂMEE completes $25M acquisition", "https://www.connectcre.com/stories/vivamee-hospitality-completes-25m-acquisition-of-two-maryland-golf-courses/"),
        ("Club + Resort Business — VIVÂMEE acquires two MD golf locations for $25M", "https://clubandresortbusiness.com/vivamee-hospitality-acquires-two-maryland-golf-locations-for-25m/"),
        ("Eye On Annapolis — VIVÂMEE acquires South River, Queenstown Harbor", "https://www.eyeonannapolis.net/2026/05/vivamee-hospitality-acquires-golf-club-at-south-river-queenstown-harbor-in-25m-deal/"),
        ("citybiz — VIVÂMEE bolsters MD golf & waterfront portfolio with $25M", "https://www.citybiz.co/article/847876/vivamee-hospitality-bolsters-maryland-golf-waterfront-portfolio-with-25-million-acquisition/"),
        ("Accountable Equity — Capital H6 hospitality fund (structure, EBITDA)", "https://accountableequity.com/capital-h6-hospitality-fund/"),
        ("Accountable Equity — Capital H6 Private Placement Memorandum (PDF, Nov 2025)", "https://accountableequity.com/wp-content/uploads/2025/11/Capital-H6-Private-Placement-Memorandum.pdf"),
        ("Accountable Equity — About Us (Josh & Melanie McCallen)", "https://accountableequity.com/about-us/"),
        ("Accountable Equity — accredited-investor / Rule 506(c)", "https://accountableequity.com/what-can-accredited-investors-access/"),
        ("Josh McCallen — VIVÂMEE / LinkedIn", "https://www.linkedin.com/in/joshmccallen/"),
    ]),
    ("Sister asset — The Golf Club at South River", [
        ("golfclubsr.com — South River golf", "https://www.golfclubsr.com/golf"),
        ("GolfLink — Golf Club at South River (Brian Ault, 1996; private since 2007)", "https://www.golflink.com/golf-courses/md/edgewater/golf-club-at-south-river"),
    ]),
]

n = 1
for group, items in SOURCE_GROUPS:
    runp(doc, group, font=HEAD, size=9.5, bold=True, color=AUBERGINE, space_before=8, space_after=2)
    for label, url in items:
        p = para(doc, space_after=2, line=1.1)
        p.paragraph_format.left_indent = Inches(0.28)
        p.paragraph_format.first_line_indent = Inches(-0.28)
        set_run(p.add_run(f"{n}.  "), font=BODY, size=8.5, bold=True, color=NAVY)
        set_run(p.add_run(label + "  —  "), font=BODY, size=8.5, color=NAVY)
        add_hyperlink(p, url, url, size=8.5, color="3A243A")
        n += 1

# ----- Closing note ---------------------------------------------------------
closing = body(doc, [
    ("Compiled from multi-source web research with adversarial cross-checking. High-confidence on property, "
     "developer, and transaction facts; sponsor projections labeled as such; one opening-date discrepancy "
     "flagged. For discussion purposes.", dict(italic=True, size=8.5, color=PLUM))],
    space_before=12)
closing.alignment = WD_ALIGN_PARAGRAPH.LEFT
top = closing._p.get_or_add_pPr()
pbdr = OxmlElement("w:pBdr")
tb = OxmlElement("w:top")
tb.set(qn("w:val"), "single"); tb.set(qn("w:sz"), "6"); tb.set(qn("w:space"), "8"); tb.set(qn("w:color"), MAUVE_HEX)
pbdr.append(tb); top.append(pbdr)

doc.save(OUT)
print("Wrote", OUT)
