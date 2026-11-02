"""Build an editable 16:9 PowerPoint slide of the journal's history (South Dakota State University colours)."""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE

PX = 6350  # EMU per design pixel (1920 px = 13.333 in)
BLUE = RGBColor(0x00, 0x33, 0xA0)     # SDSU Jackrabbit Blue
YELLOW = RGBColor(0xFF, 0xD1, 0x00)   # SDSU Yellow
GOLD = RGBColor(0x8C, 0x6D, 0x00)     # darker yellow, readable as text
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x54, 0x58, 0x5A)     # dark gray
LGRAY = RGBColor(0x9E, 0xA2, 0xA2)    # light gray
MUTED = RGBColor(0x5F, 0x63, 0x66)
GUIDE = RGBColor(0xD0, 0xD0, 0xD0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"

prs = Presentation()
prs.slide_width, prs.slide_height = Emu(1920 * PX), Emu(1080 * PX)
s = prs.slides.add_slide(prs.slide_layouts[6])


def bez(p0, p1, p2, p3, n=24):
    out = []
    for i in range(n + 1):
        t = i / n
        out.append(tuple((1-t)**3*a + 3*(1-t)**2*t*b + 3*(1-t)*t**2*c + t**3*d
                         for a, b, c, d in zip(p0, p1, p2, p3)))
    return out


def path(points, color, width_px, name, dash=False):
    fb = s.shapes.build_freeform(points[0][0] * PX, points[0][1] * PX, scale=1.0)
    fb.add_line_segments([(x * PX, y * PX) for x, y in points[1:]], close=False)
    sh = fb.convert_to_shape()
    sh.name = name
    sh.fill.background()
    sh.line.color.rgb = color
    sh.line.width = Emu(width_px * PX)
    if dash:
        sh.line.dash_style = MSO_LINE_DASH_STYLE.SQUARE_DOT
    sh._element.spPr.get_or_add_ln().set("cap", "rnd")
    return sh


def text(x, y, w, h, runs, size, color, bold=False, italic=False, align=PP_ALIGN.LEFT):
    tb = s.shapes.add_textbox(Emu(x * PX), Emu(y * PX), Emu(w * PX), Emu(h * PX))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    lines = runs if isinstance(runs, list) else [runs]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for seg in (line if isinstance(line, tuple) else (line,)):
            r = p.add_run()
            r.text = seg if isinstance(seg, str) else seg[0]
            f = r.font
            f.name, f.size, f.color.rgb, f.bold = FONT, Pt(size), color, bold
            f.italic = italic or (not isinstance(seg, str) and seg[1])
    return tb


def dot(cx, cy, r, ring):
    o = s.shapes.add_shape(MSO_SHAPE.OVAL, Emu((cx-r)*PX), Emu((cy-r)*PX), Emu(2*r*PX), Emu(2*r*PX))
    o.fill.solid(); o.fill.fore_color.rgb = WHITE
    o.line.color.rgb = ring; o.line.width = Emu(8 * PX)
    return o


# Title
text(120, 70, 1700, 70, [("How ", ("Landscape and Urban Planning", True), " came together")], 30, BLUE, bold=True)
text(120, 140, 1700, 40, "Three journals merged into one between 1974 and 1988 (publisher: Elsevier)", 16, MUTED)

# Year guides
X74, X86, X88 = 220, 900, 1260
for x in (X74, X86, X88):
    g = s.shapes.add_connector(1, Emu(x*PX), Emu(250*PX), Emu(x*PX), Emu(900*PX))
    g.line.color.rgb = GUIDE; g.line.width = Pt(1.5); g.line.dash_style = MSO_LINE_DASH_STYLE.DASH

MAIN_Y = 520
# Feeder journals
path([(X74, 360), (800, 360)] + bez((800, 360), (860, 360), (860, MAIN_Y), (X86, MAIN_Y)), YELLOW, 34, "Landscape Planning")
path([(360, 680), (520, 680)], GRAY, 10, "Urban Ecology (earlier)", dash=True)
path([(560, 680), (800, 680)] + bez((800, 680), (860, 680), (860, MAIN_Y), (X86, MAIN_Y)), GRAY, 34, "Urban Ecology")
path([(800, 800), (940, 800)], LGRAY, 10, "RRR (earlier)", dash=True)
path([(980, 800), (1160, 800)] + bez((1160, 800), (1220, 800), (1220, MAIN_Y), (X88, MAIN_Y)), LGRAY, 34, "Reclamation and Revegetation Research")

# Main journal + arrowhead
path([(X86, MAIN_Y), (1740, MAIN_Y)], BLUE, 44, "Landscape and Urban Planning")
fb = s.shapes.build_freeform(1735*PX, 478*PX, scale=1.0)
fb.add_line_segments([(1805*PX, MAIN_Y*PX), (1735*PX, 562*PX)], close=True)
arrow = fb.convert_to_shape(); arrow.name = "Arrowhead"
arrow.fill.solid(); arrow.fill.fore_color.rgb = BLUE; arrow.line.fill.background()

dot(X74, 360, 16, YELLOW); dot(X86, MAIN_Y, 18, BLUE); dot(X88, MAIN_Y, 18, BLUE)

# Labels
text(240, 300, 600, 40, "Landscape Planning", 16, GOLD, bold=True)
text(240, 400, 600, 70, ["Founded by Elsevier", "Founding editor-in-chief: Arnold E. Weddle"], 12, MUTED)
text(480, 625, 400, 40, "Urban Ecology", 16, GRAY, bold=True)
text(880, 845, 800, 40, "Reclamation and Revegetation Research", 16, RGBColor(0x7A, 0x7E, 0x80), bold=True)
text(960, 440, 800, 50, "Landscape and Urban Planning", 19, BLUE, bold=True)
text(1650, 580, 160, 40, "Today", 13, BLUE, bold=True)

# Axis + years
ax = s.shapes.add_connector(1, Emu(120*PX), Emu(900*PX), Emu(1800*PX), Emu(900*PX))
ax.line.color.rgb = BLACK; ax.line.width = Pt(2)
for x, yr, cap in [(X74, "1974", ["Journal founded"]),
                   (X86, "1986", ["Renamed; merged with", "Urban Ecology"]),
                   (X88, "1988", ["Absorbs Reclamation &", "Revegetation Research"])]:
    text(x-160, 925, 320, 50, yr, 24, BLUE, bold=True, align=PP_ALIGN.CENTER)
    text(x-160, 985, 320, 70, cap, 12, MUTED, align=PP_ALIGN.CENTER)

prs.save("journal-timeline-SDState.pptx")
print("saved")
