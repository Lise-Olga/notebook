"""Editable PowerPoint table: most-cited Landscape and Urban Planning articles (SDSU colours)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

BLUE = RGBColor(0x00, 0x33, 0xA0)      # SDSU Jackrabbit Blue
YELLOW = RGBColor(0xFF, 0xD1, 0x00)    # SDSU Yellow
TINT = RGBColor(0xEE, 0xF2, 0xFA)      # light blue row band
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x1A, 0x1A, 0x1A)
FONT = "Arial"

rows = [
    ("MEEROW, S.; NEWELL, J.; STULTS, M. Defining urban resilience: A review. "
     "LANDSCAPE AND URBAN PLANNING, 2016, 147: 38–49.", "2504"),
    ("CHIESURA, A. The role of urban parks for the sustainable city. "
     "LANDSCAPE AND URBAN PLANNING, 2004, 68(1): 129–138.", "2519"),
    ("WOLCH, J.; BYRNE, J.; NEWELL, J. Urban green space, public health, and environmental justice: "
     "The challenge of making cities ‘just green enough’. LANDSCAPE AND URBAN PLANNING, 2014, 125: 234–244.", "2693"),
    ("TZOULAS, K., et al. Promoting ecosystem and human health in urban areas using green infrastructure: "
     "A literature review. LANDSCAPE AND URBAN PLANNING, 2007, 81(3): 167–178.", "2722"),
]

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
s = prs.slides.add_slide(prs.slide_layouts[6])

tf = s.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(0.7)).text_frame
tf.paragraphs[0].alignment = PP_ALIGN.LEFT
r = tf.paragraphs[0].add_run(); r.text = "Highly cited articles in Landscape and Urban Planning"
r.font.name, r.font.size, r.font.bold, r.font.color.rgb = FONT, Pt(28), True, BLUE
bar = s.shapes.add_shape(1, Inches(0.8), Inches(1.25), Inches(1.2), Inches(0.08))
bar.fill.solid(); bar.fill.fore_color.rgb = YELLOW; bar.line.fill.background(); bar.shadow.inherit = False

tbl = s.shapes.add_table(len(rows) + 1, 2, Inches(0.8), Inches(1.65), Inches(11.7), Inches(5.0)).table
tbl.columns[0].width, tbl.columns[1].width = Inches(9.9), Inches(1.8)
tbl.first_row = True
tbl.horz_banding = False


def fill(cell, text, size, color, bg, bold=False, align=PP_ALIGN.LEFT):
    cell.fill.solid(); cell.fill.fore_color.rgb = bg
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = cell.margin_right = Inches(0.15)
    cell.margin_top = cell.margin_bottom = Inches(0.08)
    tf = cell.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.name, r.font.size, r.font.bold, r.font.color.rgb = FONT, Pt(size), bold, color


tbl.rows[0].height = Inches(0.6)
fill(tbl.cell(0, 0), "Articles", 18, WHITE, BLUE, bold=True, align=PP_ALIGN.CENTER)
fill(tbl.cell(0, 1), "Citations", 18, BLUE, YELLOW, bold=True, align=PP_ALIGN.CENTER)
for i, (art, cit) in enumerate(rows, start=1):
    bg = TINT if i % 2 == 0 else WHITE
    tbl.rows[i].height = Inches(1.1)
    fill(tbl.cell(i, 0), art, 15, INK, bg)
    fill(tbl.cell(i, 1), cit, 20, BLUE, bg, bold=True, align=PP_ALIGN.CENTER)

prs.save("journal-most-cited.pptx")
print("saved")
