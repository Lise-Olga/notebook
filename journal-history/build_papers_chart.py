"""Editable PowerPoint column chart: papers published per year in Landscape and Urban Planning."""
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

BLUE = RGBColor(0x00, 0x33, 0xA0)      # SDSU Jackrabbit Blue
PARTIAL = RGBColor(0xA9, 0xB8, 0xDD)   # light blue for the unfinished year
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x5F, 0x63, 0x66)
GRID = RGBColor(0xE3, 0xE5, 0xE8)

data = {1986: 39, 1987: 45, 1988: 60, 1989: 46, 1990: 40, 1991: 65, 1992: 51, 1993: 102,
        1994: 56, 1995: 88, 1996: 76, 1997: 80, 1998: 95, 1999: 58, 2000: 115, 2001: 107,
        2002: 75, 2003: 74, 2004: 93, 2005: 95, 2006: 124, 2007: 164, 2008: 120, 2009: 123,
        2010: 132, 2011: 194, 2012: 179, 2013: 156, 2014: 185, 2015: 188, 2016: 140,
        2017: 252, 2018: 216, 2019: 242, 2020: 183, 2021: 208, 2022: 275, 2023: 194,
        2024: 243, 2025: 193, 2026: 44}
years = list(data)

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
s = prs.slides.add_slide(prs.slide_layouts[6])


def text(x, y, w, h, t, size, color, bold=False, align=PP_ALIGN.LEFT):
    tf = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    tf.word_wrap = True
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = t
    r.font.name, r.font.size, r.font.color.rgb, r.font.bold = "Arial", Pt(size), color, bold


text(0.8, 0.45, 11.7, 0.6, "Papers published per year in Landscape and Urban Planning", 28, BLUE, bold=True)
text(0.8, 1.05, 11.7, 0.4, "Yearly output grew about sevenfold, from 39 papers in 1986 to a peak of 275 in 2022", 16, MUTED)

cd = CategoryChartData()
cd.categories = [str(y) for y in years]
cd.add_series("Papers published", [data[y] for y in years])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.6),
                        Inches(12.1), Inches(5.2), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = False
ch.font.name, ch.font.size, ch.font.color.rgb = "Arial", Pt(11), MUTED

plot = ch.plots[0]
plot.gap_width = 35
ser = plot.series[0]
ser.format.fill.solid(); ser.format.fill.fore_color.rgb = BLUE
ser.format.line.fill.background()

va = ch.value_axis
va.maximum_scale, va.minimum_scale, va.major_unit = 300, 0, 50
va.has_major_gridlines = True
va.major_gridlines.format.line.color.rgb = GRID
va.major_gridlines.format.line.width = Pt(0.75)
va.format.line.fill.background()
va.tick_labels.font.size = Pt(11)
va.has_title = True
va.axis_title.text_frame.text = "Papers"
va.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
va.axis_title.text_frame.paragraphs[0].runs[0].font.color.rgb = MUTED

ca = ch.category_axis
ca.format.line.color.rgb = RGBColor(0x9E, 0xA2, 0xA2)
ca.tick_labels.font.size = Pt(10)
ca.tick_labels.offset = 50

# 2026 is a partial year: lighter fill.
last = ser.points[len(years) - 1]
last.format.fill.solid(); last.format.fill.fore_color.rgb = PARTIAL

# Selective direct labels: first year, peak, partial year.
for y, label in [(1986, "39"), (2022, "275 (peak)"), (2026, "44*")]:
    dl = ser.points[years.index(y)].data_label
    dl.has_text_frame = True
    dl.text_frame.text = label
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    f = dl.text_frame.paragraphs[0].runs[0].font
    f.size, f.bold, f.color.rgb = Pt(12), True, INK

text(0.8, 6.9, 11.7, 0.35, "* 2026 is a partial year (data so far).", 11, MUTED)
prs.save("journal-papers-per-year.pptx")
print("saved")
