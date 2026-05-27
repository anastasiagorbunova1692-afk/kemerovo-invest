from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ── Palette ────────────────────────────────────────────────────────────────
DARK   = RGBColor(0x07, 0x09, 0x0F)
DARK2  = RGBColor(0x0D, 0x15, 0x26)
BLUE   = RGBColor(0x1E, 0x90, 0xFF)
VIOLET = RGBColor(0x7C, 0x5B, 0xF5)
LIGHT  = RGBColor(0xF4, 0xF6, 0xF9)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
MUTED  = RGBColor(0xA0, 0xA8, 0xB8)
GREY_CARD = RGBColor(0x14, 0x1C, 0x2E)

W = Inches(13.33)   # 16:9
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

blank = prs.slide_layouts[6]   # totally blank layout

# ── Helpers ────────────────────────────────────────────────────────────────
def add_slide():
    return prs.slides.add_slide(blank)

def bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def box(slide, x, y, w, h, color=None, alpha=None):
    """Add a filled rectangle, return shape."""
    sh = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE=1
    sh.line.fill.background()
    if color:
        sh.fill.solid()
        sh.fill.fore_color.rgb = color
    else:
        sh.fill.background()
    return sh

def txt(slide, text, x, y, w, h,
        size=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT,
        italic=False, wrap=True, font_name="Calibri"):
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font_name
    return txb

def label(slide, text, x, y, w=Inches(6), color=BLUE):
    txt(slide, text, x, y, w, Inches(0.3),
        size=9, bold=True, color=color, font_name="Calibri")

def heading(slide, text, x, y, w, size=28, color=WHITE):
    txt(slide, text, x, y, w, Inches(1.2),
        size=size, bold=True, color=color, font_name="Calibri")

def photo_placeholder(slide, x, y, w, h, filename):
    """Grey placeholder rectangle with dashed border and centered caption."""
    sh = slide.shapes.add_shape(1, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor(0x2A, 0x2A, 0x3A)
    ln = sh.line
    ln.color.rgb = RGBColor(0x55, 0x55, 0x77)
    ln.width = Pt(1)
    # dashed via XML
    ln_elem = sh.line._ln
    prstDash = etree.SubElement(ln_elem, qn('a:prstDash'))
    prstDash.set('val', 'dash')
    # caption text
    tf = sh.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = f"[ ФОТО: {filename} ]"
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x77, 0x77, 0x99)
    run.font.name = "Calibri"
    return sh

def divider(slide, x, y, w, color=BLUE):
    ln = slide.shapes.add_shape(1, x, y, w, Pt(2))
    ln.fill.solid()
    ln.fill.fore_color.rgb = color
    ln.line.fill.background()
    return ln

def card(slide, x, y, w, h, title, body, title_color=BLUE,
         bg_color=GREY_CARD, border_color=None):
    sh = box(slide, x, y, w, h, bg_color)
    if border_color:
        sh.line.color.rgb = border_color
        sh.line.width = Pt(1)
    else:
        sh.line.fill.background()
    txt(slide, title, x+Inches(0.2), y+Inches(0.15), w-Inches(0.4), Inches(0.35),
        size=9, bold=True, color=title_color, font_name="Calibri")
    txt(slide, body, x+Inches(0.2), y+Inches(0.5), w-Inches(0.4), h-Inches(0.6),
        size=10, color=MUTED, font_name="Calibri")

def metric_block(slide, num, sub, x, y, w, num_color=DARK, sub_color=RGBColor(0x88,0x88,0x88)):
    txt(slide, num, x, y, w, Inches(0.8),
        size=32, bold=True, color=num_color, align=PP_ALIGN.CENTER, font_name="Calibri")
    txt(slide, sub, x, y+Inches(0.75), w, Inches(0.4),
        size=8, color=sub_color, align=PP_ALIGN.CENTER, font_name="Calibri")

def table(slide, headers, rows, x, y, w, h,
          header_bg=DARK2, header_color=BLUE,
          row_bg=None, row_alt=None,
          text_color=WHITE, font_size=9,
          highlight_rows=None, highlight_color=BLUE):
    """Simple table via individual boxes."""
    n_cols = len(headers)
    col_w  = w / n_cols
    row_h  = h / (len(rows) + 1)

    # header
    for ci, hdr in enumerate(headers):
        sh = box(slide, x + ci*col_w, y, col_w, row_h, header_bg)
        sh.line.color.rgb = RGBColor(0x20,0x28,0x40)
        sh.line.width = Pt(0.5)
        txt(slide, hdr,
            x + ci*col_w + Inches(0.05), y + Inches(0.04),
            col_w - Inches(0.1), row_h,
            size=font_size, bold=True, color=header_color, font_name="Calibri")

    for ri, row in enumerate(rows):
        ry = y + row_h + ri*row_h
        is_highlight = highlight_rows and ri in highlight_rows
        for ci, cell in enumerate(row):
            bg_c = (highlight_color if is_highlight
                    else (row_alt if (row_alt and ri % 2 == 1) else (row_bg or DARK2)))
            sh = box(slide, x + ci*col_w, ry, col_w, row_h, bg_c)
            sh.line.color.rgb = RGBColor(0x20,0x28,0x40)
            sh.line.width = Pt(0.5)
            fc = WHITE if is_highlight else text_color
            fb = is_highlight
            txt(slide, str(cell),
                x + ci*col_w + Inches(0.05), ry + Inches(0.03),
                col_w - Inches(0.1), row_h,
                size=font_size, bold=fb, color=fc, font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — TITLE
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, DARK)

# grid overlay (subtle lines via thin boxes)
for i in range(0, 14):
    ln = s.shapes.add_shape(1, Inches(i), Inches(0), Pt(1), H)
    ln.fill.solid(); ln.fill.fore_color.rgb = RGBColor(0x12,0x14,0x1E)
    ln.line.fill.background()

# right photo placeholder
photo_placeholder(s, Inches(6.5), Inches(0), Inches(6.83), H, "building.jpg")

# left content
label(s, "Г. КЕМЕРОВО · 2025", Inches(0.6), Inches(1.2))
heading(s, "Инвестиции в спортивную\nинфраструктуру Сибири",
        Inches(0.6), Inches(1.6), Inches(5.5), size=34)
divider(s, Inches(0.6), Inches(3.0), Inches(0.7))

# 3 metrics row
for i, (num, sub) in enumerate([
    ("191 МЛН ₽", "ОБЪЁМ ПРОЕКТА"),
    ("5 ЛЕТ",     "ОКУПАЕМОСТЬ"),
    ("22% ROI",   "ДОХОДНОСТЬ"),
]):
    mx = Inches(0.5) + i * Inches(1.8)
    rect = box(s, mx, Inches(5.8), Inches(1.65), Inches(1.0), DARK2)
    rect.line.color.rgb = RGBColor(0x30,0x38,0x50)
    rect.line.width = Pt(1)
    txt(s, num, mx+Inches(0.1), Inches(5.85), Inches(1.45), Inches(0.45),
        size=14, bold=True, color=WHITE, font_name="Calibri")
    txt(s, sub, mx+Inches(0.1), Inches(6.3), Inches(1.45), Inches(0.3),
        size=7, color=MUTED, font_name="Calibri")

# logo
txt(s, "PROGROUP", Inches(0.6), Inches(7.0), Inches(2), Inches(0.35),
    size=14, bold=True, color=WHITE, font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — KEY METRICS
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, LIGHT)

label(s, "О ПРОЕКТЕ", Inches(0.6), Inches(0.4), color=BLUE)
heading(s, "Спортивно-развлекательный комплекс\nв г. Кемерово",
        Inches(0.6), Inches(0.75), Inches(12), size=28, color=DARK)

# 3 metric columns with dividers
col_w = Inches(4.1)
for i, (num, sub) in enumerate([
    ("191 000 000 ₽", "ОБЩИЙ ОБЪЁМ ИНВЕСТИЦИЙ"),
    ("22%",           "ГОДОВАЯ ДОХОДНОСТЬ"),
    ("2 500 М²",      "ПЛОЩАДЬ ЗДАНИЯ"),
]):
    cx = Inches(0.3) + i * col_w
    if i > 0:
        divider(s, cx - Inches(0.1), Inches(2.5), Pt(1.5))
        sh = box(s, cx - Inches(0.05), Inches(2.5), Pt(2), Inches(3))
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(0xCC,0xCC,0xCC)
        sh.line.fill.background()
    txt(s, num, cx + Inches(0.1), Inches(2.8), col_w - Inches(0.3), Inches(1.1),
        size=36, bold=True, color=DARK, align=PP_ALIGN.CENTER, font_name="Calibri")
    txt(s, sub, cx + Inches(0.1), Inches(4.1), col_w - Inches(0.3), Inches(0.5),
        size=9, color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER, font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — ABOUT PROJECT
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, DARK)

label(s, "О ПРОЕКТЕ", Inches(0.6), Inches(0.35))
heading(s, "Надёжность недвижимости. Ликвидность бизнеса.",
        Inches(0.6), Inches(0.75), Inches(12), size=26)

cards_data = [
    ("01  ИНВЕСТИЦИИ",
     "~80 000 руб за 1 м² при общем бюджете 191 млн ₽ — ниже рынка при высокой доходности."),
    ("02  ОКУПАЕМОСТЬ",
     "5–6 лет с ROI 22% годовых — одна из лучших доходностей в коммерческой недвижимости Сибири."),
    ("03  ГОТОВЫЙ АРЕНДАТОР",
     "Спортивно-развлекательная ниша, готовые проекты и управляющая компания."),
    ("04  ВАРИАНТЫ ВЫХОДА",
     "1) Готовый арендный бизнес  2) Отдельный бизнес — спортивный комплекс"),
]
cw = Inches(3.1); ch = Inches(2.0)
for i, (t, b) in enumerate(cards_data):
    cx = Inches(0.5) + (i % 2) * (cw + Inches(0.2))
    cy = Inches(1.8) + (i // 2) * (ch + Inches(0.2))
    card(s, cx, cy, cw, ch, t, b)

# card 5 full width
card(s, Inches(0.5), Inches(6.05), Inches(6.5), Inches(1.1),
     "05  ЛИКВИДНАЯ НЕДВИЖИМОСТЬ",
     "Покупка коммерческой недвижимости в центре города по цене 38 000 ₽/м² — в 3–4 раза ниже рынка, с доходностью от 600 ₽/м² в месяц и ежегодным ростом стоимости актива на 10–18%.")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — BUILDING
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, WHITE)

label(s, "ЗДАНИЕ", Inches(0.6), Inches(0.35), color=BLUE)
heading(s, "Общая стоимость здания — 95 000 000 ₽",
        Inches(0.6), Inches(0.75), Inches(7), size=24, color=DARK)

# left info card
card_bg = RGBColor(0xEA,0xF0,0xFA)
sh = box(s, Inches(0.5), Inches(1.7), Inches(5.8), Inches(4.5), card_bg)
sh.line.fill.background()

items = [
    "Площадь: 2 500 м²",
    "Стоимость: 38 000 руб за 1 м²",
    "Центр города, вблизи ТЦ и Леруа Мерлен",
    "Необходимое количество парковочных мест",
    "Низкая стоимость ремонта",
]
for ii, it in enumerate(items):
    txt(s, f"— {it}", Inches(0.7), Inches(1.85)+ii*Inches(0.42),
        Inches(5.4), Inches(0.4), size=11, color=DARK, font_name="Calibri")

# progress bars
prog_data = [("Вентиляция", 100), ("Освещение", 80), ("Полы", 100)]
txt(s, "ГОТОВНОСТЬ ЗДАНИЯ", Inches(0.7), Inches(4.05), Inches(3), Inches(0.3),
    size=9, bold=True, color=BLUE, font_name="Calibri")
for pi, (plabel, pval) in enumerate(prog_data):
    py = Inches(4.45) + pi * Inches(0.55)
    txt(s, plabel, Inches(0.7), py, Inches(1.5), Inches(0.3),
        size=10, color=DARK, font_name="Calibri")
    txt(s, f"{pval}%", Inches(5.4), py, Inches(0.5), Inches(0.3),
        size=10, color=DARK, font_name="Calibri")
    # track
    track = box(s, Inches(0.7), py+Inches(0.32), Inches(4.9), Inches(0.1),
                RGBColor(0xCC,0xD5,0xE5))
    track.line.fill.background()
    # fill
    fill_w = Inches(4.9 * pval / 100)
    bar = box(s, Inches(0.7), py+Inches(0.32), fill_w, Inches(0.1), BLUE)
    bar.line.fill.background()

# right photo placeholder
photo_placeholder(s, Inches(6.7), Inches(1.7), Inches(6.3), Inches(4.5), "floorplan.jpg")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — LIQUIDITY
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, LIGHT)

label(s, "ЛИКВИДНОСТЬ", Inches(0.6), Inches(0.35), color=BLUE)
heading(s, "Ликвидность объекта", Inches(0.6), Inches(0.75), Inches(12), size=26, color=DARK)

# two comparison cards
for i, (title, lines, bg_c) in enumerate([
    ("НАШЕ ПРЕДЛОЖЕНИЕ",
     ["Стоимость здания: 38 000 ₽/м²", "Доходность: от 600 ₽/м²"],
     RGBColor(0xEA,0xF0,0xFA)),
    ("РЫНОК КЕМЕРОВО",
     ["Средняя стоимость: 100 000–180 000 ₽/м²", "Доходность: 600–1 350 ₽/м²"],
     RGBColor(0xE0,0xE0,0xE0)),
]):
    cx = Inches(0.5) + i * Inches(6.4)
    sh = box(s, cx, Inches(1.7), Inches(6.0), Inches(1.5), bg_c)
    sh.line.fill.background()
    txt(s, title, cx+Inches(0.2), Inches(1.8), Inches(5.6), Inches(0.35),
        size=10, bold=True, color=DARK, font_name="Calibri")
    for li, line in enumerate(lines):
        txt(s, line, cx+Inches(0.2), Inches(2.22)+li*Inches(0.38),
            Inches(5.6), Inches(0.36), size=11, color=DARK, font_name="Calibri")

# growth table
txt(s, "Рост стоимости 1 м² коммерческой недвижимости в г. Кемерово",
    Inches(0.6), Inches(3.45), Inches(12), Inches(0.4),
    size=13, bold=True, color=DARK, font_name="Calibri")

table(s,
    ["Год", "Диапазон стоимости ₽/м²", "Рост за год"],
    [
        ["2021", "55 000 – 70 000", "+8–10%"],
        ["2022", "65 000 – 85 000", "+12–18%"],
        ["2023", "80 000 – 110 000", "+18–25%"],
        ["2024", "95 000 – 135 000", "+15–22%"],
        ["2025", "110 000 – 180 000", "+10–18%"],
    ],
    Inches(0.5), Inches(3.95), Inches(12.3), Inches(3.2),
    header_bg=DARK2, header_color=BLUE,
    row_bg=RGBColor(0xE8,0xEE,0xF8), row_alt=RGBColor(0xF4,0xF6,0xF9),
    text_color=DARK, font_size=11
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — MARKET
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, DARK)

label(s, "РЫНОК", Inches(0.6), Inches(0.35))
heading(s, "Кемерово — один из самых спортивных городов Сибири",
        Inches(0.6), Inches(0.75), Inches(12), size=24)

# 3 stats
for i, (num, desc) in enumerate([
    ("52%",     "Жителей Кузбасса регулярно занимаются спортом"),
    ("60,1%",   "Систематически занимающихся к концу 2023 года"),
    ("285 000", "Человек занимаются спортом в Кемерово"),
]):
    cx = Inches(0.4) + i * Inches(4.3)
    txt(s, num, cx, Inches(1.9), Inches(4), Inches(1.0),
        size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font_name="Calibri")
    txt(s, desc, cx, Inches(3.0), Inches(4), Inches(0.8),
        size=10, color=MUTED, align=PP_ALIGN.CENTER, font_name="Calibri")

# two info blocks
divider(s, Inches(0.5), Inches(4.0), Inches(12.3))

for i, (title, items) in enumerate([
    ("ФОРМАТЫ СПОРТИВНЫХ ОБЪЕКТОВ ГОРОДА",
     ["• Государственные спортивные объекты",
      "• Фитнес-центры",
      "• Ледовые арены",
      "• Базовые спортивные комплексы"]),
    ("ЧТО ОТСУТСТВУЕТ НА РЫНКЕ",
     ["• Точки притяжения для активной аудитории",
      "• Крытые развлечения нового поколения",
      "• Комплексы для семейного досуга"]),
]):
    cx = Inches(0.5) + i * Inches(6.5)
    txt(s, title, cx, Inches(4.2), Inches(6.2), Inches(0.4),
        size=10, bold=True, color=BLUE, font_name="Calibri")
    for bi, item in enumerate(items):
        txt(s, item, cx, Inches(4.7)+bi*Inches(0.5),
            Inches(6.2), Inches(0.45), size=11, color=MUTED, font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — AUDIENCE
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, WHITE)

label(s, "АУДИТОРИЯ", Inches(0.6), Inches(0.35), color=BLUE)
heading(s, "Проект для тех, кто выбирает активную жизнь",
        Inches(0.6), Inches(0.75), Inches(12), size=26, color=DARK)

aud_data = [
    ("МОЛОДЁЖЬ 18–35 ЛЕТ",
     "Основная аудитория social sports. Выбирают эмоциональный досуг, комьюнити и впечатления. Активные пользователи соцсетей."),
    ("СЕМЬИ С ДЕТЬМИ",
     "Детские школы · Семейный отдых · Совместные активности. Высокий LTV — регулярные повторные визиты и абонементы."),
    ("КОРПОРАТИВНЫЙ СЕГМЕНТ",
     "Компании тратят ~500–700 млн руб/год на тимбилдинги. Государственные компании возмещают до 70% трат на корпоративный спорт."),
]
cw = Inches(4.0); ch = Inches(3.8)
for i, (t, b) in enumerate(aud_data):
    cx = Inches(0.4) + i * Inches(4.3)
    sh = box(s, cx, Inches(1.7), cw, ch, RGBColor(0xF0,0xF4,0xFA))
    sh.line.color.rgb = RGBColor(0xCC,0xD5,0xE5)
    sh.line.width = Pt(1)
    divider(s, cx + Inches(0.2), Inches(1.7), Inches(0.5))
    txt(s, t, cx+Inches(0.2), Inches(1.9), cw-Inches(0.4), Inches(0.55),
        size=12, bold=True, color=DARK, font_name="Calibri")
    txt(s, b, cx+Inches(0.2), Inches(2.55), cw-Inches(0.4), Inches(2.7),
        size=11, color=RGBColor(0x44,0x44,0x55), font_name="Calibri")

# quote
txt(s, "«Несмотря на высокий интерес к спорту, в городе практически отсутствуют современные пространства.»",
    Inches(0.6), Inches(5.8), Inches(12), Inches(0.9),
    size=12, italic=True, color=RGBColor(0x55,0x55,0x66), font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — BEACH SPORTS
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, WHITE)

label(s, "ПЛЯЖНЫЙ СПОРТ", Inches(0.6), Inches(0.3), color=BLUE)
heading(s, "Центр пляжных видов спорта",
        Inches(0.6), Inches(0.7), Inches(7), size=26, color=DARK)

# photo placeholder
photo_placeholder(s, Inches(7.0), Inches(0.3), Inches(6.0), Inches(4.0), "beach-1.jpg")

# stats
for i, (num, sub) in enumerate([("5 КОРТОВ", "с беспыльным песком"),
                                  ("1 300 М²", "площадь здания")]):
    cx = Inches(0.5) + i * Inches(3.2)
    txt(s, num, cx, Inches(1.8), Inches(3.0), Inches(0.8),
        size=30, bold=True, color=DARK, font_name="Calibri")
    txt(s, sub, cx, Inches(2.6), Inches(3.0), Inches(0.35),
        size=11, color=RGBColor(0x77,0x77,0x88), font_name="Calibri")

divider(s, Inches(0.5), Inches(3.1), Inches(6.2))

# technologies
txt(s, "ТЕХНОЛОГИИ", Inches(0.5), Inches(3.25), Inches(6), Inches(0.3),
    size=9, bold=True, color=BLUE, font_name="Calibri")
for ti, tech in enumerate(["● Умные раздевалки", "● Face ID вход", "● Безключевой доступ"]):
    txt(s, tech, Inches(0.5) + ti * Inches(2.0), Inches(3.6),
        Inches(1.9), Inches(0.4), size=11, color=DARK, font_name="Calibri")

# CAPEX summary
sh = box(s, Inches(0.5), Inches(4.3), Inches(6.2), Inches(2.9), DARK2)
sh.line.fill.background()
txt(s, "CAPEX", Inches(0.7), Inches(4.45), Inches(3), Inches(0.3),
    size=9, bold=True, color=BLUE, font_name="Calibri")
txt(s, "30 000 000 ₽", Inches(0.7), Inches(4.8), Inches(5.8), Inches(0.7),
    size=28, bold=True, color=WHITE, font_name="Calibri")
txt(s, "Возврат инвестиций: 3–4 года", Inches(0.7), Inches(5.5), Inches(5.8), Inches(0.35),
    size=11, color=MUTED, font_name="Calibri")

for ti, tech in enumerate(["Умные раздевалки", "Face ID вход", "Безключевой доступ"]):
    pass  # already above

txt(s, "Структура дохода: аренда кортов · event · детские тренировки · турниры",
    Inches(0.7), Inches(5.95), Inches(6.0), Inches(0.5),
    size=10, color=MUTED, font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — BEACH FINANCES
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, DARK)

label(s, "ПЛЯЖНЫЙ СПОРТ — ФИНАНСЫ", Inches(0.6), Inches(0.35))
heading(s, "Финансовая модель — Центр пляжного спорта",
        Inches(0.6), Inches(0.75), Inches(12), size=24)

# CAPEX
sh = box(s, Inches(0.5), Inches(1.7), Inches(5.8), Inches(1.4), DARK2)
sh.line.color.rgb = BLUE; sh.line.width = Pt(1)
txt(s, "CAPEX", Inches(0.7), Inches(1.8), Inches(3), Inches(0.3),
    size=9, bold=True, color=BLUE, font_name="Calibri")
txt(s, "30 000 000 ₽", Inches(0.7), Inches(2.1), Inches(5.4), Inches(0.65),
    size=30, bold=True, color=WHITE, font_name="Calibri")
txt(s, "Возврат инвестиций: 3–4 года", Inches(0.7), Inches(2.8), Inches(5.4), Inches(0.3),
    size=11, color=MUTED, font_name="Calibri")

# income structure bars
txt(s, "СТРУКТУРА ДОХОДА", Inches(0.6), Inches(3.4), Inches(8), Inches(0.35),
    size=9, bold=True, color=BLUE, font_name="Calibri")

bars = [("Аренда кортов", 55), ("Event-направление", 20),
        ("Детские тренировки", 15), ("Турниры", 10)]
bar_w = Inches(7.0)
for bi, (blabel, bval) in enumerate(bars):
    by = Inches(3.85) + bi * Inches(0.72)
    txt(s, blabel, Inches(0.6), by, Inches(2.8), Inches(0.3),
        size=10, color=WHITE, font_name="Calibri")
    txt(s, f"{bval}%", Inches(10.8), by, Inches(0.6), Inches(0.3),
        size=10, color=BLUE, font_name="Calibri")
    # track
    track = box(s, Inches(3.5), by+Inches(0.06), bar_w, Inches(0.2),
                RGBColor(0x1A,0x22,0x38))
    track.line.fill.background()
    # fill
    fill = box(s, Inches(3.5), by+Inches(0.06), bar_w * bval/100, Inches(0.2), BLUE)
    fill.line.fill.background()

# KPIs
kpis = ["Выручка", "Кол-во клиентов", "LTV", "Загрузка кортов"]
for ki, kpi in enumerate(kpis):
    kx = Inches(0.5) + ki * Inches(3.0)
    sh = box(s, kx, Inches(6.7), Inches(2.7), Inches(0.55),
             RGBColor(0x14,0x1C,0x2E))
    sh.line.fill.background()
    txt(s, kpi, kx+Inches(0.1), Inches(6.75), Inches(2.5), Inches(0.4),
        size=10, bold=True, color=MUTED, align=PP_ALIGN.CENTER, font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — BEACH CAPEX TABLE
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, DARK)

label(s, "ПЛЯЖНЫЙ СПОРТ — CAPEX", Inches(0.6), Inches(0.25))
heading(s, "Инвестиции CAPEX — Центр пляжного спорта",
        Inches(0.6), Inches(0.6), Inches(12), size=22)

capex_beach = [
    ["Помещение", "Ремонт", "11 068 211"],
    ["Помещение", "Песок", "5 871 200"],
    ["Помещение", "Свет", "2 570 600"],
    ["Помещение", "Коммуникации", "1 867 000"],
    ["Помещение", "Камеры и контроль доступа", "935 000"],
    ["Мебель и техника", "Холл + Раздевалки + Ресепшн", "3 129 465"],
    ["Маркетинг", "Сайт, реклама, прочее", "1 250 000"],
    ["ЗП до открытия", "Администраторы, дизайнер, маркетолог, СММ, прочее", "1 259 500"],
    ["Прочее", "Инвентарь, ИС, мерч, прочее", "1 336 183"],
    ["ИТОГО", "", "29 717 159 ₽"],
]
table(s,
    ["Категория", "Подкатегория", "Сумма, руб"],
    capex_beach,
    Inches(0.4), Inches(1.4), Inches(12.5), Inches(5.8),
    header_bg=DARK2, header_color=BLUE,
    row_bg=RGBColor(0x0D,0x15,0x26), row_alt=RGBColor(0x10,0x18,0x2C),
    text_color=WHITE, font_size=10,
    highlight_rows=[9], highlight_color=RGBColor(0x1A,0x3A,0x6A)
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — KARTING
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, DARK)

label(s, "КАРТИНГ", Inches(0.6), Inches(0.3), color=VIOLET)
heading(s, "Картинг с многоуровневой трассой",
        Inches(0.6), Inches(0.7), Inches(7), size=26, color=VIOLET)

photo_placeholder(s, Inches(7.0), Inches(0.3), Inches(6.0), Inches(4.2), "karting-1.jpg")

txt(s, "Высокотехнологичные многоуровневые трассы. Скоростные заезды\nна экологичных электрокартах с системой хронометража.",
    Inches(0.6), Inches(1.7), Inches(6.2), Inches(0.8),
    size=12, color=MUTED, font_name="Calibri")

features = [
    ("СКОРОСТЬ И ДРАЙВ",  "Разгон до 100 км/ч за 4 секунды — ощущения настоящего спорта."),
    ("НАДЁЖНОСТЬ",        "Карты производятся в России из качественных комплектующих."),
    ("БЫСТРОЕ ОБСЛУЖИВАНИЕ", "Простая система зарядки и диагностики, минимум простоев."),
    ("ЭКОЛОГИЧНОСТЬ",     "Без выхлопных газов. Комфортно для помещений, безопасно для детей."),
    ("ТЕЛЕМЕТРИЯ",        "Дистанционное ограничение скорости, хронометраж, детские модели."),
]
for fi, (ft, fb) in enumerate(features):
    fy = Inches(2.65) + fi * Inches(0.88)
    sh = box(s, Inches(0.5), fy, Inches(6.2), Inches(0.8), DARK2)
    sh.line.color.rgb = VIOLET; sh.line.width = Pt(1)
    txt(s, ft, Inches(0.7), fy+Inches(0.07), Inches(5.8), Inches(0.28),
        size=9, bold=True, color=VIOLET, font_name="Calibri")
    txt(s, fb, Inches(0.7), fy+Inches(0.38), Inches(5.8), Inches(0.35),
        size=10, color=MUTED, font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 12 — KARTING FINANCES
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, DARK)

label(s, "КАРТИНГ — ФИНАНСЫ", Inches(0.6), Inches(0.35), color=VIOLET)
heading(s, "Финансовая модель — Картинг",
        Inches(0.6), Inches(0.75), Inches(12), size=24, color=VIOLET)

sh = box(s, Inches(0.5), Inches(1.7), Inches(5.8), Inches(1.4), DARK2)
sh.line.color.rgb = VIOLET; sh.line.width = Pt(1)
txt(s, "CAPEX", Inches(0.7), Inches(1.8), Inches(3), Inches(0.3),
    size=9, bold=True, color=VIOLET, font_name="Calibri")
txt(s, "66 000 000 ₽", Inches(0.7), Inches(2.1), Inches(5.4), Inches(0.65),
    size=30, bold=True, color=WHITE, font_name="Calibri")
txt(s, "Возврат инвестиций: 5–6 лет", Inches(0.7), Inches(2.8), Inches(5.4), Inches(0.3),
    size=11, color=MUTED, font_name="Calibri")

# donut legend as horizontal bars
txt(s, "СТРУКТУРА ДОХОДА", Inches(0.6), Inches(3.4), Inches(8), Inches(0.35),
    size=9, bold=True, color=VIOLET, font_name="Calibri")

donut_data = [("Заезды", 70, VIOLET), ("Event", 20, RGBColor(0xA4,0x7F,0xFF)),
              ("Детские тренировки", 10, RGBColor(0xC9,0xAF,0xFF))]
bar_w = Inches(7.0)
for di, (dlabel, dval, dcol) in enumerate(donut_data):
    dy = Inches(3.85) + di * Inches(0.72)
    txt(s, f"{dlabel}  {dval}%", Inches(0.6), dy, Inches(2.8), Inches(0.3),
        size=10, color=WHITE, font_name="Calibri")
    track = box(s, Inches(3.5), dy+Inches(0.06), bar_w, Inches(0.2),
                RGBColor(0x1A,0x22,0x38))
    track.line.fill.background()
    fill = box(s, Inches(3.5), dy+Inches(0.06), bar_w * dval/100, Inches(0.2), dcol)
    fill.line.fill.background()

kpis = ["Выручка", "Кол-во клиентов", "LTV", "Загрузка трассы"]
for ki, kpi in enumerate(kpis):
    kx = Inches(0.5) + ki * Inches(3.0)
    sh = box(s, kx, Inches(6.7), Inches(2.7), Inches(0.55),
             RGBColor(0x14,0x1C,0x2E))
    sh.line.fill.background()
    txt(s, kpi, kx+Inches(0.1), Inches(6.75), Inches(2.5), Inches(0.4),
        size=10, bold=True, color=MUTED, align=PP_ALIGN.CENTER, font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 13 — KARTING CAPEX TABLE
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, DARK)

label(s, "КАРТИНГ — CAPEX", Inches(0.6), Inches(0.25), color=VIOLET)
heading(s, "Инвестиции CAPEX — Картинг",
        Inches(0.6), Inches(0.6), Inches(12), size=22, color=VIOLET)

capex_kart = [
    ["Осн. оборудование", "Машины (электрокарты)", "24 800 000"],
    ["Помещение", "Проектирование", "580 000"],
    ["Помещение", "Ремонт + трасса + свет + коммуникации + камеры", "36 700 000*"],
    ["Маркетинг", "Сайт, реклама, прочее", "1 250 000"],
    ["ЗП до открытия", "Администраторы, дизайнер, механик, СММ, прочее", "1 440 500"],
    ["Прочее", "Инвентарь, ИС, мерч, прочее", "1 227 500"],
    ["ИТОГО", "", "66 000 000 ₽"],
]
table(s,
    ["Категория", "Подкатегория", "Сумма, руб"],
    capex_kart,
    Inches(0.4), Inches(1.4), Inches(12.5), Inches(4.5),
    header_bg=DARK2, header_color=VIOLET,
    row_bg=RGBColor(0x0D,0x15,0x26), row_alt=RGBColor(0x10,0x18,0x2C),
    text_color=WHITE, font_size=10,
    highlight_rows=[6], highlight_color=RGBColor(0x28,0x1A,0x4A)
)

txt(s, "*Точная стоимость после проектирования",
    Inches(0.5), Inches(6.1), Inches(8), Inches(0.3),
    size=9, color=MUTED, font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 14 — FINANCIAL MODEL
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, WHITE)

label(s, "ФИНАНСЫ", Inches(0.6), Inches(0.25), color=BLUE)
heading(s, "Финансовая модель комплекса на год",
        Inches(0.6), Inches(0.6), Inches(12), size=24, color=DARK)

fin_rows = [
    ["ВЫРУЧКА ЗА ГОД", "84 397 000 ₽"],
    ["   — Песок", "29 480 000 ₽"],
    ["   — Картинг", "50 825 000 ₽"],
    ["   — Субаренда", "1 932 000 ₽"],
    ["   — Дополнительные услуги", "2 160 000 ₽"],
    ["РАСХОДЫ", "32 680 400 ₽"],
    ["   — Песок", "13 218 400 ₽"],
    ["   — Картинг", "19 462 000 ₽"],
    ["ОПЕРАЦИОННАЯ ПРИБЫЛЬ", "51 716 600 ₽"],
    ["РАСХОДЫ УК", "7 571 660 ₽"],
    ["EBITDA", "44 144 940 ₽"],
    ["   — Эквайринг + Налоги", "2 068 664 ₽"],
    ["ЧИСТАЯ ПРИБЫЛЬ", "42 076 276 ₽"],
]
col_w = [Inches(8.5), Inches(3.8)]
row_h = Inches(0.45)
ty = Inches(1.5)
for ri, row in enumerate(fin_rows):
    is_hl = row[0] in ("EBITDA", "ЧИСТАЯ ПРИБЫЛЬ")
    is_main = row[0] in ("ВЫРУЧКА ЗА ГОД", "РАСХОДЫ", "ОПЕРАЦИОННАЯ ПРИБЫЛЬ", "РАСХОДЫ УК")
    is_indent = row[0].startswith("   ")
    bg_c = (RGBColor(0xE0,0xEE,0xFF) if is_hl
            else (RGBColor(0xF4,0xF6,0xF9) if ri % 2 == 0 else WHITE))
    for ci, (cell, cw) in enumerate(zip(row, col_w)):
        cx = Inches(0.5) if ci == 0 else Inches(9.0)
        sh = box(s, cx, ty + ri*row_h, cw, row_h, bg_c)
        sh.line.color.rgb = RGBColor(0xDD,0xDD,0xDD)
        sh.line.width = Pt(0.5)
        fc = (BLUE if is_hl else DARK)
        fb = is_hl or is_main
        txt(s, str(cell), cx+Inches(0.1), ty+ri*row_h+Inches(0.04),
            cw-Inches(0.15), row_h,
            size=10, bold=fb, color=fc,
            align=(PP_ALIGN.RIGHT if ci == 1 else PP_ALIGN.LEFT),
            font_name="Calibri")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 15 — INVESTOR OFFER
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, DARK)

label(s, "ИНВЕСТОРАМ", Inches(0.6), Inches(0.25))
heading(s, "Предложение для инвесторов", Inches(0.6), Inches(0.6), Inches(12), size=26)

# investments table (left)
inv_rows = [
    ["Общая сумма вложений", "191 000 000 ₽"],
    ["  — Покупка здания", "95 000 000 ₽"],
    ["  — Центр пляжного спорта", "30 000 000 ₽"],
    ["  — Электрокартинг", "66 000 000 ₽"],
    ["Минимальный вход", "20 000 000 ₽ = 10% доли"],
    ["Возврат инвестиций", "60 месяцев"],
]
col_w2 = [Inches(3.5), Inches(2.7)]
rh2 = Inches(0.45)
for ri, row in enumerate(inv_rows):
    is_sub = row[0].startswith("  ")
    bg_c = RGBColor(0x10,0x18,0x2C) if ri % 2 == 0 else DARK2
    for ci, (cell, cw) in enumerate(zip(row, col_w2)):
        cx = Inches(0.4) if ci == 0 else Inches(3.9)
        sh = box(s, cx, Inches(1.6)+ri*rh2, cw, rh2, bg_c)
        sh.line.color.rgb = RGBColor(0x20,0x28,0x40)
        sh.line.width = Pt(0.5)
        fc = (MUTED if is_sub else WHITE)
        txt(s, str(cell), cx+Inches(0.1), Inches(1.6)+ri*rh2+Inches(0.04),
            cw-Inches(0.15), rh2, size=10, color=fc,
            align=(PP_ALIGN.RIGHT if ci == 1 else PP_ALIGN.LEFT),
            bold=(not is_sub), font_name="Calibri")

# ROI table (right)
txt(s, "ROI ПО ГОДАМ", Inches(6.9), Inches(1.5), Inches(6), Inches(0.35),
    size=9, bold=True, color=BLUE, font_name="Calibri")
table(s,
    ["", "2025", "1 год", "2 год", "3 год", "4 год", "5 год"],
    [
        ["Чистая прибыль", "−210 381 380 ₽", "42 076 276 ₽", "42 076 276 ₽",
         "42 076 276 ₽", "42 076 276 ₽", "42 076 276 ₽"],
        ["ROI", "−%", "22,03%", "22,03%", "22,03%", "22,03%", "22,03%"],
    ],
    Inches(6.8), Inches(1.9), Inches(6.1), Inches(1.3),
    header_bg=DARK2, header_color=BLUE,
    row_bg=RGBColor(0x0D,0x15,0x26), row_alt=RGBColor(0x10,0x18,0x2C),
    text_color=WHITE, font_size=9
)

# Capitalization
txt(s, "КАПИТАЛИЗАЦИЯ ПРОЕКТА 2026–2031", Inches(0.4), Inches(4.55), Inches(12), Inches(0.35),
    size=9, bold=True, color=BLUE, font_name="Calibri")
table(s,
    ["Год", "Ликвидность +12%/год", "Доходы +5%/год", "ИТОГО"],
    [
        ["2026", "95 000 000", "42 076 276", "227 645 460"],
        ["2027", "106 400 000", "44 180 090", "245 677 733"],
        ["2028", "119 168 000", "46 389 094", "265 409 620"],
        ["2029", "133 468 160", "48 708 549", "287 021 861"],
        ["2030", "149 484 339", "51 143 976", "310 715 725"],
        ["2031", "167 422 460", "53 701 175", "336 715 415"],
    ],
    Inches(0.4), Inches(5.0), Inches(12.5), Inches(2.25),
    header_bg=DARK2, header_color=BLUE,
    row_bg=RGBColor(0x0D,0x15,0x26), row_alt=RGBColor(0x10,0x18,0x2C),
    text_color=WHITE, font_size=9,
    highlight_rows=[5], highlight_color=RGBColor(0x1A,0x3A,0x6A)
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 16 — CONTACT
# ══════════════════════════════════════════════════════════════════════════════
s = add_slide()
bg(s, DARK)

# left photo placeholder
photo_placeholder(s, Inches(0), Inches(0), Inches(4.5), H, "anton.jpg")

# right content
label(s, "КОНТАКТ", Inches(5.0), Inches(1.2))
txt(s, "ЯНКЕВИЧ\nАНТОН ПАВЛОВИЧ",
    Inches(5.0), Inches(1.6), Inches(7.8), Inches(1.4),
    size=28, bold=True, color=WHITE, font_name="Calibri")

# divider line
sh = box(s, Inches(5.0), Inches(3.2), Inches(7.5), Pt(1.5), WHITE)
sh.line.fill.background()

contacts = [
    ("Email",    "antonyankevich@gmail.com"),
    ("Телефон",  "8 (933) 333-43-99"),
    ("Telegram", "@ToniF03"),
]
for ci, (clabel, cval) in enumerate(contacts):
    cy = Inches(3.45) + ci * Inches(0.75)
    txt(s, clabel, Inches(5.0), cy, Inches(1.2), Inches(0.4),
        size=9, bold=True, color=BLUE, font_name="Calibri")
    txt(s, cval, Inches(6.3), cy, Inches(6.2), Inches(0.4),
        size=14, color=MUTED, font_name="Calibri")

# CTA button
sh = box(s, Inches(5.0), Inches(5.8), Inches(3.2), Inches(0.6), BLUE)
sh.line.fill.background()
txt(s, "НАПИСАТЬ В TELEGRAM →",
    Inches(5.0), Inches(5.83), Inches(3.2), Inches(0.5),
    size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font_name="Calibri")

# footer logo
txt(s, "PROGROUP  ·  © 2025",
    Inches(5.0), Inches(7.0), Inches(7.5), Inches(0.35),
    size=11, bold=True, color=MUTED, font_name="Calibri")

# ── Save ───────────────────────────────────────────────────────────────────
prs.save("/home/user/kemerovo-invest/presentation.pptx")
print(f"Done — {len(prs.slides)} slides saved.")
