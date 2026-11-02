"""Build the VIK presentation from the VIK template (Davydova sample deck)."""
import copy
import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION
from pptx.oxml.ns import qn

TPL, MAP, OUT = sys.argv[1], sys.argv[2], sys.argv[3]

NAVY = RGBColor(0x1A, 0x46, 0x8D)
DARK = RGBColor(0x0E, 0x1E, 0x3F)
TEXT = RGBColor(0x1F, 0x29, 0x37)
MUTED = RGBColor(0x5B, 0x66, 0x75)
TINT = RGBColor(0xEA, 0xF0, 0xF8)
LIGHT = RGBColor(0x2F, 0x80, 0xC9)      # light / lighting
LIGHT_TINT = RGBColor(0xE3, 0xF0, 0xFB)
HEAT = RGBColor(0xE0, 0x6A, 0x1B)       # heat / heating
HEAT_TINT = RGBColor(0xFD, 0xEE, 0xE2)
GREY = RGBColor(0x9A, 0xA3, 0xAE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"

prs = Presentation(TPL)
LAYOUT_WHITE = prs.slide_layouts[1]

# ---------------------------------------------------------------- helpers
def set_runs(shape, text):
    """Replace text of a shape, keeping the formatting of its first run."""
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    r0 = p0.runs[0]
    r0.text = text
    for r in p0.runs[1:]:
        r._r.getparent().remove(r._r)
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)


def tb(slide, x, y, w, h, paras, size=14, color=TEXT, bold=False, align=PP_ALIGN.LEFT,
       anchor=MSO_ANCHOR.TOP, space_after=4, margin=0.05, name=None, bullets=False):
    """paras: list of str or list of runs [(text, {opts})] ; a para may be dict with keys."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(margin))
    first = True
    for para in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(space_after)
        if bullets:
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", "228600"); pPr.set("indent", "-228600")
            bu = pPr.makeelement(qn("a:buChar"), {"char": "•"})
            pPr.append(bu)
        runs = para if isinstance(para, list) else [(para, {})]
        for text, o in runs:
            r = p.add_run()
            r.text = text
            f = r.font
            f.name = FONT
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", False)
            f.color.rgb = o.get("color", color)
            if o.get("baseline"):
                r._r.get_or_add_rPr().set("baseline", o["baseline"])
    return box


def card(slide, x, y, w, h, fill=TINT, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06, name=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        s.name = name
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1)
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    s.text_frame.text = ""
    return s


def circle_num(slide, x, y, d, n, fill=NAVY, size=14):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = fill
    c.line.fill.background(); c.shadow.inherit = False
    tf = c.text_frame
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", 0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = str(n)
    r.font.name = FONT; r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = WHITE
    return c


def arrow(slide, x, y, w, h, color=GREY, shape=MSO_SHAPE.RIGHT_ARROW):
    a = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    a.fill.solid(); a.fill.fore_color.rgb = color
    a.line.fill.background(); a.shadow.inherit = False
    return a


def new_slide(title):
    s = prs.slides.add_slide(LAYOUT_WHITE)
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx == 0:
            ph.text_frame.text = title
            ph.left = Inches(0.57); ph.top = Inches(0.0); ph.width = Inches(8.6); ph.height = Inches(1.0)
        else:
            ph._element.getparent().remove(ph._element)
    return s


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def style_chart(chart, size=12):
    chart.font.name = FONT
    chart.font.size = Pt(size)
    chart.font.color.rgb = TEXT


def table(slide, x, y, w, rows, col_w, row_h=0.36, size=12, header_fill=NAVY, zebra=True,
          fills=None, bold_first_col=False, aligns=None):
    nrows, ncols = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nrows, ncols, Inches(x), Inches(y), Inches(w), Inches(row_h * nrows))
    t = gt.table
    # drop template table style banding look
    tblPr = t._tbl.tblPr
    tblPr.set("firstRow", "1"); tblPr.set("bandRow", "0")
    for j, cw in enumerate(col_w):
        t.columns[j].width = Inches(cw)
    for i in range(nrows):
        t.rows[i].height = Inches(row_h)
        for j in range(ncols):
            cell = t.cell(i, j)
            val = rows[i][j]
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            runs = val if isinstance(val, list) else [(val, {})]
            p = tf.paragraphs[0]
            p.alignment = (aligns[j] if aligns else (PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER))
            for text, o in runs:
                r = p.add_run(); r.text = text
                r.font.name = FONT; r.font.size = Pt(o.get("size", size))
                r.font.bold = o.get("bold", i == 0 or (bold_first_col and j == 0))
                r.font.color.rgb = WHITE if i == 0 else o.get("color", TEXT)
                if o.get("baseline"):
                    r._r.get_or_add_rPr().set("baseline", o["baseline"])
            cell.fill.solid()
            if i == 0:
                cell.fill.fore_color.rgb = header_fill
            elif fills and fills.get((i, j)) is not None:
                cell.fill.fore_color.rgb = fills[(i, j)]
            else:
                cell.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFB) if (zebra and i % 2 == 0) else WHITE
    return gt


def sub(t):
    return (t, {"baseline": "-25000"})


# ------------------------------------------------ slide 1 (template title)
s1 = prs.slides[0]
PROJECT = ("Методика проектирования систем досветки и отопления высокоширотных "
           "теплиц с фотоэлектрическими модулями на основе многолетних климатических данных")
for sh in s1.shapes:
    if not sh.has_text_frame:
        continue
    t = sh.text_frame.text
    if t.startswith("Мобильный"):
        set_runs(sh, PROJECT)
        sh.width = Inches(7.0)
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(24)
    elif t.startswith("Давыдова"):
        set_runs(sh, "Дэн Юаньбяо")
    elif t.startswith("Курс"):
        set_runs(sh, "Курс обучения: 3")
    elif t.startswith("Уровень"):
        set_runs(sh, "Уровень образования: аспирантура")
    elif t.startswith("Здоровье"):
        # epigraph: two paragraphs
        tf = sh.text_frame
        ps = tf.paragraphs
        ps[0].runs[0].text = "Свет и тепло — главные ресурсы северной теплицы."
        for r in ps[0].runs[1:]:
            r._r.getparent().remove(r._r)
        if len(ps) > 1:
            ps[1].runs[0].text = "Проектировать нужно на суровую зиму, а не на среднюю."
            for r in ps[1].runs[1:]:
                r._r.getparent().remove(r._r)
        for p_ in tf.paragraphs:
            for r_ in p_.runs:
                r_.font.size = Pt(18)
notes(s1, "Титульный слайд. Представление автора и темы проекта.")

# ------------------------------------------------ remove sample slides 2..10, keep 1 and 11
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
last_id = ids[-1]
for sid in ids[1:-1]:
    prs.part.drop_rel(sid.rId)
    sldIdLst.remove(sid)
from pptx.opc.packuri import PackURI
_last_slide = prs.part.related_part(last_id.rId)
_last_slide.partname = PackURI('/ppt/slides/slide200.xml')
if _last_slide.has_notes_slide:
    _last_slide.notes_slide.part.partname = PackURI('/ppt/notesSlides/notesSlide200.xml')

# ================================================= slide 2: актуальность
s = new_slide("Актуальность разработки и исследований")
tb(s, 0.45, 1.12, 6.4, 2.75, [
    [("Теплицы позволяют выращивать овощи рядом с потребителем", {"bold": True}),
     (" — основа продовольственной безопасности северных и арктических территорий.", {})],
    [("Севернее 60° с. ш. ", {"bold": True}),
     ("досветка и отопление нужны не менее полугода и составляют основную часть энергозатрат теплицы; "
      "свет и тепло связаны: лампы греют, а ФЭ-модули на кровле затеняют растения.", {})],
    [("Мощность светильников, источник тепла и электрическое присоединение ", {"bold": True}),
     ("выбираются на стадии проекта — до того, как известны будущие зимы.", {})],
], size=15, space_after=9)
pic = s.shapes.add_picture(MAP, Inches(0.55), Inches(3.95), width=Inches(5.9))
tb(s, 0.45, 6.6, 6.4, 0.38, [
    "Исследуемые города 60–70° с. ш.: LER — Леруик, REY — Рейкьявик, TRO — Тромсё, NOR — Норильск, "
    "SPE — Санкт-Петербург, YAK — Якутск; TOR — станция валидации Тыравере"], size=10, color=MUTED)

stats = [
    ("5,5–7,3 мес.", "длится период, когда естественного света теплице не хватает (целевой DLI 17 моль·м⁻²·сут⁻¹)", LIGHT, LIGHT_TINT),
    ("238–343 кВт·ч/м²", "электроэнергия на досветку за сезон в расчётную зиму", LIGHT, LIGHT_TINT),
    ("435–2017 кВт·ч/м²", "потребность в тепле за сезон в расчётную зиму (от Леруика до Якутска)", HEAT, HEAT_TINT),
]
y = 1.15
for big, small, col, tint in stats:
    card(s, 7.15, y, 5.75, 1.75, fill=tint)
    tb(s, 7.4, y + 0.15, 5.3, 0.75, [big], size=34, bold=True, color=col)
    tb(s, 7.4, y + 0.92, 5.3, 0.75, [small], size=14, color=TEXT)
    y += 1.93
notes(s, "Слайд 2. Актуальность: теплицы на Севере, длительный период досветки и отопления, проектные решения принимаются заранее.")

# ================================================= slide 3: проблема
s = new_slide("Проблема: на какую зиму проектировать?")
card(s, 0.45, 1.15, 6.05, 2.45, fill=TINT)
tb(s, 0.7, 1.28, 5.6, 0.5, ["Типичный метеорологический год (TMY)"], size=18, bold=True, color=NAVY)
tb(s, 0.7, 1.82, 5.6, 1.7, [
    "Собирается из «наиболее типичных» месяцев многолетнего ряда; применяется в энергетических расчётах зданий и теплиц.",
    [("Описывает среднюю зиму → установка, рассчитанная по нему, может оказаться недостаточной примерно в каждую вторую зиму.", {"bold": True})],
], size=14, space_after=6)
arrow(s, 6.62, 2.05, 0.55, 0.6, color=NAVY)
card(s, 7.3, 1.15, 5.6, 2.45, fill=NAVY)
tb(s, 7.55, 1.28, 5.2, 0.5, ["Расчётная зима «1 раз в 10 лет» (p90)"], size=18, bold=True, color=WHITE)
tb(s, 7.55, 1.82, 5.2, 1.7, [
    "Значение, которое превышается лишь в одну зиму из десяти (90-й процентиль).",
    "Определяется по многолетнему ряду (44 зимы) отдельно для каждой расчётной величины: дефицита света, тепла, пиковых нагрузок.",
], size=14, color=WHITE, space_after=6)

tb(s, 0.45, 3.85, 12.4, 0.45, ["Нерешённые вопросы, на которые отвечает проект"], size=18, bold=True, color=NAVY)
qs = [
    "Типичный год никогда не сравнивали с расчётной зимой теплицы с досветкой — ни по дефициту света, ни по теплу с учётом тепловыделения ламп.",
    "За период наблюдений климат потеплел; неизвестно, одинаково ли это сказывается на световых и тепловых величинах.",
    "Данные реанализа не проверялись по самим расчётным величинам — по тёмным дням, которые формируют дефицит света.",
]
x = 0.45
for i, q in enumerate(qs, 1):
    card(s, x, 4.4, 4.0, 2.45, fill=RGBColor(0xF4, 0xF7, 0xFB))
    circle_num(s, x + 0.2, 4.55, 0.5, i)
    tb(s, x + 0.2, 5.15, 3.6, 1.65, [q], size=14)
    x += 4.2
notes(s, "Слайд 3. Постановка проблемы: TMY против расчётной зимы, три открытых вопроса.")

# ================================================= slide 4: цель и задачи
s = new_slide("Цель работы и решаемые задачи")
card(s, 0.45, 1.15, 4.3, 5.7, fill=NAVY)
tb(s, 0.7, 1.35, 3.85, 0.5, ["Цель работы"], size=20, bold=True, color=WHITE)
tb(s, 0.7, 1.95, 3.85, 4.8, [
    "Разработка методики и её программной реализации для определения расчётных параметров систем "
    "досветки и отопления высокоширотных теплиц (в том числе с фотоэлектрическими модулями на кровле) "
    "по многолетним климатическим данным с учётом оптических свойств кровли и изменения климата.",
], size=16, color=WHITE)
tasks = [
    "Разработать оптическую модель пропускания кровли (угловая зависимость остекления, диффузное излучение неба, отражение от земли и снега) для расчёта суточного интеграла ФАР на уровне растений.",
    "Разработать часовую модель теплового баланса теплицы с учётом тепловыделения светильников.",
    "Выполнить расчёт для шести городов 60–70° с. ш. за 44 зимы (1980/81–2023/24) по данным реанализа ERA5.",
    "Проверить входные данные по наземным наблюдениям сети BSRN (Леруик, Тыравере).",
    "Сравнить типичный год и опорный год EN ISO 15927-4 с расчётной зимой «1 раз в 10 лет».",
    "Оценить тренды и привести расчётные значения к современному климату.",
    "Определить мощность досветки, энергопотребление и нагрузки; составить бюджет погрешностей и рекомендации.",
]
y = 1.15
for i, t in enumerate(tasks, 1):
    circle_num(s, 5.0, y + 0.19, 0.42, i, size=13)
    tb(s, 5.55, y, 7.35, 0.8, [t], size=13.5, anchor=MSO_ANCHOR.MIDDLE, space_after=0)
    y += 0.81
notes(s, "Слайд 4. Цель и задачи проекта.")

# ================================================= slide 5: структура проекта / научный задел
s = new_slide("Научный задел и структура проекта")
mods = [
    ("ФЭ-модули на Севере", "Деградация и эффективность ФЭ-модулей в субарктическом климате: циклы замерзания–оттаивания, сравнение Санкт-Петербурга и Гонконга", "НТВ СПбГПУ, 2023–2024\nУч. зап. физфака МГУ, 2025"),
    ("Солнечный ресурс", "Сезонная изменчивость солнечных ресурсов и потенциала ФЭ-генерации в холодном климате (в т. ч. Харбин)", "Вестник ТОГУ (ВАК)"),
    ("Геометрия ФЭ-кровли", "Компромисс «свет для растений — выработка ФЭ-модулей» при разной геометрии кровли, 45–70° с. ш.", "IEEE Access, 2026\nЖТФ (ВАК)"),
    ("Оптика теплицы", "Самозатенение каркаса теплицы Venlo и угловая зависимость пропускания остекления", "PSRE 2025"),
    ("Расчётные зимы", "Типичный год против многолетнего ряда: досветка и отопление теплиц 60–70° с. ш.", "Основной результат\nпроекта"),
]
x = 0.45; w = 2.3; gap = 0.22
for i, (head, body, pub) in enumerate(mods):
    main = (i == len(mods) - 1)
    fill = NAVY if main else TINT
    fc = WHITE if main else TEXT
    card(s, x, 1.2, w, 3.85, fill=fill)
    circle_num(s, x + 0.15, 1.35, 0.46, i + 1, fill=HEAT if main else NAVY, size=14)
    tb(s, x + 0.12, 1.9, w - 0.24, 0.75, [head], size=16, bold=True, color=WHITE if main else NAVY)
    tb(s, x + 0.12, 2.6, w - 0.24, 2.4, [body], size=13, color=fc)
    card(s, x, 5.18, w, 0.85, fill=HEAT_TINT if main else RGBColor(0xF4, 0xF7, 0xFB))
    tb(s, x + 0.08, 5.2, w - 0.16, 0.8, pub.split("\n"), size=12, bold=True, color=HEAT if main else NAVY,
       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0)
    if i < len(mods) - 1:
        arrow(s, x + w + 0.02, 2.95, gap - 0.04, 0.32)
    x += w + gap
tb(s, 0.45, 6.2, 12.45, 0.7, [
    [("Итог: ", {"bold": True, "color": NAVY}),
     ("единая оптико-климатическая модель высокоширотной теплицы — от прохождения света через кровлю "
      "до мощности светильников, теплового источника и электрического присоединения.", {})]],
    size=15, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Слайд 5. Научный задел: пять блоков проекта и соответствующие публикации.")

# ================================================= slide 6: методика
s = new_slide("Разработанная методика расчёта")
# inputs
inp = [
    ("ERA5, 1980–2024, ежечасно", "суммарная и прямая радиация, температура, точка росы, ветер"),
    ("BSRN (наблюдения)", "наземные станции Леруик и Тыравере: проверка входных данных"),
    ("ERA5-Land, MERRA-2", "снежный покров, разброс реанализов"),
]
y = 1.2
for h, b in inp:
    card(s, 0.45, y, 3.1, 1.32, fill=RGBColor(0xF4, 0xF7, 0xFB), line=RGBColor(0xC9, 0xD3, 0xE0))
    tb(s, 0.6, y + 0.08, 2.85, 0.4, [h], size=13, bold=True, color=NAVY)
    tb(s, 0.6, y + 0.5, 2.85, 0.8, [b], size=12)
    y += 1.48
arrow(s, 3.62, 2.75, 0.4, 0.45)
# optical model
card(s, 4.1, 1.2, 4.0, 2.75, fill=LIGHT_TINT)
tb(s, 4.3, 1.3, 3.65, 0.45, ["Оптическая модель кровли"], size=15, bold=True, color=LIGHT)
tb(s, 4.3, 1.75, 3.65, 2.2, [
    "положение Солнца (NREL, pvlib), шаг 1 мин;",
    "модель яркости неба Переса, 145 участков Трегенза;",
    "пропускание стекла 4 мм по Френелю–Бугеру, двускатная кровля 25°;",
    "отражение от земли (альбедо 0,2) и вариант с ФЭ-полосами на 20 % кровли.",
], size=12, space_after=3)
arrow(s, 5.85, 4.0, 0.45, 0.32, shape=MSO_SHAPE.DOWN_ARROW)
card(s, 4.1, 4.4, 4.0, 1.0, fill=LIGHT)
tb(s, 4.25, 4.42, 3.7, 0.96, [[("DLI", {"bold": True}), ("ct", {"bold": True, "baseline": "-25000"}),
                              (" — суточный интеграл ФАР на уровне растений, моль·м⁻²·сут⁻¹", {})]],
   size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
# outputs
arrow(s, 8.18, 2.35, 0.4, 0.4)
arrow(s, 8.18, 4.7, 0.4, 0.4)
card(s, 8.65, 1.2, 4.25, 2.1, fill=LIGHT_TINT)
tb(s, 8.85, 1.28, 3.9, 0.4, ["Световые расчётные величины"], size=14, bold=True, color=LIGHT)
tb(s, 8.85, 1.7, 3.9, 1.6, [
    [("d", {}), sub("t"), (" = max(0, T − DLI", {}), sub("ct,t"), (")", {})],
    [("дефицит D", {}), sub("tot"), (", D", {}), sub("30"), (", длительность R", {}), sub("max"), (";", {})],
    [("энергия E", {}), sub("tot"), (" = D", {}), sub("tot"), (" / (3,6·PPE), мощность P", {}), sub("99"), ("", {})],
], size=12.5, space_after=3)
card(s, 8.65, 3.45, 4.25, 2.0, fill=HEAT_TINT)
tb(s, 8.85, 3.53, 3.9, 0.4, ["Часовой тепловой баланс"], size=14, bold=True, color=HEAT)
tb(s, 8.85, 3.95, 3.9, 1.5, [
    "учёт солнечного тепла и тепла ламп (04:00–20:00);",
    [("Q", {}), sub("heat"), (", Q", {}), sub("30"), (", P", {}), sub("heat,99.6"), ("; пик ламп и тепловых насосов (COP 3)", {})],
], size=12.5, space_after=3)
card(s, 4.1, 5.6, 8.8, 1.3, fill=NAVY)
tb(s, 4.3, 5.65, 8.45, 1.2, [
    [("Статистика 44 зим: ", {"bold": True}),
     ("распределение GEV (L-моменты), уровни повторяемости, p90; TMY (Sandia, TMY3) и EN ISO 15927-4; "
      "тренды (Сен, Манн–Кендалл), тест Петтитта; бюджет погрешностей из 7 возмущений.", {})]],
   size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
tb(s, 0.45, 5.65, 3.1, 1.25, ["Реализация: Python (pvlib, NumPy, SciPy, pandas)"], size=12, color=MUTED,
   anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Слайд 6. Структура методики: входные данные, оптическая модель кровли, световые величины, тепловой баланс, статистика.")

# ================================================= slide 7: валидация
s = new_slide("Проверка входных данных (BSRN)")
cd = CategoryChartData()
cd.categories = ["Сезонная сумма света", "Сезонный дефицит Dtot", "99-й процентиль дефицита", "Мощность досветки P99"]
cd.add_series("Леруик (6 зим, 1 869 сут.)", (16.7, -12.1, -0.7, -0.67))
cd.add_series("Тыравере (15 зим, 5 396 сут.)", (0.8, -4.1, -0.2, -0.36))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.35), Inches(1.15), Inches(7.4), Inches(5.75), cd)
ch = gf.chart
style_chart(ch, 12)
ch.has_title = True
ch.chart_title.text_frame.text = "Отклонение расчёта по ERA5 от наблюдений, %"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = True
ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False
plot = ch.plots[0]; plot.gap_width = 60; plot.overlap = -10
plot.has_data_labels = True
dl = plot.data_labels; dl.number_format = '0.0;−0.0'; dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.size = Pt(12); dl.font.bold = True
for ser, col in zip(plot.series, (LIGHT, GREY)):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col
va = ch.value_axis; va.has_major_gridlines = True
va.major_gridlines.format.line.color.rgb = RGBColor(0xE2, 0xE6, 0xEC)
va.format.line.fill.background(); va.tick_labels.font.size = Pt(11)
va.maximum_scale = 20; va.minimum_scale = -15
ca = ch.category_axis; ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
ca.tick_labels.font.size = Pt(11); ca.format.line.color.rgb = GREY

card(s, 8.0, 1.15, 4.9, 1.75, fill=LIGHT_TINT)
tb(s, 8.2, 1.22, 4.5, 1.65, [
    [("Тёмные дни ", {"bold": True, "color": LIGHT}), ("(DLI < 17): ERA5 завышает свет на ", {}),
     ("34,1 %", {"bold": True}), (" в Леруике и ", {}), ("11,8 %", {"bold": True}), (" в Тыравере.", {})],
    [("Яркие дни: ", {"bold": True, "color": LIGHT}), ("−1,5 % (Тыравере), +7,1 % (Леруик).", {})],
], size=14, space_after=6)
card(s, 8.0, 3.05, 4.9, 1.75, fill=RGBColor(0xF4, 0xF7, 0xFB))
tb(s, 8.2, 3.12, 4.5, 1.65, [
    "В Тыравере сезонная сумма света почти точна (+0,8 %), но дефицит занижен на 4,1 %: ошибки "
    "компенсируются в сумме, но не в дефиците.",
], size=14)
card(s, 8.0, 4.95, 4.9, 1.95, fill=NAVY)
tb(s, 8.2, 5.02, 4.5, 1.85, [
    [("Вывод: ", {"bold": True}),
     ("малая погрешность сезонной суммы радиации не гарантирует точности расчётных величин. "
      "Величины, зависящие от самых тёмных дней (P", {}), sub("99"), (", d", {}), sub("p99"),
     ("), почти не чувствительны к входным данным (< 1 %).", {})]],
   size=14, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Слайд 7. Валидация ERA5 по наземным наблюдениям BSRN.")

# ================================================= slide 8: TMY vs p90
s = new_slide("Типичный год против расчётной зимы")
RED_T = RGBColor(0xFB, 0xE3, 0xD3)
OK_T = RGBColor(0xE1, 0xF2, 0xE6)
rows = [
    ["Расчётная величина", "Весь ряд\n1980/81–2023/24", "Современный\nклимат (p90now)"],
    [[("Сезонный дефицит света D", {}), ("tot", {"baseline": "-25000"})], "−5,7 … −3,9", "−6,7 … −3,4"],
    [[("Сезонная потребность в тепле Q", {}), ("heat", {"baseline": "-25000"})], "−16,1 … −5,0", "−7,3 … +1,8"],
    [[("Макс. 30-суточная потребность в тепле Q", {}), ("30", {"baseline": "-25000"})], "−22,1 … −8,3", "−18,2 … −6,6"],
    [[("Расчётная тепловая нагрузка P", {}), ("heat,99.6", {"baseline": "-25000"})], "−13,0 … +0,3", "−10,8 … +10,3"],
    ["Градусо-сутки отопления HDD18", "−13,6 … −4,6", "−4,6 … +2,1"],
    [[("Опорный год EN ISO 15927-4: D", {}), ("tot", {"baseline": "-25000"})], "−6,5 … −2,1", "—"],
]
fills = {}
for i in range(1, len(rows)):
    fills[(i, 1)] = RED_T
fills[(1, 2)] = RED_T; fills[(3, 2)] = RED_T
fills[(2, 2)] = OK_T; fills[(5, 2)] = OK_T
fills[(4, 2)] = RGBColor(0xF4, 0xF7, 0xFB); fills[(6, 2)] = RGBColor(0xF4, 0xF7, 0xFB)
table(s, 0.45, 1.6, 7.9, rows, [4.3, 1.8, 1.8], row_h=0.62, size=13, fills=fills)
tb(s, 0.45, 1.1, 7.9, 0.45, ["Отклонение TMY от значения p90 в шести городах, %"], size=15, bold=True, color=NAVY)
tb(s, 0.45, 6.05, 7.9, 0.85, [
    [("■ ", {"color": HEAT}), ("TMY занижает расчётное значение (в большинстве городов)   ", {}),
     ("■ ", {"color": RGBColor(0x4C, 0xA8, 0x6B)}), ("близко к p90", {})],
    "p90now — 90-й процентиль, приведённый к уровню 2023/24 по линейному тренду",
], size=11.5, color=MUTED, space_after=2)
card(s, 8.65, 1.15, 4.25, 1.8, fill=LIGHT_TINT)
tb(s, 8.85, 1.22, 3.9, 0.75, ["18–43-й"], size=32, bold=True, color=LIGHT)
tb(s, 8.85, 1.95, 3.9, 1.0, ["процентиль — место TMY по дефициту света среди 44 зим: это типичная, а не расчётная зима"], size=12.5)
card(s, 8.65, 3.1, 4.25, 1.6, fill=RGBColor(0xF4, 0xF7, 0xFB))
tb(s, 8.85, 3.17, 3.9, 0.7, ["54 из 54"], size=32, bold=True, color=NAVY)
tb(s, 8.85, 3.85, 3.9, 0.85, ["сочетаний порогов, возмущений и городов — вывод по дефициту света не меняется"], size=12.5)
card(s, 8.65, 4.85, 4.25, 2.05, fill=NAVY)
tb(s, 8.85, 4.92, 3.9, 1.95, [
    "Для тепла недобор TMY по сезону исчезает после учёта тренда, но максимальная 30-суточная "
    "потребность (самый холодный месяц) занижается на 6,6–18,2 %.",
], size=13.5, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Слайд 8. Сравнение TMY и EN ISO 15927-4 с 90-м процентилем 44 зим.")

# ================================================= slide 9: тренды
s = new_slide("Тренды: свет стабилен, тепло снижается")
cd = CategoryChartData()
cd.categories = ["Леруик", "Рейкьявик", "Тромсё", "Норильск", "С.-Петербург", "Якутск"]
cd.add_series("Дефицит света Dtot", (-0.8, 0.9, 0.3, 0.3, 0.0, 0.4))
cd.add_series("Потребность в тепле Qheat", (-3.5, -4.8, -5.7, -2.1, -4.4, -1.7))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.35), Inches(1.15), Inches(7.6), Inches(5.75), cd)
ch = gf.chart; style_chart(ch, 12)
ch.has_title = True
ch.chart_title.text_frame.text = "Тренд за 44 зимы, % от медианы за 10 лет (наклон Сена)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = True
ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False
plot = ch.plots[0]; plot.gap_width = 70; plot.overlap = -10
plot.has_data_labels = True
dl = plot.data_labels; dl.number_format = '0.0;−0.0;0.0'; dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.size = Pt(11); dl.font.bold = True
for ser, col in zip(plot.series, (LIGHT, HEAT)):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col
va = ch.value_axis; va.has_major_gridlines = True
va.major_gridlines.format.line.color.rgb = RGBColor(0xE2, 0xE6, 0xEC)
va.format.line.fill.background(); va.maximum_scale = 2; va.minimum_scale = -7; va.major_unit = 1
va.tick_labels.font.size = Pt(11)
ca = ch.category_axis; ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW; ca.tick_labels.font.size = Pt(11)
ca.format.line.color.rgb = GREY

card(s, 8.2, 1.15, 4.7, 1.8, fill=LIGHT_TINT)
tb(s, 8.4, 1.22, 4.3, 1.7, [
    [("Свет: ", {"bold": True, "color": LIGHT}), ("значимого тренда нет ни в одном городе (p > 0,05).", {})],
    [("→ p90 дефицита света берётся прямо из многолетнего ряда.", {"bold": True})],
], size=14, space_after=6)
card(s, 8.2, 3.1, 4.7, 1.95, fill=HEAT_TINT)
tb(s, 8.4, 3.17, 4.3, 1.85, [
    [("Тепло: ", {"bold": True, "color": HEAT}), ("−1,7 … −5,7 % за 10 лет во всех городах (p ≤ 0,003).", {})],
    [("→ p90 нужно приводить к современному климату: он на 5,0–15,9 % ниже p90 всего ряда.", {"bold": True})],
], size=14, space_after=6)
card(s, 8.2, 5.2, 4.7, 1.7, fill=RGBColor(0xF4, 0xF7, 0xFB))
tb(s, 8.4, 5.27, 4.3, 1.6, [
    "Все наблюдённые расчётные зимы — 1980-х годов; для отопления сегодня они дают запас до 21,6 %. "
    "Точки перелома (тест Петтитта) — 1995–2004 гг.",
], size=13)
notes(s, "Слайд 9. Тренды световых и тепловых величин.")

# ================================================= slide 10: расчётные зимы
s = new_slide("Расчётные зимы и нагрузки")
hdr = ["Город", "Расчётная\nзима", [("R", {}), ("max", {"baseline": "-25000"}), (",\nсут", {})],
       [("E", {}), ("tot", {"baseline": "-25000"}), (",\nкВт·ч/м²", {})],
       [("P", {}), ("99", {"baseline": "-25000"}), (",\nВт/м²", {})],
       [("Q", {}), ("heat", {"baseline": "-25000"}), (",\nкВт·ч/м²", {})],
       "Тепло ламп\nв отоплении, %",
       [("P", {}), ("heat,99.6", {"baseline": "-25000"}), (",\nВт/м²", {})],
       "Эл. пик\n(COP 3), Вт/м²"]
data = [
    ["Леруик", "2018/19", "183", "261,5", "115,6", "435", "35,5", "154", "136"],
    ["Рейкьявик", "2002/03", "200", "323,5", "117,4", "650", "32,1", "263", "175"],
    ["Тромсё", "1992/93", "216", "342,8", "118,1", "1096", "23,4", "401", "225"],
    ["Норильск", "2013/14", "211", "308,5", "118,1", "1963", "13,5", "640", "304"],
    ["С.-Петербург", "1999/00", "177", "254,9", "114,6", "802", "23,6", "337", "220"],
    ["Якутск", "1983/84", "176", "238,3", "113,0", "2017", "10,3", "657", "304"],
]
fills = {}
for i in range(1, 7):
    for j in (3, 4):
        fills[(i, j)] = LIGHT_TINT
    for j in (5, 6, 7):
        fills[(i, j)] = HEAT_TINT
table(s, 0.45, 1.15, 12.45, [hdr] + data, [1.75, 1.25, 1.05, 1.45, 1.15, 1.45, 1.6, 1.4, 1.35],
      row_h=0.43, size=13, fills=fills, bold_first_col=True)
tb(s, 0.45, 4.25, 12.45, 0.35, ["Целевой DLI 17 моль·м⁻²·сут⁻¹, эффективность светильников 2,5 мкмоль/Дж, фотопериод 16 ч; "
                                "стеклянная теплица типа Venlo"], size=11, color=MUTED)
st = [
    ("113–118 Вт/м²", "мощность ламп почти одинакова во всех городах: определяется целевым DLI, фотопериодом и эффективностью светильников", LIGHT, LIGHT_TINT),
    ("10–34 %", "потребности в отоплении покрывает тепло ламп; больше всего — в морском климате", HEAT, HEAT_TINT),
    ("76–90 %", "суммы раздельных пиков составляет совмещённый пик ламп и тепловых насосов: простое сложение завышает нагрузку на 11–32 %", NAVY, TINT),
]
x = 0.45
for big, small, col, tint in st:
    card(s, x, 4.7, 4.0, 2.2, fill=tint)
    tb(s, x + 0.2, 4.78, 3.6, 0.7, [big], size=28, bold=True, color=col)
    tb(s, x + 0.2, 5.48, 3.6, 1.4, [small], size=12.5)
    x += 4.225
notes(s, "Слайд 10. Расчётные зимы по дефициту света и соответствующие нагрузки.")

# ================================================= slide 11: бюджет погрешностей
s = new_slide("Бюджет погрешностей энергии досветки")
cats = ["Порог DLI −2 моль·м⁻²·сут⁻¹", "Порог DLI +2 моль·м⁻²·сут⁻¹", "Эффективность ламп 3,0 мкмоль/Дж",
        "ФЭ-полосы на 20 % кровли", "Ошибка реанализа, Леруик (измерено)", "Ошибка реанализа, Тыравере → С.-Петербург",
        "Постоянный коэффициент пропускания τ₀", "Альбедо снега"]
vals = [-17.8, 18.7, -16.7, 10.0, -12.1, -4.1, 0.8, -0.3]
cols = [NAVY, NAVY, NAVY, LIGHT, HEAT, HEAT, GREY, GREY]
cd = CategoryChartData(); cd.categories = list(reversed(cats)); cd.add_series("Изменение Etot, %", list(reversed(vals)))
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.35), Inches(1.15), Inches(8.1), Inches(5.75), cd)
ch = gf.chart; style_chart(ch, 12)
ch.has_title = True
ch.chart_title.text_frame.text = "Изменение энергии досветки Etot, % (медиана)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = True
ch.has_legend = False
plot = ch.plots[0]; plot.gap_width = 45
plot.has_data_labels = True
dl = plot.data_labels; dl.number_format = '+0.0;−0.0'; dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.size = Pt(12); dl.font.bold = True
ser = plot.series[0]
for idx, col in enumerate(reversed(cols)):
    pt = ser.points[idx]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = col
va = ch.value_axis; va.minimum_scale = -25; va.maximum_scale = 25; va.major_unit = 5
va.has_major_gridlines = True; va.major_gridlines.format.line.color.rgb = RGBColor(0xE2, 0xE6, 0xEC)
va.format.line.fill.background(); va.tick_labels.font.size = Pt(11)
ca = ch.category_axis; ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW; ca.tick_labels.font.size = Pt(11.5)
ca.format.line.color.rgb = GREY

legend = [
    (NAVY, "Решения проектировщика", "16–20 %"),
    (LIGHT, "ФЭ-модули на кровле", "7–14 %"),
    (HEAT, "Ошибка входных данных", "до 12 %"),
    (GREY, "Упрощения оптической модели", "≤ 1,3 %"),
]
y = 1.2
for col, lab, val in legend:
    card(s, 8.75, y, 4.15, 1.0, fill=RGBColor(0xF4, 0xF7, 0xFB))
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.95), Inches(y + 0.36), Inches(0.28), Inches(0.28))
    dot.fill.solid(); dot.fill.fore_color.rgb = col; dot.line.fill.background(); dot.shadow.inherit = False
    tb(s, 9.35, y + 0.08, 2.2, 0.85, [lab], size=13, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 11.5, y + 0.08, 1.3, 0.85, [val], size=17, bold=True, color=col if col != GREY else MUTED,
       align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    y += 1.12
card(s, 8.75, 5.75, 4.15, 1.15, fill=NAVY)
tb(s, 8.9, 5.8, 3.85, 1.05, ["Точность проекта определяют выбор целевого DLI и светильников, а не детали оптической модели"],
   size=13, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Слайд 11. Бюджет погрешностей энергии досветки.")

# ================================================= slide 12: практическая значимость
s = new_slide("Практическая значимость и перспективы")
tb(s, 0.45, 1.1, 7.3, 0.45, ["Рекомендации для проектирования"], size=18, bold=True, color=NAVY)
recs = [
    [("Мощность светильников", {"bold": True}), (" — по целевому DLI, фотопериоду и эффективности ламп (≈113–118 Вт/м² при 17 моль·м⁻²·сут⁻¹ и 2,5 мкмоль/Дж).", {})],
    [("Энергия досветки", {"bold": True}), (" — по многолетнему ряду или расчётной зиме, а не по типичному году.", {})],
    [("Тепловые величины", {"bold": True}), (" — по ряду, приведённому к современному климату; особое внимание — максимальной 30-суточной потребности в тепле.", {})],
    [("Электрическое присоединение", {"bold": True}), (" — по совместному часовому моделированию ламп и тепловых насосов, а не по сумме пиков.", {})],
]
y = 1.65
for i, r in enumerate(recs, 1):
    card(s, 0.45, y, 7.3, 1.2, fill=RGBColor(0xF4, 0xF7, 0xFB))
    circle_num(s, 0.62, y + 0.35, 0.5, i)
    tb(s, 1.3, y + 0.05, 6.35, 1.1, [r], size=14, anchor=MSO_ANCHOR.MIDDLE)
    y += 1.32
card(s, 8.05, 1.15, 4.85, 2.85, fill=NAVY)
tb(s, 8.25, 1.25, 4.5, 0.45, ["Области применения"], size=16, bold=True, color=WHITE)
tb(s, 8.25, 1.72, 4.5, 2.25, [
    "тепличные комплексы Северо-Запада, Арктической зоны РФ, Якутии, Норильского промышленного района;",
    "северные регионы Китая (Харбин) и Северной Европы;",
    "проектные организации: программная реализация методики на Python.",
], size=13, color=WHITE, space_after=5, bullets=True)
card(s, 8.05, 4.15, 4.85, 2.75, fill=TINT)
tb(s, 8.25, 4.25, 4.5, 0.45, ["Перспективы"], size=16, bold=True, color=NAVY)
tb(s, 8.25, 4.72, 4.5, 2.15, [
    "климатические данные будущего для тепловых расчётов;",
    "больше станций BSRN — оценка погрешности уровней повторяемости;",
    "угловая модель остекления для почасового управления досветкой;",
    "открытый код (Zenodo) после публикации.",
], size=13, space_after=4, bullets=True)
notes(s, "Слайд 12. Рекомендации, области применения, перспективы.")

# ================================================= slide 13: публикации
s = new_slide("Публикации и апробация результатов")
G = RGBColor(0x2E, 0x8B, 0x57)
def ok(t):
    return [(t, {"color": G, "bold": True})]
pubs = [
    ["№", "Публикация", "Статус"],
    ["1", "A New Method for Evaluating PV Roof Geometry and Crop-Light Tradeoffs in Agrivoltaic Greenhouses across 45–70°N Climates // IEEE Access. 2026. DOI 10.1109/ACCESS.2026.3740614", ok("Опубликована")],
    ["2", "Анализ эффективности и исследование особенностей конструкций агровольтаических систем в высоких широтах // Журнал технической физики (ВАК)", ok("Принята к печати")],
    ["3", "Using silicon quantum dots to prevent solar cells from overheating and improve photoelectric conversion efficiency // Физика и техника полупроводников (ВАК)", ok("Принята к печати")],
    ["4", "Исследование сезонного изменения солнечных ресурсов и потенциала фотоэлектрической генерации в холодном климате (единоличный автор) // Вестник ТОГУ (ВАК)", ok("Принята, 2027")],
    ["5", "Механизмы деградации, вызванные циклами замерзания–оттаивания, в фотоэлектрических модулях в условиях холодного климата // Учёные записки физического факультета МГУ (ВАК). 2025", ok("Опубликована")],
    ["6", "Photovoltaic Potential in the Subarctic Region // Научно-технические ведомости СПбГПУ. Физ.-мат. науки. 2024. Т. 17, № 3.2", ok("Опубликована")],
    ["7", "Features of Photovoltaic Cell Degradation of Solar Power Plants in Hong Kong and Saint Petersburg // Научно-технические ведомости СПбГПУ. Физ.-мат. науки. 2023. Т. 16, № 3.2. С. 44–49", ok("Опубликована")],
    ["8", "Impact of Self-Shading on Solar Irradiance for Venlo Greenhouse Agrivoltaics in High Latitudes // PSRE 2025", ok("Доклад, 2025")],
    ["9", "Performance Challenges of Photovoltaic Modules in Subarctic Environments // ElCon, СПбГЭТУ «ЛЭТИ»", ok("Доклад")],
    ["10", "Search for New Solutions to Improve the Efficiency of Solar Cells Using in Saint Petersburg and Hong Kong // EExPolytech 2023 (IEEE). С. 348–351", ok("Опубликована")],
]
table(s, 0.45, 1.1, 12.45, pubs, [0.45, 10.2, 1.8], row_h=0.52, size=10.5,
      aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT])
notes(s, "Слайд 13. Публикации по теме проекта.")

# ================================================= slide 14: заключение
s = new_slide("Заключение")
concl = [
    "Разработана методика определения расчётных параметров досветки и отопления высокоширотной теплицы: оптическая модель кровли, часовой тепловой баланс с учётом тепла ламп и статистика 44 зим; методика реализована программно.",
    "Сезонный дефицит света определяется местоположением (медианы городов различаются до 50 %, зимы в одном городе — на 12–15 %) и не имеет значимого тренда; потребность в тепле снижается на 1,7–5,7 % за 10 лет.",
    "Типичный метеорологический год описывает типичную, а не расчётную зиму: дефицит света занижен на 3,9–5,7 % (3,4–6,7 % в современном климате), максимальная 30-суточная потребность в тепле — на 6,6–18,2 %.",
    "Мощность досветки (113–118 Вт/м²) задаётся целевым DLI, фотопериодом и эффективностью ламп; тепло ламп покрывает 10–34 % отопления; совмещённый электрический пик — 76–90 % суммы раздельных пиков.",
    "ERA5 занижает дефицит света на 12,1 % (Леруик) и 4,1 % (Тыравере); в Тыравере — даже при почти точной сезонной сумме радиации (+0,8 %).",
    "Результаты апробированы в 10 работах: 7 статей в журналах (IEEE Access, ЖТФ, ФТП, Вестник ТОГУ, Учёные записки физфака МГУ, НТВ СПбГПУ) и 3 доклада на международных конференциях (EExPolytech, ElCon, PSRE).",
]
y = 1.12
for i, c in enumerate(concl, 1):
    circle_num(s, 0.45, y + 0.24, 0.44, i, size=13)
    tb(s, 1.05, y, 11.85, 0.92, [c], size=14, anchor=MSO_ANCHOR.MIDDLE)
    y += 0.97
notes(s, "Слайд 14. Основные результаты.")

# ------------------------------------------------ last slide (template): move to end & fill
sldIdLst.remove(last_id)
sldIdLst.append(last_id)
last = prs.slides[len(prs.slides) - 1]
for sh in last.shapes:
    if not sh.has_text_frame:
        continue
    t = sh.text_frame.text
    if t.startswith("Мобильный"):
        set_runs(sh, PROJECT)
        sh.height = Inches(1.3)
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(15)
    elif t.startswith("Давыдова"):
        set_runs(sh, "Дэн Юаньбяо")
    elif t.startswith("Курс"):
        set_runs(sh, "Курс обучения: 3")
    elif t.startswith("Уровень"):
        set_runs(sh, "Уровень образования: аспирантура")
    elif t.strip().startswith("2025"):
        set_runs(sh, "2026 год")
notes(last, "Завершающий слайд.")

prs.save(OUT)
print("saved", OUT, len(prs.slides), "slides")
