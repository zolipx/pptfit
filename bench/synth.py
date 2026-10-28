"""Synthetic benchmark: common Chinese-deck layouts plus their English translation.

python bench/synth.py outdir  ->  outdir/synth.pptx, outdir/synth_tr.json, outdir/synth_gl.json
Every pattern is tight for Chinese, so the English needs fitting. Nothing here is
used by the toolkit itself; it only exists to test that the toolkit generalises.
"""
import json, os, sys

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.util import Pt

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from pptfit import units as U

# phrase bank: Chinese -> (full English, short English, concepts)
BANK = {
    "服务流程": ("Service flow", "Service flow", [("服务", "service"), ("流程", "flow")]),
    "门店概览": ("Store overview", "Store overview", [("门店", "store"), ("概览", "overview")]),
    "门店亮点": ("Store highlights", "Store highlights", [("门店", "store"), ("亮点", "highlight")]),
    "顾客点单（自助）": ("Customer ordering (self-service)", "Customer order (self-service)", [("顾客", "customer"), ("点单", "order"), ("自助", "self-service")]),
    "收银确认": ("Payment check", "Payment check", [("收银", "payment"), ("确认", "check")]),
    "研磨咖啡": ("Grinding coffee", "Grinding coffee", [("研磨", "grind"), ("咖啡", "coffee")]),
    "萃取浓缩": ("Espresso extraction", "Extract espresso", [("萃取", "extract"), ("浓缩", "espresso")]),
    "打发奶泡": ("Milk frothing", "Milk froth", [("打发", "froth"), ("奶泡", "milk")]),
    "拉花装杯": ("Latte art and pouring into cups", "Latte art & cups", [("拉花", "latte art"), ("装杯", "cup")]),
    "门店取餐（自助）": ("Pick-up in store (self-service)", "Store pick-up (self-service)", [("门店", "store"), ("取餐", "pick-up"), ("自助", "self-service")]),
    "外卖打包": ("Packing for delivery", "Delivery packing", [("外卖", "delivery"), ("打包", "pack")]),
    "提升销量": ("Higher sales", "Higher sales", [("提升", "higher"), ("销量", "sales")]),
    "日销量翻倍，顾客排队时间缩短一半。": ("Daily sales double and customers wait half as long in line.", "Daily sales double; wait time halves.", [("销量", "sales"), ("排队", "wait")]),
    "节省人力": ("Leaner staffing", "Leaner staffing", [("节省", "leaner"), ("人力", "staff")]),
    "自助点单减少收银，节省人力和培训开支。": ("Self-service ordering cuts cashier work and saves staff and training costs.", "Self-ordering cuts cashier work, staff and training costs.", [("自助", "self"), ("收银", "cashier"), ("培训", "training")]),
    "留住顾客": ("Loyal customers", "Loyal customers", [("留住", "loyal"), ("顾客", "customer")]),
    "会员积分带来回头客，复购率明显提高。": ("Member points bring customers back, and the repeat purchase rate rises clearly.", "Member points raise repeat purchases.", [("会员", "member"), ("积分", "points"), ("复购", "repeat")]),
    "咖啡机": ("Coffee machine", "Coffee machine", [("咖啡机", "coffee machine")]),
    "磨豆机": ("Grinder", "Grinder", [("磨豆机", "grinder")]),
    "吧台": ("Bar counter", "Bar", [("吧台", "bar")]),
    "区域介绍：": ("Area overview:", "Area overview:", [("区域", "area"), ("介绍", "overview")]),
    "1）顾客在入口扫码点单，订单自动发送到吧台屏幕；": ("1) Customers scan a code at the entrance to order; the order goes straight to the screen at the bar;", "1) Customers scan to order at the entrance; orders go straight to the bar screen;", [("顾客", "customer"), ("吧台", "bar")]),
    "2）咖啡师按屏幕顺序制作并呼叫取餐。": ("2) Baristas make the drinks in screen order and call customers to pick them up.", "2) Baristas make drinks in screen order and call for pick-up.", [("咖啡师", "barista"), ("取餐", "pick")]),
    "项目": ("Item", "Item", [("项目", "item")]),
    "旧门店": ("Old store", "Old store", [("旧", "old"), ("门店", "store")]),
    "新门店": ("New store", "New store", [("新", "new"), ("门店", "store")]),
    "店员人数": ("Number of staff", "Staff", [("店员", "staff")]),
    "每日杯量": ("Cups per day", "Cups/day", [("杯", "cup"), ("每日", "day")]),
    "店铺面积": ("Floor space", "Floor space", [("面积", "space")]),
    "8 人": ("8 people", "8 people", []),
    "5 人": ("5 people", "5 people", []),
    "400 杯": ("400 cups", "400 cups", []),
    "900 杯": ("900 cups", "900 cups", []),
    "60 平方米": ("60 square meters", "60 m²", []),
    "45 平方米": ("45 square meters", "45 m²", []),
    "人均杯量对比": ("Cups per staff member", "Cups per staff", [("人均", "per staff"), ("杯量", "cup")]),
    "门店数据": ("Store facts", "Store facts", [("门店", "store"), ("数据", "facts")]),
    "营业时间：7:00–22:00，座位：36 个，出杯时间：约 90 秒/杯，外卖范围 3 公里内，高峰时段每小时约 120 杯，月租金约 18000 元。": (
        "Opening hours: 7:00–22:00; seats: 36; serving time: about 90 s per cup; delivery within 3 km; about 120 cups per hour at peak; monthly rent about 18000 yuan.",
        "Hours 7:00–22:00; 36 seats; serving ~90 s/cup; delivery <3 km; ~120 cups/h at peak; rent ~18000 yuan/month.",
        [("营业", "hours"), ("座位", "seats"), ("出杯", "serving"), ("外卖", "delivery"), ("租金", "rent")]),
    "点单区域": ("Ordering area", "Order area", [("点单", "order"), ("区域", "area")]),
    "取餐区域": ("Pick-up area", "Pick-up area", [("取餐", "pick-up"), ("区域", "area")]),
    "消防通道": ("Fire lane", "Fire lane", [("消防", "fire"), ("通道", "lane")]),
    "紧急出口": ("Emergency exit", "Emergency exit", [("紧急", "emergency"), ("出口", "exit")]),
}

BLUE, GREEN, GREY = RGBColor(0x1F, 0x4E, 0x99), RGBColor(0xD9, 0xEA, 0xD3), RGBColor(0x40, 0x40, 0x40)


def text(slide, x, y, w, h, s, size=14, bold=False, align=None, anchor=None, shape=None, fill=None, line=None, font="微软雅黑"):
    sh = slide.shapes.add_shape(shape, Pt(x), Pt(y), Pt(w), Pt(h)) if shape else slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    if shape:
        if fill is None:
            sh.fill.background()
        else:
            sh.fill.solid(); sh.fill.fore_color.rgb = fill
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = line; sh.line.width = Pt(1.25)
    tf = sh.text_frame
    tf.word_wrap = True
    if anchor:
        from pptx.enum.text import MSO_ANCHOR
        tf.vertical_anchor = {"middle": MSO_ANCHOR.MIDDLE, "top": MSO_ANCHOR.TOP}[anchor]
    for i, para in enumerate(s.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if align:
            from pptx.enum.text import PP_ALIGN
            p.alignment = {"center": PP_ALIGN.CENTER, "left": PP_ALIGN.LEFT}[align]
        r = p.add_run(); r.text = para
        r.font.size, r.font.bold, r.font.name = Pt(size), bold, font
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if fill == BLUE else GREY
        rpr = r._r.get_or_add_rPr()
        ea = rpr.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}ea", {"typeface": font})
        rpr.append(ea)
    return sh


def title(slide, s):
    """Title + the decorative rule that starts right after it (tight for Chinese)."""
    t = text(slide, 40, 18, 18 * len(s) + 16, 30, s, size=18, bold=True)
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(40 + 18 * len(s) + 20), Pt(34), Pt(900), Pt(34))
    ln.line.color.rgb = BLUE; ln.line.width = Pt(1)
    return t


def footer_in_layout(prs, layout):
    """A dark footer bar living in the layout, not on the slides."""
    tmp = prs.slides.add_slide(layout)
    bar = tmp.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Pt(510), prs.slide_width, Pt(30))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0x12, 0x2B, 0x50); bar.line.fill.background()
    layout.shapes._spTree.append(bar._element)
    rid = prs.slides._sldIdLst[-1].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    prs.part.drop_rel(rid); prs.slides._sldIdLst.remove(prs.slides._sldIdLst[-1])


def picture(path):
    img = Image.new("RGB", (480, 400), (200, 210, 230))
    d = ImageDraw.Draw(img)
    for i in range(0, 480, 40):
        d.rectangle([i, 300, i + 30, 380], fill=(60, 80, 160))
    d.ellipse([180, 80, 300, 200], fill=(240, 200, 60))
    img.save(path)
    return path


def build(out):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(960), Pt(540)
    layout = prs.slide_layouts[6]
    footer_in_layout(prs, layout)
    pic = picture(os.path.join(out, "pic.png"))

    s = prs.slides.add_slide(layout)                      # 1 flowchart: look-alike boxes + arrows
    title(s, "服务流程")
    steps = ["顾客点单（自助）", "收银确认", "研磨咖啡", "萃取浓缩", "打发奶泡", "拉花装杯", "门店取餐（自助）", "外卖打包"]
    for i, st in enumerate(steps):
        r, c = divmod(i, 4)
        x, y = 60 + c * 220, 110 + r * 170
        text(s, x, y, 150, 52, st, size=18, bold=True, align="center", anchor="middle",
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=BLUE)
        if c < 3:
            a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Pt(x + 160), Pt(y + 16), Pt(50), Pt(20))
            a.fill.background(); a.line.color.rgb = BLUE

    s = prs.slides.add_slide(layout)                      # 2 cards: text on separate coloured panels
    title(s, "门店亮点")
    cards = [("提升销量", "日销量翻倍，顾客排队时间缩短一半。"), ("节省人力", "自助点单减少收银，节省人力和培训开支。"),
             ("留住顾客", "会员积分带来回头客，复购率明显提高。")]
    for i, (h, b) in enumerate(cards):
        x = 60 + i * 290
        panel = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(100), Pt(250), Pt(150))
        panel.fill.solid(); panel.fill.fore_color.rgb = GREEN; panel.line.fill.background()
        text(s, x + 10, 110, 230, 30, h, size=18, bold=True, align="center")
        text(s, x + 10, 150, 230, 90, b, size=16)

    s = prs.slides.add_slide(layout)                      # 3 picture + callouts + body column above the footer
    title(s, "门店概览")
    s.shapes.add_picture(pic, Pt(60), Pt(70), Pt(420), Pt(350))
    for i, lab in enumerate(["咖啡机", "磨豆机", "吧台"]):
        text(s, 30, 120 + i * 100, 70, 22, lab, size=11, align="center", shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=RGBColor(0xE0, 0, 0))
    text(s, 510, 80, 380, 420, "区域介绍：\n1）顾客在入口扫码点单，订单自动发送到吧台屏幕；\n2）咖啡师按屏幕顺序制作并呼叫取餐。", size=18)

    s = prs.slides.add_slide(layout)                      # 4 table with fixed columns
    title(s, "门店数据")
    rows = [["项目", "旧门店", "新门店"], ["店员人数", "8 人", "5 人"], ["每日杯量", "400 杯", "900 杯"], ["店铺面积", "60 平方米", "45 平方米"]]
    tb = s.shapes.add_table(4, 3, Pt(80), Pt(90), Pt(420), Pt(200)).table
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            cell = tb.cell(r, c); cell.text = v
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(16)
    text(s, 540, 90, 360, 160, "营业时间：7:00–22:00，座位：36 个，出杯时间：约 90 秒/杯，外卖范围 3 公里内，高峰时段每小时约 120 杯，月租金约 18000 元。",
         size=16, shape=MSO_SHAPE.RECTANGLE, fill=GREEN)

    s = prs.slides.add_slide(layout)                      # 5 chart with Chinese categories
    title(s, "人均杯量对比")
    cd = CategoryChartData(); cd.categories = ["旧门店", "新门店"]; cd.add_series("每日杯量", (400, 900))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Pt(80), Pt(80), Pt(360), Pt(300), cd)

    s = prs.slides.add_slide(layout)                      # 6 group of tags + autofit box
    title(s, "门店数据")
    a = text(s, 80, 100, 110, 30, "点单区域", size=14, align="center", shape=MSO_SHAPE.RECTANGLE, line=BLUE)
    b = text(s, 210, 100, 110, 30, "取餐区域", size=14, align="center", shape=MSO_SHAPE.RECTANGLE, line=BLUE)
    grp = s.shapes.add_group_shape([a, b])
    for i, lab in enumerate(["消防通道", "紧急出口"]):
        text(s, 80 + i * 130, 160, 110, 26, lab, size=13, align="center", shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=GREY)
    box = text(s, 420, 100, 300, 70, "营业时间：7:00–22:00，座位：36 个，出杯时间：约 90 秒/杯，外卖范围 3 公里内，高峰时段每小时约 120 杯，月租金约 18000 元。", size=14)
    bp = box.text_frame._txBody.find("{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr")
    bp.append(bp.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}normAutofit", {"fontScale": "85000"}))

    path = os.path.join(out, "synth.pptx")
    prs.save(path)
    return path


def translations(path):
    out, missing = [], []
    for u in U.make_units(path):
        src = u["src"].strip()
        if src in BANK:
            full, short, con = BANK[src]
            out.append({"id": u["id"], "src": u["src"], "full": full, "short": short,
                        "concepts": [{"zh": z, "en": e} for z, e in con]})
        else:
            missing.append(src)
    return out, missing


if __name__ == "__main__":
    outdir = sys.argv[1]
    os.makedirs(outdir, exist_ok=True)
    p = build(outdir)
    tr, missing = translations(p)
    json.dump(tr, open(os.path.join(outdir, "synth_tr.json"), "w"), ensure_ascii=False, indent=1)
    json.dump([], open(os.path.join(outdir, "synth_gl.json"), "w"))
    print("units", len(tr), "missing", missing)
