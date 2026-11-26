# -*- coding: utf-8 -*-
"""
3D 列印入門工作坊簡報產生器
用法：python build_slides.py [--font 字型名稱] [--out 輸出檔名]

動畫設計以「FDM 3D 列印」為核心：
  ● 噴嘴列印（Print）   ：封面／結尾／每頁右下角的小物件，噴嘴沿 Z 字路徑來回移動，
                           每走一趟就「擠出」一層線材（擦去效果方向與噴嘴同步），一層層往上堆。
  ● 擠出（Extrude）     ：標題、單行文字、箭頭、時間軸都像噴嘴擠線一樣由左向右畫出。
  ● 堆疊（Layer up）    ：卡片、圖片框由下往上長出來，像積層製造。
  ● 預熱（Preheat）     ：安全須知的溫度卡片會脈動放大，提醒「很燙」。
  ● 列印進度條（Progress）：倒數頁面的橘色線材條會在實際倒數時間內慢慢擠滿。
  ● 換頁轉場            ：所有頁面「由下往上擦去」，像列印平台一層層升起。
"""
import argparse
import copy
import os
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE

# ---------------------------------------------------------------- 基本設定
W, H = 13.333, 7.5
BG = RGBColor(0x1F, 0x29, 0x33)
CARD = RGBColor(0x2B, 0x38, 0x46)
CARD2 = RGBColor(0x35, 0x44, 0x54)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0xC3, 0xCC, 0xD6)
GRAY = RGBColor(0x8A, 0x97, 0xA5)
ORANGE = RGBColor(0xF2, 0x8C, 0x28)
ORANGE_D = RGBColor(0xD2, 0x70, 0x14)
ORANGE_BG = RGBColor(0x4A, 0x36, 0x24)
DARK = RGBColor(0x1F, 0x29, 0x33)
FONT = "微軟正黑體"

PLACEHOLDERS = []      # (頁碼, 說明) — 圖片預留框清單
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

SEC = {
    "open": "開場｜0–10 分",
    "basic": "基礎概念｜10–22 分",
    "ind": "業界應用｜22–34 分",
    "tour": "休息＋機台參觀｜34–44 分",
    "mw": "MakerWorld 實作｜44–92 分",
    "share": "小組發表｜92–108 分",
    "end": "總結｜108–120 分",
}


# ---------------------------------------------------------------- 文字工具
def _font(run, size, color, bold):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    f.name = FONT
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(f"{{{A}}}latin")
    for tag in ("ea", "cs"):
        old = rPr.find(f"{{{A}}}{tag}")
        if old is not None:
            rPr.remove(old)
    ea = etree.SubElement(rPr, f"{{{A}}}ea")
    ea.set("typeface", FONT)
    cs = etree.SubElement(rPr, f"{{{A}}}cs")
    cs.set("typeface", FONT)
    latin.addnext(ea)
    ea.addnext(cs)


def fill_text(tf, paras, size=22, color=WHITE, bold=False, align=PP_ALIGN.LEFT,
              anchor=MSO_ANCHOR.TOP, margin=0.05, space_after=0, line_spacing=None):
    """paras：字串，或段落清單；每段是字串或 [(文字, {size,color,bold}) ...]"""
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    m = Inches(margin)
    tf.margin_left = tf.margin_right = m
    tf.margin_top = tf.margin_bottom = Inches(min(margin, 0.05))
    if isinstance(paras, str):
        paras = [paras]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if space_after:
            p.space_after = Pt(space_after)
        if line_spacing:
            p.line_spacing = line_spacing
        runs = [(para, {})] if isinstance(para, str) else para
        for text, o in runs:
            r = p.add_run()
            r.text = text
            _font(r, o.get("size", size), o.get("color", color), o.get("bold", bold))


def tb(slide, x, y, w, h, paras, **kw):
    s = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    fill_text(s.text_frame, paras, **kw)
    return s


def box(slide, x, y, w, h, fill=CARD, line=None, radius=0.12, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
        text=None, dash=False, line_w=1.5, **kw):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = min(0.5, radius / min(w, h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
        if dash:
            s.line.dash_style = MSO_LINE.DASH
    if text is not None:
        kw.setdefault("align", PP_ALIGN.CENTER)
        kw.setdefault("anchor", MSO_ANCHOR.MIDDLE)
        fill_text(s.text_frame, text, **kw)
    return s


def placeholder(slide, page, x, y, w, h, label, size=18):
    PLACEHOLDERS.append((page, label))
    return box(slide, x, y, w, h, fill=RGBColor(0x27, 0x33, 0x3F), line=GRAY, dash=True,
               radius=0.08, text=label, size=size, color=MUTED)


# ---------------------------------------------------------------- 動畫引擎
class Anim:
    """收集一頁的動畫，最後寫成 PowerPoint 的 <p:timing>。
    群組 0 在進入投影片時自動播放；click() 之後的效果要按一下才播放。"""

    WIPE = {"left": (8, "wipe(left)"), "right": (2, "wipe(right)"),
            "bottom": (4, "wipe(down)"), "top": (1, "wipe(up)")}

    def __init__(self):
        self.groups = [[]]
        self.built = set()

    def click(self):
        self.groups.append([])

    def _add(self, start, kind, shapes, dur, **kw):
        if not isinstance(shapes, (list, tuple)):
            shapes = [shapes]
        for s in shapes:
            self.groups[-1].append(dict(start=int(start), kind=kind, spid=s.shape_id, dur=int(dur), **kw))
            if kind in ("wipe", "fade", "pulse"):
                self.built.add(s.shape_id)

    def wipe(self, shapes, start, dur=500, frm="left"):
        self._add(start, "wipe", shapes, dur, frm=frm)

    def fade(self, shapes, start, dur=400):
        self._add(start, "fade", shapes, dur)

    def motion(self, shapes, start, dur, p0, p1):
        """p0、p1：相對於版面原位置的位移（英吋）"""
        self._add(start, "motion", shapes, dur, p0=p0, p1=p1)

    def pulse(self, shapes, start, dur=350, times=3, scale=1.08):
        self._add(start, "pulse", shapes, dur, times=times, scale=scale)

    # ---- XML
    def _effect(self, e, node):
        i = self._id
        tgt = f'<p:tgtEl><p:spTgt spid="{e["spid"]}"/></p:tgtEl>'
        k = e["kind"]
        if k in ("wipe", "fade"):
            if k == "wipe":
                sub, flt = self.WIPE[e["frm"]]
                pid = 22
            else:
                sub, flt, pid = 0, "fade", 10
            x = (f'<p:par><p:cTn id="{i()}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" '
                 f'fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                 f'<p:childTnLst><p:set><p:cBhvr><p:cTn id="{i()}" dur="1" fill="hold"><p:stCondLst>'
                 f'<p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}<p:attrNameLst><p:attrName>style.visibility'
                 f'</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                 f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{i()}" dur="{e["dur"]}"/>'
                 f'{tgt}</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
        elif k == "motion":
            (x0, y0), (x1, y1) = e["p0"], e["p1"]
            path = f"M {x0 / W:.5f} {y0 / H:.5f} L {x1 / W:.5f} {y1 / H:.5f} E"
            x = (f'<p:par><p:cTn id="{i()}" presetID="0" presetClass="path" presetSubtype="0" '
                 f'accel="0" decel="0" fill="hold" nodeType="{node}"><p:stCondLst><p:cond delay="0"/>'
                 f'</p:stCondLst><p:childTnLst><p:animMotion origin="layout" path="{path}" '
                 f'pathEditMode="relative" ptsTypes="AA"><p:cBhvr><p:cTn id="{i()}" dur="{e["dur"]}" '
                 f'fill="hold"/>{tgt}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y'
                 f'</p:attrName></p:attrNameLst></p:cBhvr><p:rCtr x="0" y="0"/></p:animMotion>'
                 f'</p:childTnLst></p:cTn></p:par>')
        else:  # pulse
            sc = int(e["scale"] * 100000)
            x = (f'<p:par><p:cTn id="{i()}" presetID="6" presetClass="emph" presetSubtype="0" '
                 f'fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                 f'<p:childTnLst><p:animScale><p:cBhvr><p:cTn id="{i()}" dur="{e["dur"]}" autoRev="1" '
                 f'repeatCount="{e["times"] * 1000}" fill="hold"/>{tgt}</p:cBhvr>'
                 f'<p:by x="{sc}" y="{sc}"/></p:animScale></p:childTnLst></p:cTn></p:par>')
        return x

    def xml(self):
        n = [0]

        def nid():
            n[0] += 1
            return n[0]
        self._id = nid
        nid(); nid()  # 1 = root, 2 = mainSeq
        outer = []
        first = True
        for gi, g in enumerate(self.groups):
            if not g:
                continue
            auto = gi == 0
            starts = sorted(set(e["start"] for e in g))
            oid = nid()
            cond = '<p:cond delay="indefinite"/>'
            if auto:
                cond += '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>'
            inner = []
            for si, st in enumerate(starts):
                pid_ = nid()
                effs = []
                for ei, e in enumerate([e for e in g if e["start"] == st]):
                    if ei > 0:
                        node = "withEffect"
                    elif si == 0 and not auto:
                        node = "clickEffect"
                    else:
                        node = "afterEffect"
                    effs.append(self._effect(e, node))
                inner.append(f'<p:par><p:cTn id="{pid_}" fill="hold"><p:stCondLst><p:cond delay="{st}"/>'
                             f'</p:stCondLst><p:childTnLst>{"".join(effs)}</p:childTnLst></p:cTn></p:par>')
            outer.append(f'<p:par><p:cTn id="{oid}" fill="hold"><p:stCondLst>{cond}</p:stCondLst>'
                         f'<p:childTnLst>{"".join(inner)}</p:childTnLst></p:cTn></p:par>')
            first = False
        if first:
            return None
        bld = "".join(f'<p:bldP spid="{s}" grpId="0" animBg="1"/>' for s in sorted(self.built))
        return (f'<p:timing xmlns:p="{P}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
                f'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" '
                f'dur="indefinite" nodeType="mainSeq"><p:childTnLst>{"".join(outer)}</p:childTnLst></p:cTn>'
                f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
                f'</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>'
                f'</p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
                f'<p:bldLst>{bld}</p:bldLst></p:timing>')


def finish_slide(slide, anim, transition="wipe"):
    el = slide._element
    clr = el.find(f"{{{P}}}clrMapOvr")
    if transition == "fade":
        tr = f'<p:transition xmlns:p="{P}" spd="slow"><p:fade/></p:transition>'
    else:  # 由下往上擦去：像列印平台一層層升起
        tr = f'<p:transition xmlns:p="{P}" spd="med"><p:wipe dir="u"/></p:transition>'
    tr_el = etree.fromstring(tr)
    clr.addnext(tr_el)
    x = anim.xml()
    if x:
        tr_el.addnext(etree.fromstring(x))


# ---------------------------------------------------------------- FDM 噴嘴列印動畫
def print_object(slide, anim, cx, base_y, widths, layer_h, start=0, per_layer=250,
                 bed=True, colors=(ORANGE, ORANGE_D), nozzle_scale=1.0):
    """在 (cx, base_y) 由下往上「列印」一個物件。
    噴嘴沿 Z 字路徑來回移動，每一趟擠出一層（擦去方向與噴嘴移動同步）。回傳結束時間(ms)。"""
    t = start
    top = base_y - layer_h * len(widths)
    if bed:
        bw = max(widths) * 1.35
        b = box(slide, cx - bw / 2, base_y + 0.01, bw, max(0.05, layer_h * 0.45), fill=GRAY, radius=0.03)
        anim.wipe(b, t, 300, "left")
    layers = []
    for i, w in enumerate(widths):
        y = base_y - layer_h * (i + 1)
        s = box(slide, cx - w / 2, y + layer_h * 0.04, w, layer_h * 0.92,
                fill=colors[i % 2], radius=layer_h * 0.46)
        s.name = f"FDM 第{i + 1}層"
        layers.append((s, w, y))
    # 噴嘴：加熱塊＋噴頭，平常停在物件正上方（靜態檢視時也好看）
    nb_w, nb_h = 0.34 * nozzle_scale, 0.17 * nozzle_scale
    tip_w, tip_h = 0.16 * nozzle_scale, 0.12 * nozzle_scale
    park_tip_y = top - 0.12 * nozzle_scale
    blk = box(slide, cx - nb_w / 2, park_tip_y - tip_h - nb_h, nb_w, nb_h, fill=MUTED, radius=0.03)
    tip = box(slide, cx - tip_w / 2, park_tip_y - tip_h, tip_w, tip_h, fill=MUTED,
              shape=MSO_SHAPE.FLOWCHART_MERGE)
    blk.name, tip.name = "FDM 噴嘴加熱塊", "FDM 噴頭"
    noz = [blk, tip]
    anim.fade(noz, t, 250)
    t += 250
    travel = max(120, per_layer // 3)
    pos = (0.0, 0.0)  # 相對停靠位置的位移
    for i, (s, w, y) in enumerate(layers):
        go_right = i % 2 == 0
        xs = (cx - w / 2) if go_right else (cx + w / 2)
        xe = (cx + w / 2) if go_right else (cx - w / 2)
        dy = y - park_tip_y
        p_start = (xs - cx, dy)
        anim.motion(noz, t, travel, pos, p_start)  # 空跑到起點
        t += travel
        p_end = (xe - cx, dy)
        anim.motion(noz, t, per_layer, p_start, p_end)  # 擠出一層
        anim.wipe(s, t, per_layer, "left" if go_right else "right")
        t += per_layer
        pos = p_end
    anim.motion(noz, t, 350, pos, (0.0, 0.0))  # 回到停靠點
    return t + 350


def mini_printer(slide, anim, start=0):
    """每頁右下角的小列印動畫（品牌動態標誌）"""
    return print_object(slide, anim, 12.42, 7.12, [0.62, 0.74, 0.8, 0.74, 0.62], 0.1,
                        start=start, per_layer=210, nozzle_scale=0.6)


# ---------------------------------------------------------------- 共用版型
def new_slide(prs, sec, title=None, title_size=38):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    anim = Anim()
    tag = box(s, W - 0.5 - 3.6, 0.32, 3.6, 0.46, fill=None, line=ORANGE, radius=0.23,
              text=SEC[sec], size=16, color=ORANGE, bold=True)
    tag.name = "段落時間標籤"
    anim.fade(tag, 0, 300)
    t_shape = None
    if title:
        t_shape = tb(s, 0.6, 0.4, 8.9, 0.9, title, size=title_size, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        t_shape.name = "標題"
        anim.wipe(t_shape, 0, 600, "left")  # 擠出：標題像噴嘴擠線一樣畫出
    return s, anim


def worksheet_tag(slide, anim, part, start=0):
    t = box(slide, 0.6, 6.78, 2.75, 0.48, fill=ORANGE, radius=0.1, text=f"📝 學習單 Part {part}",
            size=18, bold=True, color=DARK)
    t.name = "學習單標籤"
    anim.wipe(t, start, 400, "left")
    return t


def progress_bar(slide, anim, x, y, w, minutes, ticks, start, label, marks=None):
    """列印進度條：橘色線材在實際倒數時間內擠滿"""
    tb(slide, x, y - 0.5, w, 0.45, label, size=20, color=MUTED, anchor=MSO_ANCHOR.BOTTOM, margin=0)
    track = box(slide, x, y, w, 0.3, fill=CARD2, radius=0.15)
    bar = box(slide, x, y, w, 0.3, fill=ORANGE, radius=0.15)
    bar.name = f"列印進度條（{minutes} 分鐘）"
    anim.wipe(bar, start, minutes * 60 * 1000, "left")
    for tk in ticks:
        tx = x + w * tk / minutes
        txt = f"{tk} 分" if tk else "開始"
        if marks and tk in marks:
            txt = marks[tk]
        tb(slide, tx - 0.8, y + 0.34, 1.6, 0.38, txt, size=16, color=MUTED, align=PP_ALIGN.CENTER, margin=0)
    return bar


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text.strip()


def arrow(slide, x, y, w, h, color=ORANGE):
    a = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    a.shadow.inherit = False
    a.fill.solid(); a.fill.fore_color.rgb = color
    a.line.fill.background()
    return a


def down_arrow(slide, x, y, w, h, color=ORANGE):
    a = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    a.shadow.inherit = False
    a.fill.solid(); a.fill.fore_color.rgb = color
    a.line.fill.background()
    return a


# ================================================================ 各頁內容
def build(out):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    pg = [0]

    def page():
        pg[0] += 1
        return pg[0]

    # ---------------- 1 封面
    n = page()
    s, an = new_slide(prs, "open")
    end = print_object(s, an, 10.15, 6.35,
                       [2.5, 2.8, 3.0, 3.1, 3.0, 2.75, 2.45, 2.25, 2.25, 2.45, 2.7],
                       0.34, start=200, per_layer=330)
    t1 = tb(s, 0.6, 1.75, 6.4, 1.0, "從零開始的 3D 列印", size=40, bold=True)
    t2 = tb(s, 0.6, 2.75, 6.4, 1.3, ["概念、產業到", "MakerWorld 實作"], size=32, bold=True, color=ORANGE)
    t3 = tb(s, 0.6, 4.5, 6.2, 1.1, ["日期：＿＿＿＿＿＿＿＿", "主辦單位：＿＿＿＿＿＿"], size=20, color=MUTED,
            space_after=6)
    t4 = tb(s, 0.6, 5.85, 6.0, 0.5, "FDM 入門工作坊｜2 小時", size=20, color=GRAY)
    an.wipe(t1, 400, 700, "left")
    an.wipe(t2, 1000, 800, "left")
    an.fade([t3, t4], 1700, 600)
    notes(s, """
【時間】約 0.5 分鐘（講者 A）
【這頁要講】歡迎大家！今天兩小時從「3D 列印是什麼」一路做到「在 MakerWorld 挑出一個能真的送印的模型」。
【動畫】右邊的噴嘴會一層層把物件「印」出來——這就是今天的主角 FDM 列印方式，可以邊播邊說：「等一下你們會看到真的機台這樣印。」
【互動】請大家看一下桌上的學習單與樣品盒，先不要打開樣品盒。
""")
    finish_slide(s, an, "fade")

    # ---------------- 2 今天的行程（時間軸按實際分鐘比例）
    n = page()
    s, an = new_slide(prs, "open", "今天的行程")
    stages = [("開場", 0, 10, False), ("基礎概念", 10, 22, False), ("業界應用", 22, 34, False),
              ("機台參觀", 34, 44, True), ("MakerWorld 實作", 44, 92, True), ("小組發表", 92, 108, True),
              ("總結", 108, 120, False)]
    x0, x1, ly = 0.9, 12.45, 3.75

    def X(m):
        return x0 + (x1 - x0) * m / 120
    t = 500
    for i, (name, a_, b_, hands) in enumerate(stages):
        seg = box(s, X(a_) + 0.03, ly, X(b_) - X(a_) - 0.06, 0.3, fill=ORANGE if hands else GRAY, radius=0.15)
        dur = int(140 + 22 * (b_ - a_))
        an.wipe(seg, t, dur, "left")  # 擠出：時間軸像一條線材被擠出來
        mid = (X(a_) + X(b_)) / 2
        above = i % 2 == 0
        ly_txt = ly - 1.25 if above else ly + 0.5
        lw = 3.2 if "Maker" in name else 2.3
        lab = tb(s, mid - lw / 2, ly_txt, lw, 1.0,
                 [[(name, {"bold": True, "size": 22})], [(f"{a_}–{b_} 分", {"color": MUTED, "size": 20})]],
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM if above else MSO_ANCHOR.TOP, margin=0)
        an.fade(lab, t + dur // 2, 300)
        t += dur
    lg1 = box(s, 0.9, 5.65, 0.4, 0.26, fill=GRAY, radius=0.13)
    lg1t = tb(s, 1.4, 5.5, 2.0, 0.55, "聽講", size=20, color=MUTED, anchor=MSO_ANCHOR.MIDDLE, margin=0)
    lg2 = box(s, 2.6, 5.65, 0.4, 0.26, fill=ORANGE, radius=0.13)
    lg2t = tb(s, 3.1, 5.5, 5.5, 0.55, "動手做：超過一半的時間！", size=20, color=MUTED,
              anchor=MSO_ANCHOR.MIDDLE, margin=0)
    an.fade([lg1, lg1t, lg2, lg2t], t, 400)
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 1 分鐘（講者 A）
【這頁要講】時間軸是照實際分鐘比例畫的：灰色是聽講、橘色是動手。強調「聽講只有大約 24 分鐘，其他都是你們動手」。
最長的一段是 MakerWorld 實作（48 分鐘），最後每組要上台發表自己選的送印模型。
【互動】問一句：「有人以前用過 3D 印表機嗎？舉手讓我看看。」快速掃一眼人數即可，不展開。
""")
    finish_slide(s, an)

    # ---------------- 3 認識你的小組
    n = page()
    s, an = new_slide(prs, "open", "認識你的小組")
    cw, gap, cy = 2.3, 0.2, 1.6
    for i in range(5):
        x = 0.6 + i * (cw + gap)
        c = box(s, x, cy, cw, 2.75, fill=CARD, radius=0.15)
        num = box(s, x + cw / 2 - 0.42, cy + 0.25, 0.84, 0.84, fill=ORANGE, shape=MSO_SHAPE.OVAL,
                  text=str(i + 1), size=30, bold=True, color=DARK)
        txt = tb(s, x + 0.1, cy + 1.2, cw - 0.2, 1.45,
                 [[(f"第 {i + 1} 組", {"bold": True, "size": 26})],
                  [("認養工作人員", {"size": 20, "color": MUTED})],
                  [("＿＿＿＿＿", {"size": 22, "color": ORANGE})]],
                 align=PP_ALIGN.CENTER, margin=0)
        an.wipe([c, num, txt], 500 + i * 260, 450, "bottom")  # 堆疊：卡片由下往上長出
    msg = box(s, 0.6, 4.7, 12.1, 0.9, fill=ORANGE_BG, radius=0.15,
              text=[[("有問題？舉手找", {}), ("「你們組的工作人員」", {"color": ORANGE, "bold": True})]],
              size=26)
    roles = tb(s, 0.6, 5.75, 11.0, 0.5, "A 主講｜B 業界講者・控台｜C 列印站｜D MakerWorld｜E 行政",
               size=20, color=MUTED, margin=0)
    an.wipe(msg, 1900, 600, "left")
    an.fade(roles, 2400, 400)
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 1 分鐘（講者 A）
【這頁要講】25 位學員分 5 組、每組 5 人，每組有一位認養的工作人員（名字請事先填在卡片上）。
有任何問題——操作卡住、找不到模型、想去列印站看——先找自己組的工作人員，不用等主講。
【互動】請 5 位工作人員依序舉手、說一句自我介紹（各 5 秒），讓各組認得自己的工作人員。
""")
    finish_slide(s, an)

    # ---------------- 4 安全須知（醒目）
    n = page()
    s, an = new_slide(prs, "open", [[("⚠ ", {"color": ORANGE}), ("安全須知", {})]])
    t1c = box(s, 0.6, 1.55, 5.95, 1.6, fill=ORANGE, radius=0.18,
              text=[[("噴嘴", {"size": 24, "bold": True})], [("約 200–250°C", {"size": 40, "bold": True})]],
              color=DARK)
    t2c = box(s, 6.75, 1.55, 5.95, 1.6, fill=ORANGE, radius=0.18,
              text=[[("熱床", {"size": 24, "bold": True})], [("約 55–100°C", {"size": 40, "bold": True})]],
              color=DARK)
    an.wipe([t1c, t2c], 500, 500, "bottom")
    an.pulse([t1c, t2c], 1100, 320, times=3, scale=1.05)  # 預熱：溫度卡片脈動
    rules = [("✋", "不要伸手進機台"), ("↔", "A1 熱床會前後移動，保持一臂距離"), ("!", "取件和工具只由工作人員操作")]
    for i, (ic, txt) in enumerate(rules):
        y = 3.45 + i * 0.95
        row = box(s, 0.6, y, 12.1, 0.8, fill=CARD, radius=0.12)
        icon = box(s, 0.8, y + 0.1, 0.6, 0.6, fill=ORANGE_BG, shape=MSO_SHAPE.OVAL, text=ic, size=20,
                   color=ORANGE, bold=True)
        tx = tb(s, 1.65, y, 10.8, 0.8, txt, size=26, bold=True, anchor=MSO_ANCHOR.MIDDLE, margin=0)
        an.wipe([row, icon, tx], 1300 + i * 350, 450, "left")
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 1.5 分鐘（講者 A）— 這頁一定要講清楚、講慢一點
【這頁要講】
1. 噴嘴大約 200–250°C、熱床大約 55–100°C，碰到會燙傷，印完也要一段時間才會降溫。
2. 機台運作中不要把手伸進去。
3. A1 是開放式機台，熱床會「前後」移動，參觀時請保持一臂距離，也不要把東西放在機台前面。
4. 取件、拿刮板和工具都只由工作人員操作。
【互動】請全體一起複誦一次：「不伸手、一臂距離、工作人員取件」。
""")
    finish_slide(s, an)

    # ---------------- 5 暖身問題
    n = page()
    s, an = new_slide(prs, "open")
    q = tb(s, 0.8, 1.5, 11.7, 2.4, ["如果今天可以印任何東西，", "你想印什麼？"], size=48, bold=True,
           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    sub = tb(s, 0.8, 3.95, 11.7, 0.6, "寫在學習單 Part 0，限時 1 分鐘", size=24, color=ORANGE,
             align=PP_ALIGN.CENTER, bold=True)
    an.wipe(q, 300, 1200, "left")
    an.fade(sub, 1500, 400)
    progress_bar(s, an, 2.2, 5.35, 8.9, 1, [0, 1], 1900, "⏱ 列印進度：橘色線材跑滿＝時間到")
    worksheet_tag(s, an, 0, 1500)
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 2 分鐘（講者 A）：說明 20 秒＋學員寫 1 分鐘＋分享 30 秒
【這頁要講】請大家在學習單 Part 0 寫下：如果今天可以印任何東西，你想印什麼？沒有標準答案，越具體越好。
畫面上的橘色進度條會在 1 分鐘內跑滿，跑滿就是時間到。
【互動】時間到後請 2–3 位自願者分享，主講可以回應「這個今天就有機會在 MakerWorld 找到！」
【提醒】這時 C 正在列印站啟動 6 台示範列印（4 台 A1＋2 台 P2S），講者不用等，直接進下一段；之後總結時會展示成品。
""")
    finish_slide(s, an)

    # ---------------- 6 什麼是 3D 列印
    n = page()
    s, an = new_slide(prs, "basic", "什麼是 3D 列印？")
    for side, (x, head, line, img, frm) in enumerate([
        (0.6, "積層製造｜3D 列印", "材料一層一層「加上去」", "圖：一層層堆疊成形示意", "bottom"),
        (6.95, "切削加工｜CNC", "把整塊材料「削掉」", "圖：CNC 切削加工示意", "top")]):
        c = box(s, x, 1.5, 5.8, 4.65, fill=CARD if side else ORANGE_BG, radius=0.18)
        h_ = tb(s, x + 0.3, 1.65, 5.2, 0.6, head, size=28, bold=True, color=ORANGE if side == 0 else WHITE)
        l_ = tb(s, x + 0.3, 2.25, 5.2, 0.55, line, size=24, color=MUTED)
        ph = placeholder(s, n, x + 0.3, 2.95, 5.2, 2.95, img)
        st = 500 + side * 1300
        an.wipe(c, st, 500, frm)
        an.wipe([h_, l_], st + 300, 500, "left")
        # 積層＝由下往上長出；切削＝由上往下（材料被「削」出來）
        an.wipe(ph, st + 600, 700, frm)
    vs = box(s, 6.2, 3.4, 0.9, 0.9, fill=ORANGE, shape=MSO_SHAPE.OVAL, text="vs", size=24, bold=True,
             color=DARK)
    an.fade(vs, 1650, 300)
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 1.5 分鐘（講者 A）
【這頁要講】3D 列印的正式名稱叫「積層製造」：材料一層一層加上去，需要的地方才放材料。
對照的是傳統的「切削加工」（例如 CNC）：從一整塊材料把不要的部分削掉。
一句話記：3D 列印是「加法」，CNC 是「減法」。
【動畫】左邊圖框由下往上長出（加法），右邊由上往下出現（減法），可以點一下這個對比。
【互動】問：「你覺得哪一種比較浪費材料？為什麼？」等一兩位回答即可。
""")
    finish_slide(s, an)

    # ---------------- 7 四步驟流程
    n = page()
    s, an = new_slide(prs, "basic", "3D 列印四步驟")
    steps = [("建模", "用 CAD 畫出 3D 模型"), ("切片", "軟體切成一層層路徑"),
             ("列印", "機台逐層擠出成形"), ("後處理", "拆支撐、修整表面")]
    bw, bg_ = 2.6, 0.57
    t = 500
    for i, (nm, ds) in enumerate(steps):
        x = 0.6 + i * (bw + bg_)
        c = box(s, x, 2.0, bw, 3.3, fill=CARD, radius=0.18)
        num = box(s, x + bw / 2 - 0.4, 2.25, 0.8, 0.8, fill=ORANGE, shape=MSO_SHAPE.OVAL, text=str(i + 1),
                  size=28, bold=True, color=DARK)
        nm_ = tb(s, x + 0.15, 3.2, bw - 0.3, 0.7, nm, size=32, bold=True, align=PP_ALIGN.CENTER)
        ds_ = tb(s, x + 0.2, 3.95, bw - 0.4, 1.2, ds, size=22, color=MUTED, align=PP_ALIGN.CENTER)
        an.wipe([c, num, nm_, ds_], t, 500, "bottom")
        t += 450
        if i < 3:
            ar = arrow(s, x + bw + 0.08, 3.45, bg_ - 0.16, 0.42)
            an.wipe(ar, t, 250, "left")
            t += 200
    tip = tb(s, 0.6, 5.65, 11.0, 0.55, "這四步就是今天下午你會走過的流程", size=22, color=ORANGE, margin=0)
    an.fade(tip, t + 100, 400)
    worksheet_tag(s, an, 1, t)
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 2 分鐘（講者 A）
【這頁要講】3D 列印不是按一個鍵就好，有四個步驟：
1. 建模：用 CAD 軟體畫出 3D 模型（今天我們跳過，直接用 MakerWorld 別人做好的模型）。
2. 切片：用軟體（今天用 Bambu Studio）把模型切成一層層、算出噴嘴要走的路徑。
3. 列印：機台照著路徑一層層擠出線材。
4. 後處理：拆支撐、修整表面。
【互動】請學員在學習單 Part 1 把四個步驟依序填上，給 30 秒；搶答時間會考這題。
""")
    finish_slide(s, an)

    # ---------------- 8 FDM 怎麼運作
    n = page()
    s, an = new_slide(prs, "basic", "FDM 怎麼運作？")
    chain = [("線材", "捲在料盤上的塑膠線"), ("送料", "齒輪把線材往前推"),
             ("加熱融化", "噴嘴約 200–250°C"), ("堆疊", "畫完一層，再往上一層")]
    t = 500
    for i, (nm, ds) in enumerate(chain):
        y = 1.55 + i * 1.2
        row = box(s, 0.6, y, 6.3, 0.95, fill=CARD, radius=0.14)
        num = box(s, 0.78, y + 0.15, 0.65, 0.65, fill=ORANGE, shape=MSO_SHAPE.OVAL, text=str(i + 1), size=22,
                  bold=True, color=DARK)
        tx = tb(s, 1.65, y, 5.1, 0.95, [[(nm + "　", {"bold": True, "size": 26}), (ds, {"color": MUTED, "size": 21})]],
                anchor=MSO_ANCHOR.MIDDLE, margin=0)
        # 線材是往下送進噴嘴的，所以這頁由上往下依序出現
        an.wipe([row, num, tx], t, 450, "top")
        t += 450
        if i < 3:
            da = down_arrow(s, 0.98, y + 0.96, 0.3, 0.23)
            an.wipe(da, t - 100, 200, "top")
    ph = placeholder(s, n, 7.3, 1.55, 5.4, 4.6, "圖：FDM 噴嘴擠出示意\n（自行拍攝）", size=20)
    an.wipe(ph, 600, 800, "bottom")
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 1.5 分鐘（講者 A）
【這頁要講】FDM（熔融沉積成型）是今天 6 台機器都在用的方式，原理很像「會自己移動的熱熔膠槍」：
線材從料盤送出 → 送料齒輪把線往前推 → 噴嘴加熱到約 200–250°C 把線材融化擠出 → 畫完一層，平台下降（或噴頭上升）再畫下一層。
【動畫】每頁右下角那個小噴嘴就是在示範這件事：來回走一趟＝印一層。
【互動】可以拿起一卷線材或一根線材給前排摸一下。
""")
    finish_slide(s, an)

    # ---------------- 9 三種技術比較（只帶過）
    n = page()
    s, an = new_slide(prs, "basic", "三種技術一頁比較")
    cols = ["", "原理", "常見材料", "優點", "適合用途"]
    rows = [("FDM", "擠出熔融線材", "PLA、PETG、TPU", "便宜、好上手", "原型、教具"),
            ("SLA", "雷射固化樹脂", "光敏樹脂", "細節精緻", "珠寶、牙科"),
            ("SLS", "雷射燒結粉末", "尼龍粉末", "免支撐、強度高", "功能零件")]
    cwid = [1.5, 2.4, 3.2, 2.5, 2.5]
    xs = [0.6]
    for c_ in cwid[:-1]:
        xs.append(xs[-1] + c_)
    hdr = []
    for j, c_ in enumerate(cols):
        if c_:
            hdr.append(tb(s, xs[j], 1.55, cwid[j], 0.6, c_, size=22, bold=True, color=ORANGE,
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE))
    an.wipe(hdr, 500, 500, "left")
    t = 900
    for i, r in enumerate(rows):
        y = 2.25 + i * 1.15
        bgc = ORANGE_BG if i == 0 else CARD
        band = box(s, 0.6, y, 12.1, 1.0, fill=bgc, radius=0.12)
        cells = [band]
        for j, val in enumerate(r):
            cells.append(tb(s, xs[j], y, cwid[j], 1.0, val, size=26 if j == 0 else 21, bold=(j == 0),
                            color=ORANGE if (j == 0 and i == 0) else WHITE, align=PP_ALIGN.CENTER,
                            anchor=MSO_ANCHOR.MIDDLE))
        an.wipe(cells, t, 450, "bottom")
        t += 350
    foot = tb(s, 0.6, 5.85, 11.0, 0.5, "★ 今天的機台都是 FDM，另外兩種知道就好", size=20, color=MUTED, margin=0)
    an.fade(foot, t + 100, 400)
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 1 分鐘（講者 A）— 這頁只帶過，不細講
【這頁要講】3D 列印不只一種：
FDM 擠出熔融線材，便宜好上手，今天用的就是它。
SLA 用光固化液態樹脂，細節很精緻，常用在珠寶、牙科。
SLS 用雷射燒結尼龍粉末，不需要支撐、強度高，適合功能零件。
【互動】不做互動，直接進下一頁「動手摸樣品」。
""")
    finish_slide(s, an)

    # ---------------- 10 動手摸樣品（3 分鐘）
    n = page()
    s, an = new_slide(prs, "basic", "動手摸樣品")
    lead = tb(s, 0.6, 1.4, 11.5, 0.6, "打開樣品盒，組內傳著看、摸、比一比", size=24, color=ORANGE, bold=True,
              margin=0)
    an.wipe(lead, 400, 600, "left")
    items = [("層高", "看側面層紋的粗細"), ("填充", "看內部的格狀結構"),
             ("支撐", "找拆掉支撐的痕跡"), ("材料", "比 PLA／PETG／TPU 軟硬")]
    for i, (k, v) in enumerate(items):
        x = 0.6 + (i % 2) * 6.15
        y = 2.2 + (i // 2) * 1.35
        c = box(s, x, y, 5.95, 1.15, fill=CARD, radius=0.14)
        tx = tb(s, x + 0.3, y, 5.5, 1.15, [[(k + "　", {"bold": True, "color": ORANGE, "size": 26}),
                                              (v, {"size": 22})]], anchor=MSO_ANCHOR.MIDDLE, margin=0)
        an.wipe([c, tx], 900 + i * 300, 450, "bottom")
    progress_bar(s, an, 0.6, 5.5, 10.6, 3, [0, 1, 2, 3], 2200, "⏱ 3 分鐘倒數：橘色線材擠滿＝時間到")
    worksheet_tag(s, an, 2, 2000)
    mini_printer(s, an, 200)
    notes(s, """
【時間】3 分鐘動手＋約 30 秒收尾（講者 A，各組工作人員協助）
【這頁要講】請各組打開樣品盒，傳著看、摸，比較四件事，並把觀察寫在學習單 Part 2：
1. 層高：側面的層紋粗細（細層高比較光滑、但印比較久）。
2. 填充：有剖開的樣品可以看內部格狀結構。
3. 支撐：找找看哪裡有拆掉支撐留下的痕跡。
4. 材料：PLA 硬、PETG 有點韌、TPU 軟到可以彎。
【互動】進度條會在 3 分鐘內跑滿。時間到請各組用一句話說「哪個樣品最讓你意外」。
【提醒】請各組工作人員確認樣品盒收回、零件沒有少。
""")
    finish_slide(s, an)

    # ---------------- 11–13 業界案例
    cases = [
        ("案例 1｜醫療", "依病患影像，印出專屬手術模擬模型", ["客製器械", "手術模擬模型"],
         "圖：依病患影像印出的手術模擬模型"),
        ("案例 2｜製造業", "產線治具夾具，快速又便宜地客製", ["治具", "夾具", "快速改版"],
         "圖：產線治具／夾具實例"),
        ("案例 3｜文創與消費品", "小量生產也划算，商品可以客製", ["小量生產", "客製商品"],
         "圖：文創／客製商品實例"),
    ]
    case_notes = [
        """
【時間】約 2 分鐘（講者 B）
【這頁要講】醫療是 3D 列印很成熟的應用：
1. 客製器械：依每個人的身體尺寸做出合適的輔具或器械。
2. 手術模擬模型：把病患的 CT／MRI 影像轉成 3D 模型印出來，醫師可以在手術前拿在手上看、練習，降低手術風險。
重點一句話：「每個病人都不一樣，3D 列印剛好擅長『只做一個』。」
【講者可穿插自己的原型開發經驗】
（B 請在這裡分享一段自己的經驗，例如：第一次用 3D 列印做原型是什麼東西？改了幾版？比傳統做法省了多少時間？）
【互動】問：「你覺得為什麼醫療特別適合用 3D 列印？」
""",
        """
【時間】約 2 分鐘（講者 B）
【這頁要講】工廠產線上常需要「治具」與「夾具」：固定零件、對位、檢測用的小工具。
傳統要找人開模或 CNC 加工，貴又要等；用 3D 列印可以今天畫、明天就裝上產線，不合用就改一版再印。
重點一句話：「快、便宜、可以一直改。」
【講者可穿插自己的原型開發經驗】
（B 請分享：在業界做治具／原型時，3D 列印幫你縮短了多少開發時間？有沒有失敗的例子？）
【互動】請學員想一想：「你身邊有沒有什麼東西，如果可以客製一個小工具會更好用？」
""",
        """
【時間】約 2 分鐘（講者 B）
【這頁要講】文創與消費品：以前要開模就得一次做幾千個，現在可以只做 10 個、50 個，也能每一件都不一樣（刻名字、改尺寸）。
很多設計師在 MakerWorld 分享模型、也販售客製商品——這也是等一下要講「授權」的原因。
重點一句話：「小量也划算，每一件都能不一樣。」
【講者可穿插自己的原型開發經驗】
（B 請分享：有沒有看過或做過把 3D 列印做成商品的例子？從原型到商品中間遇到什麼問題？）
【互動】問：「如果要賣自己印的東西，你覺得要注意什麼？」（為後面的授權頁鋪梗）
""",
    ]
    for ci, (title, key, chips, img) in enumerate(cases):
        n = page()
        s, an = new_slide(prs, "ind", title)
        ph = placeholder(s, n, 0.6, 1.5, 7.1, 4.85, img, size=20)
        an.wipe(ph, 400, 800, "bottom")
        k = tb(s, 8.1, 1.7, 4.6, 2.2, key, size=30, bold=True, anchor=MSO_ANCHOR.TOP, margin=0,
               line_spacing=1.1)
        an.wipe(k, 1100, 800, "left")
        cy = 4.15
        for j, ch in enumerate(chips):
            c = box(s, 8.1, cy + j * 0.7, 3.4, 0.56, fill=None, line=ORANGE, radius=0.28, text=ch, size=20,
                    color=ORANGE, bold=True)
            an.wipe(c, 1800 + j * 250, 350, "left")
        mini_printer(s, an, 200)
        notes(s, case_notes[ci])
        finish_slide(s, an)

    # ---------------- 14 搶答時間（按一下出下一題）
    n = page()
    s, an = new_slide(prs, "ind", "搶答時間！")
    qs = [("Q1", "3D 列印屬於哪一種製造方式？", ["A　切削加工", "B　積層製造", "C　射出成型"]),
          ("Q2", "四步驟中，「切片」是第幾步？", ["A　第 1 步", "B　第 2 步", "C　第 4 步"]),
          ("Q3", "製造業常用 3D 列印做什麼？", ["A　產線治具與夾具", "B　大量生產零件", "C　取代所有模具"])]
    hint = tb(s, 0.6, 1.35, 11.0, 0.5, "舉手搶答，答對的組加分！", size=22, color=ORANGE, bold=True, margin=0)
    an.fade(hint, 400, 400)
    for i, (qn, qt, opts) in enumerate(qs):
        x = 0.6 + i * 4.1
        if i > 0:
            an.click()  # 按一下出下一題
        c = box(s, x, 2.0, 3.9, 3.9, fill=CARD, radius=0.18)
        qn_ = tb(s, x + 0.25, 2.15, 3.4, 0.6, qn, size=30, bold=True, color=ORANGE, margin=0)
        qt_ = tb(s, x + 0.25, 2.8, 3.45, 1.3, qt, size=22, bold=True, margin=0)
        op_ = tb(s, x + 0.25, 4.2, 3.45, 1.9, opts, size=21, color=MUTED, margin=0, space_after=8)
        st = 600 if i == 0 else 0
        an.wipe([c, qn_], st, 400, "bottom")
        an.wipe([qt_, op_], st + 300, 500, "left")
        if i == 0:
            mini_printer(s, an, 200)
    worksheet_tag(s, an, 3, 300)
    notes(s, """
【時間】約 3 分鐘（講者 B 主持，A 協助記分）
【操作】進入此頁會先出現 Q1；按一下滑鼠出現 Q2，再按一下出現 Q3。
【題目與答案】（答案只寫在這裡，不顯示在投影片上）
Q1：3D 列印屬於哪一種製造方式？　答案：B 積層製造（對應 Part 1 第 6 頁）
Q2：四步驟（建模→切片→列印→後處理）中，「切片」是第幾步？　答案：B 第 2 步（對應 Part 1 第 7 頁）
Q3：製造業常用 3D 列印做什麼？　答案：A 產線治具與夾具（對應 Part 3 案例 2）
【互動】每題請最先舉手的組回答，答對加 1 分；答錯可開放其他組搶答。最後提醒大家把答案補寫到學習單 Part 3。
""")
    finish_slide(s, an)

    # ---------------- 15 休息與參觀分流（10 分鐘）
    n = page()
    s, an = new_slide(prs, "tour", "休息與參觀分流")
    l_ = box(s, 0.6, 1.5, 5.4, 1.75, fill=ORANGE_BG, radius=0.16,
             text=[[("第 1–3 組", {"size": 32, "bold": True, "color": ORANGE})],
                   [("先參觀列印站（C 帶領）", {"size": 22})]])
    r_ = box(s, 7.3, 1.5, 5.4, 1.75, fill=CARD, radius=0.16,
             text=[[("第 4–5 組", {"size": 32, "bold": True})], [("先休息、喝水、上廁所", {"size": 22})]])
    sw = box(s, 6.15, 1.95, 1.0, 0.85, fill=None, shape=MSO_SHAPE.LEFT_RIGHT_ARROW, line=None)
    sw.fill.solid(); sw.fill.fore_color.rgb = ORANGE
    swt = tb(s, 5.3, 3.3, 2.7, 0.45, "5 分鐘後交換", size=20, bold=True, color=ORANGE, align=PP_ALIGN.CENTER,
             margin=0)
    an.wipe(l_, 400, 500, "bottom")
    an.wipe(r_, 700, 500, "bottom")
    an.wipe([sw, swt], 1100, 400, "left")
    cl = tb(s, 0.6, 3.85, 3.0, 0.6, "參觀時填 Part 4：", size=22, bold=True, anchor=MSO_ANCHOR.MIDDLE, margin=0)
    an.fade(cl, 1400, 300)
    for j, ch in enumerate(["噴嘴", "熱床", "送料機構", "A1 vs P2S 差異"]):
        wch = 3.0 if j == 3 else 1.7
        xch = 3.6 + sum([1.9, 1.9, 1.9][:j]) if j < 3 else 3.6 + 5.7
        c = box(s, xch, 3.87, wch, 0.56, fill=None, line=ORANGE, radius=0.28, text=ch, size=20, bold=True,
                color=ORANGE)
        an.wipe(c, 1600 + j * 200, 300, "left")
    progress_bar(s, an, 0.6, 5.45, 10.6, 10, [0, 5, 10], 2300, "⏱ 10 分鐘倒數（中間點＝交換）",
                 marks={5: "5 分：交換！", 10: "10 分"})
    worksheet_tag(s, an, 4, 1400)
    mini_printer(s, an, 200)
    notes(s, """
【時間】10 分鐘（列印站由 C 負責，E 協助分流與計時）
【這頁要講】接下來分兩梯：第 1–3 組先跟 C 去列印站參觀、第 4–5 組先休息；5 分鐘後交換。
參觀時請填學習單 Part 4：找到噴嘴、熱床、送料機構在哪裡，並比較 A1（開放式）和 P2S（封閉式）的差異。
【安全再提醒】A1 熱床會前後移動，保持一臂距離；不要伸手進機台。
【互動】C 可以在 A1 前面問：「你覺得封閉式的 P2S 多了一個外殼，有什麼好處？」（溫度穩定、比較安全、適合 PETG 等材料）
【操作】進度條 10 分鐘跑滿，跑到正中間時提醒交換。B 在控台留意時間。
""")
    finish_slide(s, an)

    # ---------------- 16 MakerWorld 是什麼
    n = page()
    s, an = new_slide(prs, "mw", "MakerWorld 是什麼？")
    d = tb(s, 0.6, 1.45, 6.4, 1.0, [[("Bambu Lab 的", {}), ("3D 模型分享平台", {"color": ORANGE, "bold": True})]],
           size=28, margin=0, anchor=MSO_ANCHOR.MIDDLE)
    an.wipe(d, 400, 700, "left")
    feats = [("🔍", "搜尋", "找到想印的模型"), ("⬇", "下載", "取得模型檔案"),
             ("⚙", "直接開啟", "套用建議的列印設定")]
    for i, (ic, k, v) in enumerate(feats):
        y = 2.7 + i * 1.2
        c = box(s, 0.6, y, 6.3, 1.0, fill=CARD, radius=0.14)
        icn = box(s, 0.8, y + 0.17, 0.66, 0.66, fill=ORANGE_BG, shape=MSO_SHAPE.OVAL, text=ic, size=20,
                  color=ORANGE, bold=True)
        tx = tb(s, 1.7, y, 5.1, 1.0, [[(k + "　", {"bold": True, "size": 26}), (v, {"color": MUTED, "size": 22})]],
                anchor=MSO_ANCHOR.MIDDLE, margin=0)
        an.wipe([c, icn, tx], 1000 + i * 350, 450, "bottom")
    ph = placeholder(s, n, 7.3, 1.5, 5.4, 4.65, "截圖：MakerWorld 首頁", size=20)
    an.wipe(ph, 600, 800, "bottom")
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 1 分鐘（講者 D）
【這頁要講】MakerWorld 是 Bambu Lab 經營的 3D 模型分享平台。上面有大量別人設計好的模型，可以：
1. 搜尋：用關鍵字和篩選找到想印的東西；
2. 下載：取得模型檔；
3. 直接用作者「建議的列印設定」在 Bambu Studio 開啟——不用自己從零調參數。
這就是為什麼今天我們可以跳過「建模」，直接從「切片」開始。
【互動】問：「剛剛 Part 0 寫想印什麼的同學，等一下就試著搜搜看！」
""")
    finish_slide(s, an)

    # ---------------- 17 示範步驟（按一下出下一步）
    n = page()
    s, an = new_slide(prs, "mw", "示範步驟")
    demo = [("① 登入", "登入帳號", "截圖：登入畫面"),
            ("② 搜尋與篩選", "關鍵字＋篩選", "截圖：搜尋與篩選"),
            ("③ 看模型頁", ["時間、克數", "設定、授權"], "截圖：模型頁資訊"),
            ("④ 在 Bambu Studio 開啟", ["一鍵開啟", "切片軟體"], "截圖：Bambu Studio 開啟"),
            ("⑤ 認識參數", ["層高、填充", "支撐"], "截圖：層高／填充／支撐設定")]
    cw = 2.3
    for i, (nm, ds, img) in enumerate(demo):
        x = 0.6 + i * (cw + 0.15)
        if i > 0:
            an.click()
        ph = placeholder(s, n, x, 1.55, cw, 1.9, img, size=16)
        c = box(s, x, 3.6, cw, 2.55, fill=CARD, radius=0.14)
        nm_ = tb(s, x + 0.12, 3.72, cw - 0.24, 0.95, nm, size=22, bold=True, color=ORANGE, margin=0)
        ds_ = tb(s, x + 0.12, 4.7, cw - 0.24, 1.4, ds, size=20, color=WHITE, margin=0)
        st = 500 if i == 0 else 0
        an.wipe(ph, st, 500, "bottom")
        an.wipe(c, st + 200, 450, "bottom")
        an.wipe([nm_, ds_], st + 400, 450, "left")
        if i == 0:
            mini_printer(s, an, 200)
    notes(s, """
【時間】約 4 分鐘（講者 D 邊講邊在投影幕上實際操作）
【操作】進入此頁先出現步驟①；之後每按一下出現下一步，配合現場示範的節奏。
【這頁要講】
① 登入：登入 MakerWorld 帳號（之後才能在 Bambu Studio 直接開啟）。
② 搜尋與篩選：輸入關鍵字，學會用篩選條件縮小範圍。
③ 看模型頁：重點看四個資訊——列印時間、耗材克數、列印設定（Print Profile）、授權。
④ 在 Bambu Studio 開啟：用建議設定直接開啟。
⑤ 認識三個參數：層高（越小越精細也越久）、填充（內部密度）、支撐（懸空的地方要不要撐）。
【互動】示範到③時問：「這個模型要印多久？用幾克？」請學員從畫面上找答案。
""")
    finish_slide(s, an)

    # ---------------- 18 授權小提醒
    n = page()
    s, an = new_slide(prs, "mw", "授權小提醒")
    ok = box(s, 0.6, 1.5, 5.95, 2.1, fill=CARD, radius=0.18,
             text=[[("✓ 自己印來用", {"size": 30, "bold": True})], [("個人列印通常沒問題", {"size": 22, "color": MUTED})]])
    ng = box(s, 6.75, 1.5, 5.95, 2.1, fill=ORANGE, radius=0.18,
             text=[[("! 要拿去賣", {"size": 30, "bold": True})], [("必須確認授權允許商用", {"size": 22})]],
             color=DARK)
    an.wipe(ok, 400, 500, "bottom")
    an.wipe(ng, 900, 500, "bottom")
    an.pulse(ng, 1500, 300, times=2, scale=1.04)
    tip = tb(s, 0.6, 4.0, 6.0, 2.2,
             [[("怎麼看？", {"bold": True, "color": ORANGE, "size": 24})],
              [("模型頁會標示授權類型", {"size": 22})],
              [("看到 NC＝不可商用", {"size": 22})]], margin=0, space_after=8)
    an.wipe(tip, 1900, 600, "left")
    ph = placeholder(s, n, 6.75, 3.95, 5.95, 2.2, "截圖：模型頁的授權標示位置")
    an.wipe(ph, 2100, 600, "bottom")
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 1 分鐘（講者 D）
【這頁要講】MakerWorld 上的模型都是別人的作品，各有授權。
自己印來用、學習用，通常沒有問題；但如果要拿去賣，一定要先確認授權允許商用。
常見的 Creative Commons 授權裡，有「NC（NonCommercial）」字樣就是不可商用。
【互動】在示範頁面上指出授權標示的位置，提醒大家：等一下選送印候選時，「授權允許」是條件之一。
""")
    finish_slide(s, an)

    # ---------------- 19 小組任務說明（40 分鐘）
    n = page()
    s, an = new_slide(prs, "mw", "小組任務說明")
    badge = box(s, 9.6, 1.45, 3.1, 0.95, fill=ORANGE, radius=0.16, text="40 分鐘", size=36, bold=True,
                color=DARK)
    an.wipe(badge, 300, 400, "left")
    a1 = box(s, 0.6, 1.45, 8.8, 0.95, fill=CARD, radius=0.14)
    a1t = tb(s, 0.85, 1.45, 8.4, 0.95, [[("① 尋寶任務　", {"bold": True, "color": ORANGE}), ("完成 Part 5 表格", {})]],
             size=24, anchor=MSO_ANCHOR.MIDDLE, margin=0)
    an.wipe([a1, a1t], 600, 500, "left")
    a2 = box(s, 0.6, 2.6, 12.1, 2.05, fill=ORANGE_BG, radius=0.16)
    a2t = tb(s, 0.85, 2.68, 11.6, 0.6, [[("② 選出一個送印候選　", {"bold": True, "color": ORANGE}),
                                         ("（Part 6）條件：", {})]], size=24, margin=0, anchor=MSO_ANCHOR.MIDDLE)
    an.wipe([a2, a2t], 1000, 500, "bottom")
    conds = ["≤ 60 分鐘", "≤ 20 g", "免支撐為佳", "授權允許"]
    for j, cnd in enumerate(conds):
        c = box(s, 0.85 + j * 2.95, 3.4, 2.75, 1.05, fill=ORANGE, radius=0.14, text=cnd, size=30, bold=True,
                color=DARK)
        an.wipe(c, 1500 + j * 280, 350, "bottom")
    a3 = tb(s, 0.6, 4.8, 12.0, 0.6, [[("③ 加分題　", {"bold": True, "color": ORANGE}),
                                     ("找需要支撐的模型，說明支撐加在哪", {})]], size=22, margin=0,
            anchor=MSO_ANCHOR.MIDDLE)
    an.wipe(a3, 2800, 500, "left")
    progress_bar(s, an, 3.6, 6.0, 7.9, 40, [0, 10, 20, 30, 40], 3300, "")
    tb(s, 3.6, 5.5, 7.9, 0.45, "⏱ 40 分鐘倒數", size=20, color=MUTED, margin=0, anchor=MSO_ANCHOR.BOTTOM)
    worksheet_tag(s, an, "5・6", 3100)
    mini_printer(s, an, 200)
    notes(s, """
【時間】這頁會停留約 40 分鐘（講者 D 說明 1 分鐘，其餘由各組工作人員陪同）
【這頁要講】接下來 40 分鐘是小組任務：
① 尋寶任務：照學習單 Part 5 的表格，在 MakerWorld 找到指定的資訊。
② 選出一個送印候選，寫在 Part 6，條件有四個：列印時間 ≤ 60 分鐘、耗材 ≤ 20 g、最好免支撐、授權允許。
③ 提早完成的組做加分題：找一個需要支撐的模型，說明支撐會加在哪裡、為什麼。
【操作】下方進度條會在 40 分鐘內跑滿（每 10 分鐘一格）。B 在控台留意，剩 10 分鐘、剩 5 分鐘時口頭提醒。
【互動】工作人員巡組時可以問：「你們的候選為什麼符合這四個條件？」——這也是等一下發表的內容。
""")
    finish_slide(s, an)

    # ---------------- 20 發表規則
    n = page()
    s, an = new_slide(prs, "share", "發表規則")
    lead = tb(s, 0.6, 1.4, 12.0, 0.6, [[("每組 3 分鐘", {"bold": True, "color": ORANGE}), ("，說清楚這 4 件事：", {})]],
              size=26, margin=0)
    an.wipe(lead, 400, 600, "left")
    pts = ["① 選了什麼", "② 為什麼選", "③ 預估時間與耗材", "④ 需不需要支撐"]
    for j, p_ in enumerate(pts):
        c = box(s, 0.6 + j * 3.07, 2.2, 2.9, 1.2, fill=CARD, radius=0.14, text=p_, size=22, bold=True)
        an.wipe(c, 1000 + j * 280, 400, "bottom")
    ol = tb(s, 0.6, 3.85, 6.0, 0.55, "發表順序", size=24, bold=True, color=ORANGE, margin=0)
    an.fade(ol, 2200, 300)
    t = 2400
    for j in range(5):
        x = 0.6 + j * 2.5
        c = box(s, x, 4.5, 2.0, 0.95, fill=ORANGE if j == 0 else CARD2, radius=0.14,
                text=f"第 {j + 1} 組", size=24, bold=True, color=DARK if j == 0 else WHITE)
        an.wipe(c, t, 300, "left")
        t += 300
        if j < 4:
            ar = arrow(s, x + 2.07, 4.8, 0.36, 0.34, color=GRAY)
            an.wipe(ar, t, 150, "left")
            t += 120
    tm = tb(s, 0.6, 5.7, 11.0, 0.5, "A 主持・B 計時，時間到會舉牌提醒", size=20, color=MUTED, margin=0)
    an.fade(tm, t, 300)
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 16 分鐘（A 主持、B 計時）：5 組 × 3 分鐘＋換場
【這頁要講】每組 3 分鐘，請說清楚：選了什麼模型、為什麼選、預估列印時間與耗材克數、需不需要支撐。
發表順序預設第 1 組到第 5 組（如要調整，請修改這頁下方的組別順序）。
【操作】B 計時：剩 30 秒舉牌、時間到舉牌。
【互動】每組發表後，A 可以問台下一個問題，例如：「如果這個模型要拿去賣，可以嗎？」或「這個模型印的時候要注意什麼？」
""")
    finish_slide(s, an)

    # ---------------- 21 今天的成果
    n = page()
    s, an = new_slide(prs, "end", "今天的成果")
    ph = placeholder(s, n, 0.6, 1.5, 7.2, 4.85, "照片：今日 6 台示範列印成品", size=20)
    an.wipe(ph, 400, 900, "bottom")
    rh = tb(s, 8.3, 1.5, 4.4, 0.55, "回顧四步驟", size=24, bold=True, color=ORANGE, margin=0)
    an.fade(rh, 1000, 300)
    t = 1200
    for j, st_ in enumerate(["建模", "切片", "列印", "後處理"]):
        y = 2.15 + j * 0.95
        c = box(s, 8.3, y, 4.4, 0.7, fill=CARD, radius=0.35, text=f"{j + 1}　{st_}", size=24, bold=True)
        an.wipe(c, t, 350, "bottom")
        t += 300
    done = tb(s, 8.3, 5.95, 4.4, 0.5, "今天你都走過一遍了！", size=22, color=ORANGE, bold=True, margin=0)
    an.wipe(done, t, 500, "left")
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 1 分鐘（講者 A）
【這頁要講】開場時 C 啟動的 6 台示範列印已經完成了！（請 C 把成品拿到台前，或在照片框放上成品照）
回顧四步驟：建模（今天用 MakerWorld 別人做好的）→ 切片（Bambu Studio）→ 列印（剛剛在列印站看到的）→ 後處理（拆支撐）。
【互動】把成品傳下去給各組摸一摸，問：「跟你早上摸的樣品比，層紋看得出來嗎？」
""")
    finish_slide(s, an)

    # ---------------- 22 課後想自己印？
    n = page()
    s, an = new_slide(prs, "end", "課後想自己印？")
    lead = tb(s, 0.6, 1.4, 11.0, 0.6, "列印教室使用流程", size=24, color=MUTED, margin=0)
    an.fade(lead, 400, 300)
    flow = ["線上預約", "老師審核", ["到場掃", "QR Code 報到"], ["結束時登記", "並釋放設備"]]
    t = 700
    for j, f_ in enumerate(flow):
        x = 0.6 + j * 2.95
        shp = MSO_SHAPE.PENTAGON if j == 0 else MSO_SHAPE.CHEVRON
        c = s.shapes.add_shape(shp, Inches(x), Inches(2.2), Inches(3.2), Inches(1.75))
        c.shadow.inherit = False
        c.fill.solid(); c.fill.fore_color.rgb = ORANGE if j % 2 == 0 else CARD2
        c.line.fill.background()
        c.adjustments[0] = 0.22
        fill_text(c.text_frame, f_, size=20, bold=True, color=DARK if j % 2 == 0 else WHITE,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.0)
        c.text_frame.margin_left = Inches(0.3 if j else 0.1)
        c.text_frame.margin_right = Inches(0.15)
        lab = tb(s, x + 0.1, 4.05, 2.8, 0.45, f"STEP {j + 1}", size=18, bold=True, color=GRAY,
                 align=PP_ALIGN.CENTER, margin=0)
        an.wipe([c, lab], t, 450, "left")  # 擠出：流程一段段被擠出來
        t += 400
    warn = box(s, 0.6, 4.85, 5.6, 0.85, fill=None, line=ORANGE, line_w=2.25, radius=0.14,
               text="⚠ 線材需要自備", size=26, bold=True, color=ORANGE)
    an.wipe(warn, t, 450, "left")
    an.pulse(warn, t + 500, 300, times=2, scale=1.05)
    link = tb(s, 6.6, 4.85, 6.1, 0.85, "預約連結：＿＿＿＿＿＿＿＿", size=20, color=MUTED, anchor=MSO_ANCHOR.MIDDLE,
              margin=0)
    an.fade(link, t + 300, 300)
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 1 分鐘（講者 A）
【這頁要講】課後想自己印，可以使用列印教室，流程四步：
1. 線上預約時段與機台；2. 等老師審核通過；3. 當天到場掃 QR Code 報到；4. 結束時登記並釋放設備，讓下一位可以使用。
特別提醒：線材需要自備。
（預約連結請自行填在投影片上，或另外貼 QR Code）
【互動】問：「今天選的送印候選，有沒有人想課後自己印？」鼓勵學員把 Part 6 的模型帶回去印。
""")
    finish_slide(s, an)

    # ---------------- 23 反思與問卷
    n = page()
    s, an = new_slide(prs, "end", "反思與問卷")
    l1 = tb(s, 0.6, 1.6, 7.2, 0.8, [[("請完成學習單 ", {}), ("Part 7", {"color": ORANGE})]], size=32, bold=True,
            margin=0, anchor=MSO_ANCHOR.MIDDLE)
    an.wipe(l1, 400, 600, "left")
    prompts = ["今天最有收穫的一件事", "還想學什麼、有什麼疑問"]
    for j, p_ in enumerate(prompts):
        c = box(s, 0.6, 2.7 + j * 1.1, 7.0, 0.9, fill=CARD, radius=0.14)
        tx = tb(s, 0.9, 2.7 + j * 1.1, 6.5, 0.9, p_, size=24, anchor=MSO_ANCHOR.MIDDLE, margin=0)
        an.wipe([c, tx], 1000 + j * 350, 450, "bottom")
    l2 = tb(s, 0.6, 5.1, 7.2, 0.6, "寫完請掃右邊 QR Code 填問卷 →", size=24, bold=True, color=ORANGE, margin=0)
    an.wipe(l2, 1800, 600, "left")
    ph = placeholder(s, n, 8.6, 1.6, 3.8, 3.8, "QR Code：課後問卷", size=20)
    an.wipe(ph, 1200, 900, "bottom")
    worksheet_tag(s, an, 7, 1800)
    mini_printer(s, an, 200)
    notes(s, """
【時間】約 4 分鐘（講者 A，E 協助確認問卷回收）
【這頁要講】請大家完成學習單 Part 7 的反思：今天最有收穫的一件事、還想學什麼或有什麼疑問。
寫完後掃描 QR Code 填課後問卷，大約 2 分鐘。
【互動】請各組工作人員確認組員都有掃到 QR Code；E 留意填答人數。
""")
    finish_slide(s, an)

    # ---------------- 24 謝謝參與＋合照
    n = page()
    s, an = new_slide(prs, "end")
    # 結尾印一座獎盃：底座 → 杯腳 → 杯身
    print_object(s, an, 9.9, 6.35, [2.3, 1.5, 0.6, 0.6, 0.6, 1.3, 2.3, 2.9, 3.2, 3.3, 3.3],
                 0.34, start=200, per_layer=300)
    t1 = tb(s, 0.6, 2.2, 6.5, 1.0, "謝謝參與！", size=40, bold=True)
    t2 = tb(s, 0.6, 3.25, 6.5, 0.8, "來拍張大合照 📸", size=32, bold=True, color=ORANGE)
    t3 = tb(s, 0.6, 4.3, 6.3, 1.0, ["請各組帶著學習單與成品，", "到前方集合"], size=22, color=MUTED)
    an.wipe(t1, 400, 700, "left")
    an.wipe(t2, 1100, 700, "left")
    an.fade(t3, 1800, 500)
    notes(s, """
【時間】約 2 分鐘（講者 A，E 負責拍照）
【這頁要講】謝謝大家今天的參與！從概念、產業應用、看機台，到自己在 MakerWorld 挑出能送印的模型，大家都完成了。
【互動】請各組帶著學習單與示範成品到前方合照；E 拍照。提醒學員帶走個人物品、學習單依工作人員指示繳回或帶走。
【動畫】右邊的噴嘴會印出一座獎盃，可以等它印完再拍照當背景。
""")
    finish_slide(s, an, "fade")

    prs.save(out)
    return pg[0]


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--font", default="微軟正黑體", help="字型（找不到微軟正黑體時可改用 Noto Sans TC）")
    ap.add_argument("--out", default="3D列印工作坊.pptx")
    a = ap.parse_args()
    FONT = a.font
    here = os.path.dirname(os.path.abspath(__file__))
    out = a.out if os.path.isabs(a.out) else os.path.join(here, a.out)
    total = build(out)
    print(f"已輸出 {out}，共 {total} 頁")
    print("圖片預留框：")
    for p_, lab in PLACEHOLDERS:
        print(f"  第 {p_} 頁：{lab.replace(chr(10), '')}")
