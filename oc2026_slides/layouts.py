# -*- coding: utf-8 -*-
"""レイアウト実装。1枚に載せる情報量を構造で縛り、文字を大きく保つ。"""
from deckkit import *
from deckkit import MIN_CARD, MIN_BODY, T, C, M, SLIDE_W, SLIDE_H, CONTENT_W
from deckkit import fn_height, WARNINGS, CUR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def _split(s, n=2):
    for d in ("｜", "|"):
        if d in s:
            parts = [x.strip() for x in s.split(d)]
            break
    else:
        parts = [s.strip()]
    parts += [""] * (n - len(parts))
    return parts[:n]

# ------------------------------------------------------------------ 表紙
def L_title(slide, s):
    bg(slide, C["navy"])
    px = 8.95
    rect(slide, px, 0, SLIDE_W - px, SLIDE_H, fill="16375F")
    rect(slide, px, 0, 0.12, SLIDE_H, fill=C["gold"])
    pw = SLIDE_W - px - 0.75
    if s.get("photo_slot"):
        textbox(slide, px + 0.40, 0.62, pw, 0.34, "［写真を差し込む位置］", 18, True, "C9DCF0",
                nowrap=True)
        hh = est_lines(s["photo_slot"], pw, 16) * (16 / 72.0) * 1.35
        textbox(slide, px + 0.40, 1.06, pw, hh + 0.08, s["photo_slot"], 16, False, "C9DCF0",
                line_spacing=1.35)
    if s.get("photo_caption"):
        textbox(slide, px + 0.40, SLIDE_H - 1.75, pw, 1.30, s["photo_caption"], 20, True,
                "E2ECF7", line_spacing=1.35, anchor=MSO_ANCHOR.BOTTOM)
    tw = px - M["l"] - 0.6
    ksz = 20
    while ksz > 15 and text_w_in(s.get("kicker", ""), ksz) > tw:
        ksz -= 1
    textbox(slide, M["l"], 1.00, tw, 0.40, s.get("kicker", ""), ksz, True, C["gold"],
            spc=2.0, nowrap=True)
    nl = len(s["title"].split("\n"))
    tsz = 54
    while tsz > 34 and est_lines(s["title"], tw, tsz) > nl:
        tsz -= 2
    th = est_lines(s["title"], tw, tsz) * (tsz / 72.0) * 1.26
    textbox(slide, M["l"], 1.70, tw, th + 0.10, s["title"], tsz, True, C["white"],
            line_spacing=1.26)
    y = 1.70 + th + 0.46
    rect(slide, M["l"], y, 1.9, 0.09, fill=C["gold"])
    y += 0.46
    if s.get("lead"):
        lh = est_lines(s["lead"], tw, 30) * (30 / 72.0) * 1.35
        textbox(slide, M["l"], y, tw, lh + 0.08, s["lead"], 30, False, "DCE8F5",
                line_spacing=1.35)
    pts = s.get("body_points", [])[:2]
    by = SLIDE_H - 0.62 - len(pts) * 0.46
    for i, p in enumerate(pts):
        textbox(slide, M["l"], by + i * 0.46, tw, 0.44, p, 20, False, "BBD0E6", nowrap=True)

# ------------------------------------------------------------------ 区切り
def L_section(slide, s):
    tone = s.get("tone", "navy")
    base = {"navy": C["navy"], "teal": "0B5F58", "brand": "0C4C9E"}.get(tone, C["navy"])
    bg(slide, base)
    if s.get("big_number"):
        textbox(slide, M["l"] - 0.12, 0.72, 3.2, 2.6, s["big_number"], 150, True,
                "3A6699", line_spacing=1.0, nowrap=True)
    tsz = 50
    while tsz > 34 and est_lines(s["title"], 11.4, tsz) > 1:
        tsz -= 2
    th = est_lines(s["title"], 11.4, tsz) * (tsz / 72.0) * 1.28
    pts = s.get("body_points", [])[:3]
    y = SLIDE_H - 1.05 - len(pts) * 0.60
    ty = y - 0.56 - th
    rect(slide, M["l"], ty - 0.44, 1.9, 0.09, fill=C["gold"])
    textbox(slide, M["l"], ty, 11.4, th + 0.08, s["title"], tsz, True, C["white"],
            line_spacing=1.28)
    for p in pts:
        rect(slide, M["l"] + 0.03, y + 0.16, 0.17, 0.17, fill=C["gold"])
        textbox(slide, M["l"] + 0.50, y, 10.4, 0.50, p, 26, False, "CFE0F2")
        y += 0.60

# ------------------------------------------------------------------ アジェンダ
def L_agenda(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    if s.get("lead"):
        textbox(slide, M["l"], y, CONTENT_W, 0.5, s["lead"], T["lead"], False, C["gray"])
        y += 0.78
    rows = s.get("body_points", [])[:3]
    eng_h = 1.02 if s.get("engagement") else 0.0
    gap = 0.26
    avail = SLIDE_H - M["b"] - y - eng_h
    h = max(0.86, min(1.34, (avail - gap * (len(rows) - 1)) / max(len(rows), 1)))
    for i, raw in enumerate(rows):
        star = raw.startswith("*")
        txt = raw[1:].strip() if star else raw
        fill = C["gold_lt"] if star else C["bg"]
        rect(slide, M["l"], y, CONTENT_W, h, fill=fill)
        rect(slide, M["l"], y, 0.11, h, fill=C["gold"] if star else C["brand_lt"])
        d = min(0.72, h - 0.34)
        rect(slide, M["l"] + 0.42, y + (h - d) / 2, d, d,
             fill=C["gold"] if star else C["brand"], shape=MSO_SHAPE.OVAL)
        textbox(slide, M["l"] + 0.42, y + (h - d) / 2 + d * 0.15, d, d * 0.74, str(i + 1),
                max(22, min(32, int(d * 44))), True,
                C["white"] if not star else C["gold_tx"],
                align=PP_ALIGN.CENTER, nowrap=True)
        size = 28
        while size > 22 and est_lines(txt, CONTENT_W - 2.0, size) > 1:
            size -= 1
        textbox(slide, M["l"] + 1.52, y + 0.16, CONTENT_W - 2.02, h - 0.32, txt, size,
                star, C["ink"], line_spacing=1.3, anchor=MSO_ANCHOR.MIDDLE, tag="bullet")
        y += h + gap
    engage_strip(slide, s.get("engagement", ""))

# ------------------------------------------------------------------ 箇条書き
def L_bullets(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    has_photo = bool(s.get("photo_slot"))
    w = (CONTENT_W - 4.35) if has_photo else CONTENT_W
    if s.get("lead"):
        n = est_lines(s["lead"], w, T["lead"])
        hh = n * (T["lead"] / 72.0) * 1.35
        textbox(slide, M["l"], y, w, hh + 0.05, s["lead"], T["lead"], True, C["brand"],
                line_spacing=1.35)
        y += hh + 0.42
    fn_h = fn_height(s.get("footnote", ""))
    eng_h = 1.05 if s.get("engagement") else 0.0
    if has_photo:
        ptop = 2.05
        pbot = SLIDE_H - M["b"] - fn_h - eng_h - 0.10
        photo_slot(slide, SLIDE_W - M["r"] - 4.0, ptop, 4.0, max(2.4, pbot - ptop),
                   s.get("photo_caption", ""), s["photo_slot"])
    avail = SLIDE_H - M["b"] - y - fn_h - eng_h
    bullet_block(slide, M["l"], y, w, s.get("body_points", []), max_h=avail)
    engage_strip(slide, s.get("engagement", ""),
                 y=SLIDE_H - M["b"] - fn_h - 0.90 if s.get("engagement") else None)
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ 3枚カード
def L_three_cards(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    if s.get("lead"):
        n = est_lines(s["lead"], CONTENT_W, T["lead"])
        hh = n * (T["lead"] / 72.0) * 1.35
        textbox(slide, M["l"], y, CONTENT_W, hh + 0.05, s["lead"], T["lead"], False,
                C["gray"], line_spacing=1.35)
        y += hh + 0.40
    cards = s.get("body_points", [])[:3]
    gap = 0.30
    w = (CONTENT_W - gap * (len(cards) - 1)) / max(len(cards), 1)
    h = min(4.20, SLIDE_H - M["b"] - y - fn_height(s.get("footnote", "")))
    tones = [(C["brand"], C["brand_lt"]), (C["teal"], C["teal_lt"]), (C["gold"], C["gold_lt"])]
    for i, raw in enumerate(cards):
        t, body = _split(raw, 2)
        x = M["l"] + i * (w + gap)
        fg, bgc = tones[i % 3]
        rect(slide, x, y, w, h, fill=bgc)
        rect(slide, x, y, w, 0.10, fill=fg)
        textbox(slide, x + 0.34, y + 0.36, 1.2, 0.46, f"0{i+1}", 26, True,
                C["gold_tx"] if i % 3 == 2 else fg, spc=1.2, nowrap=True)
        ts = 30
        while ts > 22 and est_lines(t, w - 0.68, ts) > 1:
            ts -= 1
        th = est_lines(t, w - 0.68, ts) * (ts / 72.0) * 1.25
        textbox(slide, x + 0.34, y + 1.00, w - 0.68, th + 0.05, t, ts, True, C["ink"],
                line_spacing=1.25)
        by = y + 1.00 + th + 0.32
        bs = 26
        while bs > 20 and est_lines(body, w - 0.68, bs) * (bs / 72.0) * 1.42 > h - (by - y) - 0.34:
            bs -= 1
        textbox(slide, x + 0.34, by, w - 0.68, h - (by - y) - 0.30, body, bs, False,
                C["ink"], line_spacing=1.42)
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ 4段の階段
def L_four_steps(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    if s.get("lead"):
        textbox(slide, M["l"], y, CONTENT_W, 0.52, s["lead"], 26, True, C["brand_dk"])
        y += 0.76
    steps = s.get("body_points", [])[:4]
    gap = 0.26
    w = (CONTENT_W - gap * 3) / 4
    bottom = SLIDE_H - M["b"] - fn_height(s.get("footnote", "")) - 0.06
    rise = min(0.36, max(0.0, (bottom - y - 2.92) / 3.0))
    tops = [y + (3 - i) * rise for i in range(4)]
    tone = [("6E9BCE", C["bg"]), ("2E76C0", C["brand_lt"]), (C["brand"], C["brand_lt"]),
            (C["gold"], C["gold_lt"])]
    for i, raw in enumerate(steps):
        year, title, money = (_split(raw, 3) + [""] * 3)[:3] if raw.count("｜") >= 2 else (
            _split(raw, 2) + [""])[:3]
        x = M["l"] + i * (w + gap)
        top = tops[i]
        fg, bgc = tone[i]
        rect(slide, x, top, w, bottom - top, fill=bgc)
        rect(slide, x, top, w, 0.74, fill=fg)
        textbox(slide, x, top + 0.16, w, 0.50, year, 26, True,
                C["white"] if i != 3 else C["gold_tx"], align=PP_ALIGN.CENTER, nowrap=True)
        ts = 24
        while ts > MIN_CARD and est_lines(title, w - 0.44, ts) > 2:
            ts -= 1
        th = est_lines(title, w - 0.44, ts) * (ts / 72.0) * 1.3
        textbox(slide, x + 0.22, top + 1.02, w - 0.44, th + 0.06, title, ts, True, C["ink"],
                line_spacing=1.3)
        if money:
            pill(slide, x + 0.18, bottom - 0.80, w - 0.36, 0.62, money, 26,
                 C["coral"] if i else C["teal"], min_size=MIN_CARD)
            if top + 1.02 + th > bottom - 0.92:
                WARNINGS.append(f"p{CUR[0]:02d} [階段の段が窮屈] {year} の説明文を短くすること")
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ 巨大数字
def L_big_number(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    num = s.get("big_number", "")
    pw = 5.25
    fn_h = fn_height(s.get("footnote", ""))
    eng_h = 1.05 if s.get("engagement") else 0.0
    ph = min(3.60, SLIDE_H - M["b"] - y - fn_h - eng_h - 0.05)
    lab = s.get("big_number_label", "")
    ls_ = 26
    while ls_ > 18 and est_lines(lab, pw - 0.6, ls_) > 2:
        ls_ -= 1
    lh = est_lines(lab, pw - 0.6, ls_) * (ls_ / 72.0) * 1.3
    size = T["huge"]
    while size > 54 and (text_w_in(num, size) > pw - 0.5
                         or size / 72.0 * 1.16 > ph - lh - 0.70):
        size -= 4
    rect(slide, M["l"], y + 0.05, pw, ph, fill=C["brand_lt"])
    nh = size / 72.0 * 1.16
    ntop = y + 0.05 + max(0.20, (ph - lh - 0.34 - nh) / 2)
    textbox(slide, M["l"], ntop, pw, nh + 0.1, num, size, True, C["brand"],
            align=PP_ALIGN.CENTER, line_spacing=1.16, nowrap=True)
    textbox(slide, M["l"] + 0.3, y + 0.05 + ph - lh - 0.26, pw - 0.6, lh + 0.08, lab, ls_,
            True, C["ink"], align=PP_ALIGN.CENTER, line_spacing=1.3)
    bx = M["l"] + pw + 0.45
    bullet_block(slide, bx, y + 0.14, CONTENT_W - pw - 0.45, s.get("body_points", []),
                 size=T["body_s"], gap=0.34, marker=C["brand"], max_h=ph - 0.18)
    engage_strip(slide, s.get("engagement", ""))
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ クイズ
def L_quiz(slide, s):
    bg(slide, C["bg"])
    rect(slide, 0, 0, SLIDE_W, 0.16, fill=C["gold"])
    if s.get("kicker"):
        textbox(slide, M["l"], 0.68, CONTENT_W, 0.42, s["kicker"].upper(), 20, True,
                C["gold_tx"], spc=2.4, align=PP_ALIGN.CENTER)
    ts = 44 if text_w_in(s["title"], 44) <= CONTENT_W else 38
    th = est_lines(s["title"], CONTENT_W, ts) * (ts / 72.0) * 1.28
    textbox(slide, M["l"], 1.30, CONTENT_W, th + 0.06, s["title"], ts, True, C["ink"],
            align=PP_ALIGN.CENTER, line_spacing=1.28)
    y = 1.30 + th + 0.55
    ch = s.get("body_points", [])[:3]
    gap = 0.28
    w = (CONTENT_W - gap * (len(ch) - 1)) / max(len(ch), 1)
    h = 2.45
    for i, raw in enumerate(ch):
        x = M["l"] + i * (w + gap)
        rect(slide, x, y, w, h, fill=C["white"], line=C["line"], line_w=1.4)
        rect(slide, x + (w - 0.72) / 2, y + 0.28, 0.72, 0.72, fill=C["brand"],
             shape=MSO_SHAPE.OVAL)
        textbox(slide, x + (w - 0.72) / 2, y + 0.38, 0.72, 0.50, "ABC"[i], 28, True,
                C["white"], align=PP_ALIGN.CENTER, nowrap=True)
        cs = 28
        while cs > 22 and est_lines(raw, w - 0.5, cs) > 2:
            cs -= 1
        textbox(slide, x + 0.25, y + 1.20, w - 0.5, 1.05, raw, cs, True, C["ink"],
                align=PP_ALIGN.CENTER, line_spacing=1.3)
        y2 = y
    engage_strip(slide, s.get("engagement", ""), y=y + h + 0.30)

# ------------------------------------------------------------------ 棒グラフ
def L_stat_bar(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    if s.get("lead"):
        textbox(slide, M["l"], y, CONTENT_W, 0.54, s["lead"], 26, True, C["brand_dk"])
        y += 0.74
    rows = []
    for raw in s.get("body_points", []):
        hi = raw.startswith("*")
        lab, val = _split(raw[1:] if hi else raw, 2)
        try:
            v = float(val.replace("%", ""))
        except ValueError:
            v = 0.0
        rows.append((lab, val, v, hi))
    vmax = max([r[2] for r in rows] or [1])
    fn_h = fn_height(s.get("footnote", ""))
    avail = SLIDE_H - M["b"] - y - fn_h - 0.10
    n = max(len(rows), 1)
    bh = min(0.86, max(0.50, (avail - 0.34 * (n - 1)) / n))
    gapy = min(0.36, max(0.18, (avail - bh * n) / max(n - 1, 1)))
    lab_w = 3.45
    val_w = 2.05
    bar_x = M["l"] + lab_w + 0.28
    bar_max = CONTENT_W - lab_w - 0.28 - val_w
    for lab, val, v, hi in rows:
        lsz = 26
        while lsz > 20 and est_lines(lab, lab_w, lsz) > 1:
            lsz -= 1
        textbox(slide, M["l"], y + (bh - lsz / 72.0 * 1.35) / 2, lab_w, lsz / 72.0 * 1.5 + 0.06,
                lab, lsz, hi, C["ink"] if hi else C["gray"], align=PP_ALIGN.RIGHT)
        wpx = max(bar_max * (v / vmax), 0.12)
        rect(slide, bar_x, y, wpx, bh, fill=C["coral"] if hi else "C9D4E0")
        vsz = 36 if hi else 28
        while vsz > 22 and text_w_in(val, vsz) > val_w - 0.2:
            vsz -= 2
        textbox(slide, bar_x + wpx + 0.18, y + (bh - vsz / 72.0 * 1.3) / 2, val_w,
                vsz / 72.0 * 1.45 + 0.06, val, vsz, True,
                C["coral"] if hi else C["gray"])
        y += bh + gapy
    footnote(slide, s.get("footnote", ""))


# ------------------------------------------------------------------ 写真主役
def L_photo_focus(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    pw = 6.55
    fn_h = fn_height(s.get("footnote", ""))
    photo_slot(slide, M["l"], y, pw, SLIDE_H - M["b"] - y - fn_h - 0.05,
               s.get("photo_caption", ""), s.get("photo_slot", ""))
    bx = M["l"] + pw + 0.5
    bullet_block(slide, bx, y + 0.08, CONTENT_W - pw - 0.5, s.get("body_points", []),
                 size=T["body_s"], gap=0.36, max_h=SLIDE_H - M["b"] - y - fn_h - 0.15)
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ 2カラム
def L_two_column(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    lw = 6.55
    if s.get("lead"):
        n = est_lines(s["lead"], lw, T["lead"])
        hh = n * (T["lead"] / 72.0) * 1.35
        textbox(slide, M["l"], y, lw, hh + 0.05, s["lead"], T["lead"], True, C["brand"],
                line_spacing=1.35)
        y2 = y + hh + 0.36
    else:
        y2 = y
    fn_h = fn_height(s.get("footnote", ""))
    bullet_block(slide, M["l"], y2, lw, s.get("body_points", []), size=T["body_s"], gap=0.34,
                 max_h=SLIDE_H - M["b"] - y2 - fn_h - 0.10)
    px = M["l"] + lw + 0.45
    pwid = CONTENT_W - lw - 0.45
    panel = s.get("panel", {})
    if panel:
        ph = SLIDE_H - M["b"] - y - fn_h - 0.10
        rect(slide, px, y, pwid, ph, fill=C["bg"])
        rect(slide, px, y, pwid, 0.09, fill=C["gold"])
        pts = 24
        while pts > 20 and est_lines(panel.get("title", ""), pwid - 0.64, pts) > 1:
            pts -= 1
        textbox(slide, px + 0.32, y + 0.40, pwid - 0.64, 0.54, panel.get("title", ""), pts,
                True, C["ink"])
        cy = y + 1.16
        for it in panel.get("items", [])[:4]:
            k, v = _split(it, 2)
            textbox(slide, px + 0.32, cy, pwid - 0.64, 0.38, k, 20, False, C["gray"])
            vs = 30
            while vs > 22 and est_lines(v, pwid - 0.64, vs) > 1:
                vs -= 1
            textbox(slide, px + 0.32, cy + 0.40, pwid - 0.64, 0.56, v, vs, True, C["coral"])
            cy += 1.16
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ まとめ
def L_summary(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    rows = s.get("body_points", [])[:3]
    eng_h = 1.34 if s.get("engagement") else 0.0
    fn_h = fn_height(s.get("footnote", ""))
    gap = 0.24
    avail = SLIDE_H - M["b"] - y - eng_h - fn_h
    h = max(0.92, min(1.55, (avail - gap * (len(rows) - 1)) / max(len(rows), 1)))
    for i, raw in enumerate(rows):
        rect(slide, M["l"], y, CONTENT_W, h, fill=C["bg"])
        rect(slide, M["l"], y, 0.11, h, fill=[C["brand"], C["teal"], C["gold"]][i % 3])
        nsz = max(30, min(44, int(h * 72 * 0.62)))
        nbh = nsz / 72.0 * 1.4
        textbox(slide, M["l"] + 0.42, y + (h - nbh) / 2, 0.9, nbh, str(i + 1), nsz, True,
                [C["brand_dk"], C["teal"], C["gold_tx"]][i % 3], align=PP_ALIGN.LEFT, nowrap=True)
        size = 28
        while size > 22 and est_lines(raw, CONTENT_W - 2.15, size) * (size / 72.0) * 1.34 > h - 0.3:
            size -= 1
        textbox(slide, M["l"] + 1.50, y + 0.16, CONTENT_W - 2.0, h - 0.32, raw, size, True,
                C["ink"], line_spacing=1.34, anchor=MSO_ANCHOR.MIDDLE, tag="bullet")
        y += h + gap
    engage_strip(slide, s.get("engagement", ""))
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ クロージング
def L_closing(slide, s):
    bg(slide, C["navy"])
    rect(slide, 0, 0, SLIDE_W, 0.16, fill=C["gold"])
    ts = 48 if text_w_in(s["title"], 48) <= 11.0 else 40
    th = est_lines(s["title"], 11.0, ts) * (ts / 72.0) * 1.3
    textbox(slide, 1.15, 1.75, 11.0, th + 0.1, s["title"], ts, True, C["white"],
            align=PP_ALIGN.CENTER, line_spacing=1.3)
    y = 1.75 + th + 0.5
    rect(slide, (SLIDE_W - 1.9) / 2, y, 1.9, 0.09, fill=C["gold"])
    y += 0.5
    for p in s.get("body_points", [])[:3]:
        textbox(slide, 1.15, y, 11.0, 0.56, p, 26, False, "CFE0F2", align=PP_ALIGN.CENTER)
        y += 0.66
    if s.get("lead"):
        textbox(slide, 1.15, SLIDE_H - 1.20, 11.0, 0.56, s["lead"], 24, True, C["gold"],
                align=PP_ALIGN.CENTER)

RENDER = {
    "title": L_title, "section": L_section, "agenda": L_agenda, "bullets": L_bullets,
    "three_cards": L_three_cards, "four_steps": L_four_steps, "big_number": L_big_number,
    "quiz": L_quiz, "stat_bar": L_stat_bar, "photo_focus": L_photo_focus,
    "two_column": L_two_column, "summary": L_summary, "closing": L_closing,
}
