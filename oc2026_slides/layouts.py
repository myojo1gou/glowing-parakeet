# -*- coding: utf-8 -*-
"""レイアウト実装。1枚に載せる情報量を構造で縛り、文字を大きく保つ。"""
from deckkit import *
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
    rect(slide, 7.75, 0, SLIDE_W - 7.75, SLIDE_H, fill="16375F")
    rect(slide, 7.75, 0, 0.12, SLIDE_H, fill=C["gold"])
    if s.get("photo_slot"):
        textbox(slide, 8.15, 0.55, 4.4, 0.3, "［写真を差し込む位置］", 13, False, "7E93AD")
        textbox(slide, 8.15, 0.92, 4.4, 1.2, s["photo_slot"], 14, False, "9FB3CB",
                line_spacing=1.35)
    if s.get("photo_caption"):
        textbox(slide, 8.15, SLIDE_H - 1.35, 4.4, 0.95, s["photo_caption"], 17, True,
                "D8E4F2", line_spacing=1.35, anchor=MSO_ANCHOR.BOTTOM)
    textbox(slide, M["l"], 1.15, 6.6, 0.34, s.get("kicker", ""), 17, True, C["gold"], spc=2.4)
    tsz = 52
    while tsz > 38 and est_lines(s["title"], 6.6, tsz) > 3:
        tsz -= 2
    th = est_lines(s["title"], 6.6, tsz) * (tsz / 72.0) * 1.24
    textbox(slide, M["l"], 1.80, 6.6, th + 0.08, s["title"], tsz, True, C["white"],
            line_spacing=1.24)
    y = 1.80 + th + 0.52
    rect(slide, M["l"], y, 1.9, 0.09, fill=C["gold"])
    y += 0.50
    if s.get("lead"):
        textbox(slide, M["l"], y, 6.6, 0.55, s["lead"], 25, False, "CFE0F2", line_spacing=1.35)
    for i, p in enumerate(s.get("body_points", [])[:2]):
        textbox(slide, M["l"], 6.18 + i * 0.46, 6.6, 0.44, p, 19, False, "9FB3CB")

# ------------------------------------------------------------------ 区切り
def L_section(slide, s):
    tone = s.get("tone", "navy")
    base = {"navy": C["navy"], "teal": "0B5F58", "brand": "0C4C9E"}.get(tone, C["navy"])
    bg(slide, base)
    if s.get("big_number"):
        textbox(slide, M["l"] - 0.12, 0.75, 3.2, 2.6, s["big_number"], 150, True,
                "2C5688", line_spacing=1.0)
    rect(slide, M["l"], 3.35, 1.9, 0.09, fill=C["gold"])
    textbox(slide, M["l"], 3.75, 10.5, 1.5, s["title"], 50, True, C["white"], line_spacing=1.25)
    y = 5.45
    for p in s.get("body_points", [])[:3]:
        rect(slide, M["l"] + 0.03, y + 0.12, 0.15, 0.15, fill=C["gold"])
        textbox(slide, M["l"] + 0.45, y, 10.0, 0.42, p, 22, False, "BFD3E8")
        y += 0.52

# ------------------------------------------------------------------ アジェンダ
def L_agenda(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    if s.get("lead"):
        textbox(slide, M["l"], y, CONTENT_W, 0.5, s["lead"], T["lead"], False, C["gray"])
        y += 0.78
    rows = s.get("body_points", [])[:3]
    eng_h = 1.02 if s.get("engagement") else 0.0
    gap = 0.24
    avail = SLIDE_H - M["b"] - y - eng_h
    h = min(1.16, (avail - gap * (len(rows) - 1)) / max(len(rows), 1))
    for i, raw in enumerate(rows):
        star = raw.startswith("*")
        txt = raw[1:].strip() if star else raw
        fill = C["gold_lt"] if star else C["bg"]
        rect(slide, M["l"], y, CONTENT_W, h, fill=fill)
        rect(slide, M["l"], y, 0.11, h, fill=C["gold"] if star else C["brand_lt"])
        d = min(0.72, h - 0.34)
        rect(slide, M["l"] + 0.42, y + (h - d) / 2, d, d,
             fill=C["gold"] if star else C["brand"], shape=MSO_SHAPE.OVAL)
        textbox(slide, M["l"] + 0.42, y + (h - d) / 2 + d * 0.16, d, d * 0.72, str(i + 1),
                int(d * 42), True, C["white"] if not star else "5C4204",
                align=PP_ALIGN.CENTER, nowrap=True)
        size = 25 if text_w_in(txt, 25) <= CONTENT_W - 2.0 else 22
        textbox(slide, M["l"] + 1.45, y + 0.16, CONTENT_W - 1.95, h - 0.32, txt, size,
                star, C["ink"], line_spacing=1.3, anchor=MSO_ANCHOR.MIDDLE)
        y += h + gap
    engage_strip(slide, s.get("engagement", ""))

# ------------------------------------------------------------------ 箇条書き
def L_bullets(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    has_photo = bool(s.get("photo_slot"))
    w = (CONTENT_W - 4.9) if has_photo else CONTENT_W
    if s.get("lead"):
        n = est_lines(s["lead"], w, T["lead"])
        hh = n * (T["lead"] / 72.0) * 1.35
        textbox(slide, M["l"], y, w, hh + 0.05, s["lead"], T["lead"], True, C["brand"],
                line_spacing=1.35)
        y += hh + 0.42
    if has_photo:
        photo_slot(slide, SLIDE_W - M["r"] - 4.55, 1.95, 4.55, 3.6,
                   s.get("photo_caption", ""), s["photo_slot"])
    fn_h = 0.0
    if s.get("footnote"):
        fn_h = est_lines(s["footnote"], CONTENT_W, T["small"]) * (T["small"] / 72.0) * 1.3 + 0.25
    eng_h = 1.05 if s.get("engagement") else 0.0
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
    h = min(4.20, SLIDE_H - M["b"] - y - (0.58 if s.get("footnote") else 0.0))
    tones = [(C["brand"], C["brand_lt"]), (C["teal"], C["teal_lt"]), (C["gold"], C["gold_lt"])]
    for i, raw in enumerate(cards):
        t, body = _split(raw, 2)
        x = M["l"] + i * (w + gap)
        fg, bgc = tones[i % 3]
        rect(slide, x, y, w, h, fill=bgc)
        rect(slide, x, y, w, 0.10, fill=fg)
        textbox(slide, x + 0.34, y + 0.34, 1.2, 0.42, f"0{i+1}", 24, True, fg, spc=1.2)
        ts = 26
        while ts > 20 and est_lines(t, w - 0.68, ts) > 1:
            ts -= 1
        th = est_lines(t, w - 0.68, ts) * (ts / 72.0) * 1.25
        textbox(slide, x + 0.34, y + 0.90, w - 0.68, th + 0.05, t, ts, True, C["ink"],
                line_spacing=1.25)
        by = y + 0.90 + th + 0.30
        bs = 21
        while bs > 18 and est_lines(body, w - 0.68, bs) * (bs / 72.0) * 1.42 > h - (by - y) - 0.34:
            bs -= 1
        textbox(slide, x + 0.34, by, w - 0.68, h - (by - y) - 0.30, body, bs, False,
                C["ink"], line_spacing=1.42)
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ 4段の階段
def L_four_steps(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    if s.get("lead"):
        textbox(slide, M["l"], y, CONTENT_W, 0.46, s["lead"], 22, True, C["brand"])
        y += 0.68
    steps = s.get("body_points", [])[:4]
    gap = 0.26
    w = (CONTENT_W - gap * 3) / 4
    bottom = SLIDE_H - M["b"] - (0.62 if s.get("footnote") else 0.10)
    rise = min(0.40, max(0.0, (bottom - y - 2.55) / 3.0))
    tops = [y + (3 - i) * rise for i in range(4)]
    tone = [("7FA6D4", C["bg"]), ("3E82C8", C["brand_lt"]), (C["brand"], C["brand_lt"]),
            (C["gold"], C["gold_lt"])]
    for i, raw in enumerate(steps):
        year, title, money = (_split(raw, 3) + [""] * 3)[:3] if raw.count("｜") >= 2 else (
            _split(raw, 2) + [""])[:3]
        x = M["l"] + i * (w + gap)
        top = tops[i]
        fg, bgc = tone[i]
        rect(slide, x, top, w, bottom - top, fill=bgc)
        rect(slide, x, top, w, 0.66, fill=fg)
        textbox(slide, x, top + 0.14, w, 0.44, year, 23, True,
                C["white"] if i != 3 else "5C4204", align=PP_ALIGN.CENTER, nowrap=True)
        ts = 21 if text_w_in(title, 21) <= (w - 0.5) * 2 else 19
        th = est_lines(title, w - 0.5, ts) * (ts / 72.0) * 1.3
        textbox(slide, x + 0.25, top + 0.92, w - 0.5, th + 0.06, title, ts, True, C["ink"],
                line_spacing=1.3)
        if money:
            ms = 20 if text_w_in(money, 20) <= w - 0.44 else 17
            pill(slide, x + 0.22, bottom - 0.78, w - 0.44, 0.56, money, ms,
                 C["coral"] if i else C["teal"])
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ 巨大数字
def L_big_number(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    num = s.get("big_number", "")
    pw = 6.05
    size = T["huge"]
    while size > 54 and text_w_in(num, size) > pw - 0.5:
        size -= 4
    ph = min(3.25, SLIDE_H - M["b"] - y - (0.60 if s.get("footnote") else 0.05))
    rect(slide, M["l"], y + 0.05, pw, ph, fill=C["brand_lt"])
    nh = size / 72.0 * 1.16
    textbox(slide, M["l"], y + 0.05 + (ph - nh - 0.95) / 2 + 0.12, pw, nh + 0.1, num, size,
            True, C["brand"], align=PP_ALIGN.CENTER, line_spacing=1.16, nowrap=True)
    textbox(slide, M["l"] + 0.3, y + 0.05 + ph - 0.98, pw - 0.6, 0.86,
            s.get("big_number_label", ""), 22, True, C["ink"], align=PP_ALIGN.CENTER,
            line_spacing=1.3)
    bx = M["l"] + pw + 0.45
    bullet_block(slide, bx, y + 0.14, CONTENT_W - pw - 0.45, s.get("body_points", []),
                 size=22, gap=0.34, marker=C["gold"], max_h=ph - 0.2)
    engage_strip(slide, s.get("engagement", ""))
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ クイズ
def L_quiz(slide, s):
    bg(slide, C["bg"])
    rect(slide, 0, 0, SLIDE_W, 0.16, fill=C["gold"])
    if s.get("kicker"):
        textbox(slide, M["l"], 0.72, CONTENT_W, 0.34, s["kicker"].upper(), 17, True,
                C["gold"], spc=2.4, align=PP_ALIGN.CENTER)
    ts = 42 if text_w_in(s["title"], 42) <= CONTENT_W else 34
    th = est_lines(s["title"], CONTENT_W, ts) * (ts / 72.0) * 1.28
    textbox(slide, M["l"], 1.30, CONTENT_W, th + 0.06, s["title"], ts, True, C["ink"],
            align=PP_ALIGN.CENTER, line_spacing=1.28)
    y = 1.30 + th + 0.55
    ch = s.get("body_points", [])[:3]
    gap = 0.28
    w = (CONTENT_W - gap * (len(ch) - 1)) / max(len(ch), 1)
    h = 2.15
    for i, raw in enumerate(ch):
        x = M["l"] + i * (w + gap)
        rect(slide, x, y, w, h, fill=C["white"], line=C["line"], line_w=1.4)
        rect(slide, x + (w - 0.66) / 2, y + 0.26, 0.66, 0.66, fill=C["brand"],
             shape=MSO_SHAPE.OVAL)
        textbox(slide, x + (w - 0.66) / 2, y + 0.36, 0.66, 0.46, "ABC"[i], 26, True,
                C["white"], align=PP_ALIGN.CENTER)
        cs = 24 if text_w_in(raw, 24) <= (w - 0.5) * 2 else 21
        textbox(slide, x + 0.25, y + 1.10, w - 0.5, 0.92, raw, cs, True, C["ink"],
                align=PP_ALIGN.CENTER, line_spacing=1.3)
        y2 = y
    engage_strip(slide, s.get("engagement", ""), y=y + h + 0.30)

# ------------------------------------------------------------------ 棒グラフ
def L_stat_bar(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    if s.get("lead"):
        textbox(slide, M["l"], y, CONTENT_W, 0.46, s["lead"], 22, True, C["brand"])
        y += 0.62
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
    lab_w = 3.15
    bar_x = M["l"] + lab_w + 0.25
    bar_max = CONTENT_W - lab_w - 0.25 - 1.65
    bh = 0.62
    gapy = 0.34
    for lab, val, v, hi in rows:
        textbox(slide, M["l"], y + 0.10, lab_w, 0.5, lab, 21 if len(lab) <= 9 else 18,
                hi, C["ink"] if hi else C["gray"], align=PP_ALIGN.RIGHT)
        wpx = max(bar_max * (v / vmax), 0.12)
        rect(slide, bar_x, y, wpx, bh, fill=C["coral"] if hi else "C9D4E0")
        textbox(slide, bar_x + wpx + 0.18, y + 0.02, 1.5, 0.6, val, 29 if hi else 22, True,
                C["coral"] if hi else C["gray"])
        y += bh + gapy
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ 写真主役
def L_photo_focus(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    pw = 6.85
    photo_slot(slide, M["l"], y, pw, SLIDE_H - M["b"] - y - (0.55 if s.get("footnote") else 0),
               s.get("photo_caption", ""), s.get("photo_slot", ""))
    bx = M["l"] + pw + 0.5
    bullet_block(slide, bx, y + 0.08, CONTENT_W - pw - 0.5, s.get("body_points", []),
                 size=22, gap=0.34, max_h=SLIDE_H - M["b"] - y - 0.6)
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ 2カラム
def L_two_column(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    lw = 6.35
    if s.get("lead"):
        n = est_lines(s["lead"], lw, T["lead"])
        hh = n * (T["lead"] / 72.0) * 1.35
        textbox(slide, M["l"], y, lw, hh + 0.05, s["lead"], T["lead"], True, C["brand"],
                line_spacing=1.35)
        y2 = y + hh + 0.36
    else:
        y2 = y
    bullet_block(slide, M["l"], y2, lw, s.get("body_points", []), size=22, gap=0.32,
                 max_h=SLIDE_H - M["b"] - y2 - (0.6 if s.get("footnote") else 0.1))
    px = M["l"] + lw + 0.45
    pwid = CONTENT_W - lw - 0.45
    panel = s.get("panel", {})
    if panel:
        ph = SLIDE_H - M["b"] - y - (0.6 if s.get("footnote") else 0.1)
        rect(slide, px, y, pwid, ph, fill=C["bg"])
        rect(slide, px, y, pwid, 0.09, fill=C["gold"])
        textbox(slide, px + 0.32, y + 0.38, pwid - 0.64, 0.5, panel.get("title", ""), 21,
                True, C["ink"])
        cy = y + 1.02
        for it in panel.get("items", [])[:4]:
            k, v = _split(it, 2)
            textbox(slide, px + 0.32, cy, pwid - 0.64, 0.34, k, 17, False, C["gray"])
            vs = 26 if text_w_in(v, 26) <= pwid - 0.64 else 21
            textbox(slide, px + 0.32, cy + 0.36, pwid - 0.64, 0.5, v, vs, True, C["coral"])
            cy += 1.02
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ まとめ
def L_summary(slide, s):
    bg(slide, C["white"])
    y = header(slide, s.get("kicker", ""), s["title"])
    rows = s.get("body_points", [])[:3]
    eng_h = 1.02 if s.get("engagement") else 0.0
    fn_h = 0.55 if s.get("footnote") else 0.0
    gap = 0.24
    avail = SLIDE_H - M["b"] - y - eng_h - fn_h
    h = min(1.34, (avail - gap * (len(rows) - 1)) / max(len(rows), 1))
    for i, raw in enumerate(rows):
        rect(slide, M["l"], y, CONTENT_W, h, fill=C["bg"])
        rect(slide, M["l"], y, 0.11, h, fill=[C["brand"], C["teal"], C["gold"]][i % 3])
        textbox(slide, M["l"] + 0.42, y + (h - 0.62) / 2, 0.8, 0.68, str(i + 1), 40, True,
                [C["brand"], C["teal"], C["gold"]][i % 3], align=PP_ALIGN.LEFT, nowrap=True)
        size = 25
        while size > 20 and est_lines(raw, CONTENT_W - 2.0, size) * (size / 72.0) * 1.34 > h - 0.3:
            size -= 1
        textbox(slide, M["l"] + 1.35, y + 0.16, CONTENT_W - 1.85, h - 0.32, raw, size, True,
                C["ink"], line_spacing=1.34, anchor=MSO_ANCHOR.MIDDLE)
        y += h + gap
    engage_strip(slide, s.get("engagement", ""))
    footnote(slide, s.get("footnote", ""))

# ------------------------------------------------------------------ クロージング
def L_closing(slide, s):
    bg(slide, C["navy"])
    rect(slide, 0, 0, SLIDE_W, 0.16, fill=C["gold"])
    ts = 46 if text_w_in(s["title"], 46) <= 11.0 else 38
    th = est_lines(s["title"], 11.0, ts) * (ts / 72.0) * 1.3
    textbox(slide, 1.15, 1.75, 11.0, th + 0.1, s["title"], ts, True, C["white"],
            align=PP_ALIGN.CENTER, line_spacing=1.3)
    y = 1.75 + th + 0.5
    rect(slide, (SLIDE_W - 1.9) / 2, y, 1.9, 0.09, fill=C["gold"])
    y += 0.5
    for p in s.get("body_points", [])[:3]:
        textbox(slide, 1.15, y, 11.0, 0.5, p, 23, False, "BFD3E8", align=PP_ALIGN.CENTER)
        y += 0.58
    if s.get("lead"):
        textbox(slide, 1.15, SLIDE_H - 1.15, 11.0, 0.5, s["lead"], 20, True, C["gold"],
                align=PP_ALIGN.CENTER)

RENDER = {
    "title": L_title, "section": L_section, "agenda": L_agenda, "bullets": L_bullets,
    "three_cards": L_three_cards, "four_steps": L_four_steps, "big_number": L_big_number,
    "quiz": L_quiz, "stat_bar": L_stat_bar, "photo_focus": L_photo_focus,
    "two_column": L_two_column, "summary": L_summary, "closing": L_closing,
}
