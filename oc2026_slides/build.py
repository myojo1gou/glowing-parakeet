# -*- coding: utf-8 -*-
import json, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Inches
import deckkit
from deckkit import SLIDE_W, SLIDE_H, new_slide, notes, pagenum, WARNINGS
from layouts import RENDER

def build(spec_path, out_path):
    spec = json.load(open(spec_path, encoding="utf-8"))
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    total = 0
    for i, s in enumerate(spec["slides"]):
        slide = new_slide(prs)
        fn = RENDER.get(s["layout"])
        if fn is None:
            raise SystemExit(f"unknown layout {s['layout']} @ slide {i+1}")
        fn(slide, s)
        if s["layout"] not in ("title", "closing"):
            pagenum(slide, i + 1)
        notes(slide, s.get("speaker_notes", ""))
        total += int(s.get("est_seconds", 0))
    prs.save(out_path)
    print(f"saved: {out_path}  slides={len(spec['slides'])}  想定{total//60}分{total%60}秒")
    if WARNINGS:
        print(f"\n--- 収まり警告 {len(WARNINGS)}件 ---")
        for w in WARNINGS:
            print(" ", w)
    else:
        print("収まり警告なし")

if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
