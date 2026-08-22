#!/bin/bash
# pptx -> pdf -> png（目視確認用）
set -e
PPTX="$1"; OUT="${2:-/home/user/glowing-parakeet/oc2026_slides/preview}"
mkdir -p "$OUT"; rm -f "$OUT"/*.png "$OUT"/*.pdf
soffice --headless --convert-to pdf --outdir "$OUT" "$PPTX" >/dev/null 2>&1
PDF="$OUT/$(basename "${PPTX%.pptx}").pdf"
python3 - "$PDF" "$OUT" <<'PY'
import sys, pymupdf
doc = pymupdf.open(sys.argv[1]); out = sys.argv[2]
for i, page in enumerate(doc):
    page.get_pixmap(dpi=110).save(f"{out}/s{i+1:02d}.png")
print(f"rendered {len(doc)} pages -> {out}")
PY
