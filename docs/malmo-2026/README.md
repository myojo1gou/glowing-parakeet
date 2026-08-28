# Malmö 2026 — Kitakyushu / Japan case material

Documents prepared for **Per-Arne Nilsson** (Senior Strategist, Environmental Department,
City of Malmö), for the meeting in Malmö on **9 September 2026** during the Nordic field study
(7–11 September 2026: Danish Energy Agency · Copenhagen Business School · Malmö · Lolland).

## Files

| File | What it is |
|---|---|
| `japan-transition-report.html` / `.pdf` | **Main report**, 21 pages A4. Nine Japanese cases read for ownership, financing and risk; synthesis; fourteen questions for Malmö. |
| `kitakyushu-briefing.html` / `.pdf` | **Short handout**, 8 pages A4. Kitakyushu only. Kept as the version to hand over at the meeting itself. |

Both are print-styled A4, self-contained, no external assets. The byline on page 1 of each
(`[name]`, `[affiliation]`) is a highlighted placeholder — fill it in the HTML before re-rendering.

## Rebuilding the PDFs

```sh
chromium --headless --disable-gpu --no-sandbox --no-pdf-header-footer \
  --print-to-pdf=japan-transition-report.pdf japan-transition-report.html
chromium --headless --disable-gpu --no-sandbox --no-pdf-header-footer \
  --print-to-pdf=kitakyushu-briefing.pdf kitakyushu-briefing.html
```

Any Chromium build works. Japanese glyphs in the source list need a CJK font installed
(IPAGothic is used here).

## Structure of the main report

- **Part I — Framing.** The three questions (who owns / who pays / who carries risk), Kitakyushu at
  a glance including the 64% industrial share of emissions, four phases 1901–present, and the
  national offshore-wind policy timeline.
- **Part II — Nine cases.** Each read the same way: what it is, the numbers, the ownership and
  financing structure, what it demonstrates, what is unresolved.
  1. Municipal estate — third-party ownership PPA at scale (Decarbonisation Leading Area)
  2. Second-life assets — the reuse-panel demonstration at Hibikinada Biotope
  3. Eco-Town — Shinryo and Recycle Tech, circular industry
  4. Hibiki offshore wind farm — a municipally-tendered consortium
  5. Green Energy Port Hibiki — the port as industrial policy
  6. FLOAT RAISER — a state-subsidised floating-wind construction vessel
  7. Ørsted in Japan — the ownership model Japan does not have
  8. Gotō, Nagasaki — E-WIND, citizens' power, a training association (the counter-case)
  9. FAIS and the Science and Research Park — thirty years, weak economic result
- **Part III — Synthesis.** Nine-case comparison table, five findings, the four missing layers.
- **Part IV — Fourteen questions** for the City of Malmö, grouped A–E.
- **Part V — Sources and method**, plus a note on what the report leaves out.

## Source material

Documents from the Kitakyushu study visit of **22 September 2023** by the Kyoto University Public
Finance Study Group, held in Google Drive (folder `SCANLINE`):

| File | Case |
|---|---|
| `IMG_0023.pdf` | 北九州市環境局「脱炭素先行地域にかかる北九州市の取り組み」2023-09-22 — Cases 1, 3; Part I |
| `IMG_0024.pdf` / `IMG_0025.pdf` | リユースパネル実証（新菱・北九州パワー・東京センチュリー・北九州市）— Case 2 |
| `北九州0026/0027/0028.pdf` | 新菱／リサイクルテック — Case 3 |
| `北九州0002.pdf` | ひびきウインドエナジー「北九州響灘洋上ウインドファーム建設工事の概要」— Case 4 |
| `北九州0005.pdf` | グリーンエネルギーポートひびき — Case 5 |
| `北九州0006.pdf` | Next Generation Energy Kitakyushu — Part I, Case 5 |
| `北九州0003.pdf` | FLOAT RAISER（戸田建設・吉田組／OWFC）— Case 6 |
| `北九州0001.pdf` | オーステッド・ジャパン会社概要 — Case 7 |
| `北九州0004.pdf` | 有限会社イー・ウィンド（五島市）— Case 8 |

Case 9 draws on the author's unpublished research on Kitakyushu's industrial structure and the
Science and Research Park (FAIS), based on *Kitakyushu City Economic Accounts* (1996–2020) and FAIS
programme documents, with Pittsburgh as the comparator.

No claims are made about Malmö, Lolland or Swedish/Danish practice — those are posed as questions
in Part IV instead.
