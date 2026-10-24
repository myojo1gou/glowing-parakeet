#!/usr/bin/env python3
"""Fill the RESER 2026 template with the paper's content.
Usage: python3 build_deck.py <unpacked_dir>
"""
import os, re, shutil, sys
from xml.sax.saxutils import escape

UNP = sys.argv[1]
EMU = 914400

# ---- palette (sampled from the template artwork) ----
NAVY = "193F61"
BLUE = "104486"
GREEN = "129152"
SAGE = "4C886C"
YELLOW = "EEBC37"
LIGHT = "ADC1D8"
WHITE = "FFFFFF"
GREY = "5A6B7C"
INK = "1F2A37"
FONT = "Calibri"

_id = [1000]
def nid():
    _id[0] += 1
    return _id[0]

def emu(v):
    return str(int(round(v * EMU)))

# ---------- text helpers ----------
def run(text, sz=14, color=INK, b=False, i=False, font=FONT):
    rpr = f'<a:rPr lang="en-US" sz="{int(sz*100)}" b="{1 if b else 0}" i="{1 if i else 0}" dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{font}"/><a:cs typeface="{font}"/></a:rPr>'
    sp = ' xml:space="preserve"' if text != text.strip() or "  " in text else ""
    return f'<a:r>{rpr}<a:t{sp}>{escape(text)}</a:t></a:r>'

def para(runs, algn="l", bullet=None, lvl=0, spcAft=6, spcBef=0, indent_in=0.22, lnSpc=None):
    """runs: list of run xml strings, or a plain string."""
    if isinstance(runs, str):
        runs = [run(runs)]
    attrs = f'algn="{algn}"'
    if lvl:
        attrs += f' lvl="{lvl}"'
    inner = ""
    if lnSpc:
        inner += f'<a:lnSpc><a:spcPct val="{int(lnSpc*100000)}"/></a:lnSpc>'
    inner += f'<a:spcBef><a:spcPts val="{int(spcBef*100)}"/></a:spcBef><a:spcAft><a:spcPts val="{int(spcAft*100)}"/></a:spcAft>'
    if bullet:
        attrs += f' marL="{emu(indent_in*(lvl+1))}" indent="-{emu(indent_in)}"'
        inner += f'<a:buClr><a:srgbClr val="{bullet[1] if isinstance(bullet, tuple) else SAGE}"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="{bullet[0] if isinstance(bullet, tuple) else bullet}"/>'
    else:
        inner += "<a:buNone/>"
    return f'<a:p><a:pPr {attrs}>{inner}</a:pPr>{"".join(runs)}</a:p>'

def bullets(items, sz=14, color=INK, char="•", bcolor=SAGE, spcAft=6, lvl=0):
    out = []
    for it in items:
        if isinstance(it, str):
            out.append(para([run(it, sz, color)], bullet=(char, bcolor), spcAft=spcAft, lvl=lvl))
        else:  # list of runs
            out.append(para(it, bullet=(char, bcolor), spcAft=spcAft, lvl=lvl))
    return out

# ---------- shape helpers ----------
def shape(x, y, w, h, paras=None, name="Shape", geom="rect", fill=None, alpha=None,
          line=None, lw=0.75, anchor="t", inset=(0.1, 0.06, 0.1, 0.06), radius=None,
          wrap=True, shadow=False, dash=None, txbox=False):
    i = nid()
    sppr = f'<a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
    av = ""
    if geom == "roundRect":
        av = f'<a:gd name="adj" fmla="val {int((radius or 0.08)*50000)}"/>'
    sppr += f'<a:prstGeom prst="{geom}"><a:avLst>{av}</a:avLst></a:prstGeom>'
    if fill:
        a = f'<a:alpha val="{int(alpha*1000)}"/>' if alpha is not None else ""
        sppr += f'<a:solidFill><a:srgbClr val="{fill}">{a}</a:srgbClr></a:solidFill>'
    else:
        sppr += "<a:noFill/>"
    if line:
        d = f'<a:prstDash val="{dash}"/>' if dash else ""
        sppr += f'<a:ln w="{int(lw*12700)}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill>{d}</a:ln>'
    else:
        sppr += "<a:ln><a:noFill/></a:ln>"
    if shadow:
        sppr += '<a:effectLst><a:outerShdw blurRad="63500" dist="25400" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="193F61"><a:alpha val="18000"/></a:srgbClr></a:outerShdw></a:effectLst>'
    l, t, r, b = inset
    body = f'<a:bodyPr wrap="{"square" if wrap else "none"}" lIns="{emu(l)}" tIns="{emu(t)}" rIns="{emu(r)}" bIns="{emu(b)}" rtlCol="0" anchor="{anchor}"><a:normAutofit/></a:bodyPr><a:lstStyle/>'
    body += "".join(paras) if paras else '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    tx = ' txBox="1"' if txbox else ""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="{escape(name)}"/><p:cNvSpPr{tx}/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr>{sppr}</p:spPr><p:txBody>{body}</p:txBody></p:sp>')

def text(x, y, w, h, paras, name="Text", anchor="t", inset=(0, 0, 0, 0), **kw):
    return shape(x, y, w, h, paras, name=name, anchor=anchor, inset=inset, txbox=True, **kw)

def title(t, sub=None):
    """Slide title at a fixed position on every content slide."""
    out = [text(0.55, 1.32, 12.2, 0.6, [para([run(t, 26, NAVY, b=True)], spcAft=0)], name="Title", anchor="t")]
    if sub:
        out.append(text(0.55, 1.95, 12.2, 0.4, [para([run(sub, 14, GREY, i=True)], spcAft=0)], name="Subtitle"))
    return out

def card(x, y, w, h, paras, fill=WHITE, alpha=92, line=None, anchor="t", inset=(0.18, 0.14, 0.18, 0.12), name="Card", shadow=True, radius=0.06):
    return shape(x, y, w, h, paras, name=name, geom="roundRect", fill=fill, alpha=alpha, line=line,
                 anchor=anchor, inset=inset, shadow=shadow, radius=radius)

def badge(x, y, d, label, fill=NAVY, color=WHITE, sz=16, name="Badge"):
    return shape(x, y, d, d, [para([run(label, sz, color, b=True)], algn="ctr", spcAft=0)], name=name,
                 geom="ellipse", fill=fill, anchor="ctr", inset=(0, 0, 0, 0))

def arrow(x, y, w, h, geom="downArrow", fill=SAGE, name="Arrow"):
    return shape(x, y, w, h, None, name=name, geom=geom, fill=fill)

def table(x, y, w, colw, rows, sz=11, header_fill=NAVY, header_color=WHITE, row_h=0.4, header_h=0.42,
          first_col_bold=True, zebra="F3F6F9", name="Table", col_align=None):
    """rows[0] is the header. colw: relative widths. Cells: str or list of paragraph xml."""
    i = nid()
    tot = sum(colw)
    cols = [w * c / tot for c in colw]
    grid = "".join(f'<a:gridCol w="{emu(c)}"/>' for c in cols)
    trs = []
    for ri, r in enumerate(rows):
        is_h = ri == 0
        h = header_h if is_h else row_h
        tcs = []
        for ci, cell in enumerate(r):
            algn = (col_align[ci] if col_align else "l")
            if isinstance(cell, str):
                b = is_h or (first_col_bold and ci == 0)
                col = header_color if is_h else (NAVY if (first_col_bold and ci == 0) else INK)
                cell_paras = [para([run(cell, sz, col, b=b)], algn=algn, spcAft=0)]
            else:
                cell_paras = cell
            fill = header_fill if is_h else (zebra if (ri % 2 == 0 and zebra) else WHITE)
            bdr = f'<a:lnL w="6350"><a:solidFill><a:srgbClr val="D5DCE4"/></a:solidFill></a:lnL><a:lnR w="6350"><a:solidFill><a:srgbClr val="D5DCE4"/></a:solidFill></a:lnR><a:lnT w="6350"><a:solidFill><a:srgbClr val="D5DCE4"/></a:solidFill></a:lnT><a:lnB w="6350"><a:solidFill><a:srgbClr val="D5DCE4"/></a:solidFill></a:lnB>'
            tcs.append(f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>{"".join(cell_paras)}</a:txBody>'
                       f'<a:tcPr marL="{emu(0.07)}" marR="{emu(0.07)}" marT="{emu(0.04)}" marB="{emu(0.04)}" anchor="{"ctr" if is_h else "t"}">{bdr}<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></a:tcPr></a:tc>')
        trs.append(f'<a:tr h="{emu(h)}">{"".join(tcs)}</a:tr>')
    total_h = header_h + row_h * (len(rows) - 1)
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{i}" name="{escape(name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(total_h)}"/></p:xfrm>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>{grid}</a:tblGrid>{"".join(trs)}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')

def cell(parts, sz=10.5, algn="l"):
    """Table cell with one paragraph per part; a part may be str or (str, bold)."""
    out = []
    for p in parts:
        if isinstance(p, tuple):
            out.append(para([run(p[0], sz, NAVY if p[1] else INK, b=p[1])], algn=algn, spcAft=2))
        else:
            out.append(para([run(p, sz, INK)], algn=algn, spcAft=2))
    return out

# ---------- slide content ----------
S = {}      # slide number -> list of shape xml
N = {}      # slide number -> speaker notes

def R(t, sz=14, color=INK, **kw):
    return run(t, sz, color, **kw)

# Slide 1 : title -------------------------------------------------------
S[1] = [
    text(1.25, 2.72, 10.9, 1.05,
         [para([R("Reconfiguring Inherited Service Infrastructures for Green Public Service Transition", 28, NAVY, b=True)], spcAft=0, lnSpc=1.0)],
         name="Deck title"),
    text(1.25, 3.72, 10.9, 0.55,
         [para([R("Municipally owned arrangements in Helsinki, Aalborg and Skåne", 18, GREEN, b=True)], spcAft=0)],
         name="Deck subtitle"),
    shape(1.07, 4.48, 11.19, 0.78, None, name="Author panel", fill=WHITE, alpha=82),
    text(1.25, 4.55, 7.4, 0.65,
         [para([R("[Presenter name]", 15, NAVY, b=True), R("  ·  [Affiliation]", 15, INK)], spcAft=2),
          para([R("[Co-author names and affiliations]", 12.5, GREY)], spcAft=0)],
         name="Authors", anchor="ctr"),
    text(8.7, 4.55, 3.45, 0.65,
         [para([R("36th RESER International Conference", 12.5, NAVY, b=True)], algn="r", spcAft=2),
          para([R("University of Brasília, Brazil · 2026", 12, GREY)], algn="r", spcAft=0)],
         name="Conference", anchor="ctr"),
]
N[1] = ("Title. Introduce yourself and co-authors. One sentence on the puzzle: municipalities commit politically to green transition, "
        "but heat and waste still have to be delivered every day, and the companies that deliver them must judge whether an investment pays. "
        "This paper looks at how those two logics meet in actual investment decisions in three Nordic municipal arrangements.")

# Slide 2 : motivation + research question -------------------------------
S[2] = title("Green transition is decided one investment at a time") + [
    text(0.55, 2.1, 6.1, 4.6, [
        para([R("The setting", 15, NAVY, b=True)], spcAft=4),
        *bullets([
            "Municipalities seek lower emissions and more resource recovery, yet homes must stay heated and waste treated during the change.",
            "Existing networks offer a route to new production technologies, but replacing plants or adding recovery requires capital and dependable income.",
            "A climate target gives a reason to invest; it leaves technology, expenditure and responsibility to owners and companies.",
        ], sz=14, spcAft=7),
        para([R("Why municipally owned enterprises", 15, NAVY, b=True)], spcAft=4, spcBef=6),
        *bullets([
            "The owner can instruct the company and support borrowing, while depending on distributions and facing demands for affordable services.",
            "Managers weigh reliability, returns and debt service. Investment decisions make this relationship observable.",
        ], sz=14, spcAft=7),
    ], name="Setting"),
    card(7.0, 2.1, 5.75, 2.55, [
        para([R("Research question", 13, SAGE, b=True)], spcAft=6),
        para([R("How do political objectives and business appraisal interact in municipal green transition investment, and how do institutions and ownership practices shape that interaction?", 15.5, NAVY, b=True)], spcAft=0, lnSpc=1.05),
    ], fill=WHITE, alpha=94),
    card(7.0, 4.85, 5.75, 1.85, [
        para([R("What we look at", 13, SAGE, b=True)], spcAft=6),
        para([R("How public owners respond when a politically supported investment still faces an adverse appraisal or an unsettled revenue basis: ", 13.5, INK),
              R("adapting a production pathway, allocating costs through a supply agreement, qualifying support through financing and purchasing conditions.", 13.5, NAVY, b=True)], spcAft=0),
    ], fill=WHITE, alpha=94),
]
N[2] = ("Frame the problem: the climate target is given politically, but it does not say which plant, at what cost, paid by whom. "
        "Municipally owned enterprises sit at exactly that intersection. State the research question and preview the three types of response "
        "we find in the records: staging (Helsinki), contractual cost allocation (Aalborg), conditional financing and purchasing (Sysav/Malmö).")

# Slide 3 : analytical framework (Figure 1 rebuilt) --------------------
fx = 0.55
S[3] = title("Analytical framework: how investability forms") + [
    shape(0.5, 2.0, 8.05, 4.5, None, name="Figure backdrop", geom="roundRect", fill=WHITE, alpha=86, radius=0.03),
    shape(fx, 2.12, 7.95, 0.6, [
        para([R("Institutional conditions", 14, WHITE, b=True), R("   rules and established practices: authority · cost recovery · purchasing conditions", 12, WHITE)], algn="ctr", spcAft=0)],
        name="Institutions", geom="roundRect", fill=NAVY, anchor="ctr", radius=0.15),
    arrow(4.33, 2.75, 0.4, 0.24, fill=LIGHT),
    shape(fx, 3.02, 3.4, 0.85, [
        para([R("Political objectives", 14, NAVY, b=True)], algn="ctr", spcAft=2),
        para([R("publicly stated reasons for intervention", 11.5, GREY)], algn="ctr", spcAft=0)],
        name="Political", geom="roundRect", fill=WHITE, line=NAVY, lw=1.25, anchor="ctr"),
    arrow(4.05, 3.3, 0.95, 0.28, geom="leftRightArrow", fill=YELLOW),
    text(3.95, 3.6, 1.15, 0.3, [para([R("tension or alignment", 9.5, GREY, i=True)], algn="ctr", spcAft=0)], name="Tension label"),
    shape(5.1, 3.02, 3.4, 0.85, [
        para([R("Business appraisal", 14, NAVY, b=True)], algn="ctr", spcAft=2),
        para([R("costs, revenues, financing, operating requirements", 11.5, GREY)], algn="ctr", spcAft=0)],
        name="Appraisal", geom="roundRect", fill=WHITE, line=NAVY, lw=1.25, anchor="ctr"),
    text(fx, 3.9, 7.95, 0.26, [para([R("municipal councils and company boards: appraisal and authorisation", 11, GREY, i=True)], algn="ctr", spcAft=0)], name="Actors"),
    shape(fx, 4.22, 2.5, 0.8, [para([R("Owner capital and guarantees", 12.5, INK, b=True)], algn="ctr", spcAft=2), para([R("financing support", 11, SAGE, i=True)], algn="ctr", spcAft=0)],
          name="Ch1", geom="roundRect", fill="E3EAF2", anchor="ctr"),
    shape(fx + 2.725, 4.22, 2.5, 0.8, [para([R("Prospective owner purchases", 12.5, INK, b=True)], algn="ctr", spcAft=2), para([R("conditional revenue", 11, SAGE, i=True)], algn="ctr", spcAft=0)],
          name="Ch2", geom="roundRect", fill="E3EAF2", anchor="ctr"),
    shape(fx + 5.45, 4.22, 2.5, 0.8, [para([R("Agreements between companies", 12.5, INK, b=True)], algn="ctr", spcAft=2), para([R("cost allocation", 11, SAGE, i=True)], algn="ctr", spcAft=0)],
          name="Ch3", geom="roundRect", fill="E3EAF2", anchor="ctr"),
    arrow(1.6, 5.06, 0.4, 0.24, fill=LIGHT), arrow(6.8, 5.06, 0.4, 0.24, fill=LIGHT),
    shape(fx, 5.36, 3.6, 0.55, [para([R("Approved financing", 14, WHITE, b=True)], algn="ctr", spcAft=0)], name="Financing", geom="roundRect", fill=GREEN, anchor="ctr"),
    text(4.15, 5.36, 0.7, 0.55, [para([R("+", 22, NAVY, b=True)], algn="ctr", spcAft=0)], name="Plus", anchor="ctr"),
    shape(4.9, 5.36, 3.6, 0.55, [para([R("Credible cost recovery", 14, WHITE, b=True)], algn="ctr", spcAft=0)], name="Recovery", geom="roundRect", fill=GREEN, anchor="ctr"),
    text(fx, 5.98, 7.95, 0.4, [para([R("= Investability", 15, NAVY, b=True), R("    the definition used in this study", 11.5, GREY, i=True)], algn="ctr", spcAft=0)],
         name="Investability", anchor="ctr"),
    card(8.85, 2.0, 3.9, 2.05, [
        para([R("Investability", 14, GREEN, b=True)], spcAft=4),
        para([R("Approved financing combined with credible cost recovery. It asks who accepts the financial commitments and how investment and provision costs are expected to be met, rather than scoring completed business cases.", 12.5, INK)], spcAft=0),
    ]),
    card(8.85, 4.2, 3.9, 1.3, [
        para([R("Revisability", 14, YELLOW, b=True), R("  (secondary)", 11.5, GREY)], spcAft=4),
        para([R("Changes in technology, scale, timing or activities: can the two conditions be maintained through the change?", 12.5, INK)], spcAft=0),
    ]),
    card(8.85, 5.65, 3.9, 1.15, [para([R("Source: Authors' own elaboration (Figure 1 of the paper). Both terms are analytical tools, following Bruton et al. (2015) rather than assigning political and economic rationality to separate camps.", 10, GREY)], spcAft=0)], alpha=86, shadow=False, inset=(0.15, 0.1, 0.15, 0.08)),
]
N[3] = ("Walk the figure top to bottom. Institutions decide who may act and what enters the economic evaluation. Councils and boards hold political "
        "objectives and business appraisal in tension or alignment. Owners can intervene through three channels: capital and guarantees, prospective "
        "purchases, and agreements between group companies. Investability is defined narrowly as approved financing plus credible cost recovery. "
        "Revisability is the secondary question.")

# Slide 4 : literature positioning ---------------------------------------
cw, gap = 3.9, 0.25
cx = [0.55, 0.55 + cw + gap, 0.55 + 2 * (cw + gap)]
lit = [
    ("Municipal climate governance", NAVY,
     "Service provision as a means of implementing climate policy (Kern & Alber 2009). Mayors also back renewables for employment, earnings and local development (Busch & McCormick 2014).",
     "We therefore record the purposes invoked for each investment."),
    ("Public service ecosystems and public value", GREEN,
     "Value creation mediated among actors, institutions and competing values (Osborne et al. 2022; Trischler et al. 2023); conflict over public value foregrounded (Skålén et al. 2024).",
     "We examine that mediation at the moment of investment, when purposes meet assets, expenditure and obligations."),
    ("Municipally owned enterprises, multiple principals", SAGE,
     "Hybrid organisations combining public control and corporate form (Bruton et al. 2015). Joint delivery raises collective action problems (Voorn et al. 2019; Voorn & van Genugten 2022).",
     "We ask whose agreement is required and how a collective position becomes an owner's commitment."),
]
S[4] = title("Where the paper sits")
for (hd, col, body, use), x in zip(lit, cx):
    S[4] += [
        card(x, 2.15, cw, 3.45, [
            para([R(hd, 15, col, b=True)], spcAft=8),
            para([R(body, 13, INK)], spcAft=10),
            para([R("In this paper: ", 12.5, col, b=True), R(use, 12.5, INK, i=True)], spcAft=0),
        ], inset=(0.2, 0.55, 0.2, 0.15)),
        badge(x + 0.2, 1.98, 0.42, "", fill=col),
    ]
S[4] += [
    card(0.55, 5.8, 12.2, 0.95, [
        para([R("Gap addressed: ", 14, NAVY, b=True),
              R("investment decisions are where an owner endorses a purpose, provides capital or considers purchasing an output, each with different consequences for the company's business case. Most studies stop at objectives; we follow expenditure, income and obligations.", 13.5, INK)], spcAft=0)],
        fill="EEF3F8", alpha=96, anchor="ctr", shadow=False),
]
N[4] = ("Three literatures. Climate governance tells us why municipalities invest; the public service ecosystem literature tells us that value is "
        "negotiated among actors and institutions; the MOE and multiple-principals literature tells us that owners and managers do not automatically agree. "
        "Our contribution is to look at the point where those three meet: the investment decision and its financing and recovery terms.")

# Slide 5 : research design -----------------------------------------------
cases = [
    ("H", "Helsinki · Helen Oy", NAVY, WHITE, "District heat and power",
     "City-owned company incorporated 2015 (Act 626/2013). Single owner; council and board decisions on production pathway and capital."),
    ("A", "Aalborg · municipal heat group", GREEN, WHITE, "District heat",
     "Aalborg Forsyning A/S group: Aalborg Varme A/S buys heat from Nordjyllandsværket A/S. Council guarantees; intra-group agreements."),
    ("S", "Skåne · Sysav", YELLOW, NAVY, "Waste treatment and resource recovery",
     "Owned jointly by 14 municipalities since the 1970s. Malmö (45.98%) as the largest owner; subsidiary Sysav Industri AB."),
]
S[5] = title("Research design: three arrangements, fifteen documentary records")
for k, (ltr, nm, col, tcol, sector, desc) in enumerate(cases):
    x = 0.55 + k * (cw + gap)
    S[5] += [
        card(x, 2.15, cw, 2.0, [
            para([R(nm, 14.5, NAVY, b=True)], spcAft=2),
            para([R(sector, 11.5, GREY, i=True)], spcAft=6),
            para([R(desc, 12.5, INK)], spcAft=0),
        ], inset=(0.2, 0.2, 0.2, 0.15)),
        badge(x + cw - 0.75, 2.3, 0.55, ltr, fill=col, color=tcol, sz=18),
    ]
S[5] += [
    card(0.55, 4.4, 6.0, 2.4, [
        para([R("Case logic", 14, SAGE, b=True)], spcAft=4),
        *bullets([
            "Exploratory, purposive comparison (Seawright & Gerring 2008); not national models, not an isolated ownership effect.",
            "Case = the municipal service arrangement; observation = a documented investment decision sequence.",
            "Started from three substantive expectations (Finnish earnings and green investment, Aalborg mayoral leadership, Sysav's long cooperation) and kept the records that qualify them.",
        ], sz=12.5, spcAft=4),
    ], inset=(0.2, 0.14, 0.2, 0.1)),
    card(6.75, 4.4, 6.0, 2.4, [
        para([R("Documents and procedure", 14, SAGE, b=True)], spcAft=4),
        *bullets([
            "15 records (H1–H5, A1–A6, S1–S4): council minutes, annual reports, agreements, regulator guidance, CJEU material, press; Finnish, Danish and Swedish read in the original.",
            "Document analysis after Bowen (2009): authorship, purpose, proximity to the decision.",
            "Per episode: purposes invoked, economic assessment, action taken, actor and authority. Financing evidence and cost-recovery evidence read separately; a blank is a limit of the record, not proof of inaction.",
        ], sz=12.5, spcAft=4),
    ], inset=(0.2, 0.14, 0.2, 0.1)),
]
N[5] = ("Three arrangements with different owner structures: a single city owner, a municipal group with several companies, and a 14-owner company. "
        "Emphasise that this is exploratory and purposive. Fifteen records, read in the original languages, each coded for purposes, appraisal, "
        "action and actor. The key methodological point: we never treat a blank cell as evidence that nothing happened.")

# Slide 6..8 : case slides with timeline -----------------------------------
def timeline(events, y, col, x0=0.55, w=12.2):
    """events: list of (date, label). Draws a horizontal line with dots and labels alternating above/below."""
    out = [shape(x0, y, w, 0.04, None, name="Timeline", fill=LIGHT)]
    n = len(events)
    step = w / n
    for k, (d, lab) in enumerate(events):
        cx_ = x0 + step * (k + 0.5)
        out.append(shape(cx_ - 0.11, y - 0.09, 0.22, 0.22, None, name=f"Dot {k}", geom="ellipse", fill=col, line=WHITE, lw=1.5))
        out.append(text(cx_ - step / 2 + 0.05, y - 0.62, step - 0.1, 0.45,
                        [para([R(d, 12, col, b=True)], algn="ctr", spcAft=0)], name=f"Date {k}", anchor="b"))
        out.append(text(cx_ - step / 2 + 0.05, y + 0.2, step - 0.1, 1.05,
                        [para([R(lab, 11.5, INK)], algn="ctr", spcAft=0)], name=f"Event {k}"))
    return out

def case_slide(n, ltr, col, tcol, t, sub, events, left_hd, left_items, right_hd, right_paras, tl_y=3.0):
    S[n] = title(t, sub) + [badge(12.2, 1.3, 0.55, ltr, fill=col, color=tcol, sz=18)]
    S[n] += timeline(events, tl_y, col)
    S[n] += [
        card(0.55, 4.5, 6.4, 2.3, [para([R(left_hd, 14, col if col != YELLOW else NAVY, b=True)], spcAft=4)] +
             bullets(left_items, sz=12.5, spcAft=4), inset=(0.2, 0.14, 0.2, 0.1)),
        card(7.15, 4.5, 5.6, 2.3, [para([R(right_hd, 14, GREEN, b=True)], spcAft=4)] + right_paras,
             fill="EEF3F8", alpha=96, inset=(0.2, 0.14, 0.2, 0.1)),
    ]

case_slide(6, "H", NAVY, WHITE,
    "Helsinki: political objective survives an adverse appraisal",
    "Helen Oy, decentralised heat pathway and coal closures (records H1–H5)",
    [("Jun 2014", "Council incorporates Helen; capital plan targets ≥50% equity through the 2016–2021 investment peak"),
     ("2 Dec 2015", "Council selects decentralised production (EUR 360m) over new CHP (EUR 960m) and Hanasaari closure"),
     ("1 Apr 2023", "Hanasaari B closed, after replacement heat capacity was secured"),
     ("Apr 2025", "Salmisaari, the last coal plant, closed; national coal ban from May 2029"),
     ("2025 accounts", "Fuel purchases EUR 336m → 216m; allowance costs EUR 78m → 25m (Helen 2026)")],
    "What the records show",
    ["Board judged every option unattractive on purely commercial grounds; if the city acted for climate or land use, it preferred the less burdensome staged pathway.",
     "Capitalisation and plant selection were separate decisions: capital prepared first, pathway chosen later under uncertainty.",
     "Cost recovery tied to municipal finances through limits on heat-price rises and lower expected distributions; Consumer Ombudsman intervened on pricing (KKV 2022)."],
    "Reading",
    [para([R("A politically authorised ", 13, INK), R("modification of the pathway", 13, NAVY, b=True), R(" through which public purposes would be pursued, incorporating the board's conditional preference.", 13, INK)], spcAft=6),
     para([R("The 2025 profitability is a ", 13, INK), R("separate observation", 13, NAVY, b=True), R(": the accounts do not isolate the return on the 2015 choice from price movements.", 13, INK)], spcAft=0)])
N[6] = ("Helsinki is the staging case. Capital was prepared before the pathway was chosen. In December 2015 the board said no option was commercially attractive, "
        "but if the city wanted to act for climate reasons it should take the less burdensome decentralised route, which the council did. Closures followed "
        "only once replacement heat existed. Be careful with the 2025 numbers: falling fuel and allowance costs are reported by management, not attributed to the 2015 decision.")

case_slide(7, "A", GREEN, WHITE,
    "Aalborg: who pays is settled by contract inside the municipal group",
    "Aalborg Forsyning group, boiler agreement, network guarantee and the Vest heat pump (records A1–A6)",
    [("2015–2016", "Municipality acquires Nordjyllandsværket from Vattenfall; parent guarantee for decommissioning"),
     ("May 2017", "Council approves a fossil-free strategy"),
     ("9 Nov 2023", "Vest 20 MW wastewater heat pump abandoned, attributed to the excess-heat price cap"),
     ("4 Dec 2023", "Regulator places wastewater heat outside the cap; the cap itself abolished July 2025"),
     ("26 Jun 2025", "Council guarantee DKK 385m for Varme's network: 17.3% of Varme's year-end assets"),
     ("2025 reports", "150 MW electric boiler assembled; Varme pays fixed capacity and variable costs at self-cost; commissioning summer 2026")],
    "What the records show",
    ["The boiler agreement places the payment obligation on Aalborg Varme, the purchaser, through a supply agreement with the producer company.",
     "The guarantee exposes the municipality to borrowing risk; final risk assessment and guarantee commission are not reported in the published minutes.",
     "Danish heat pricing allows necessary costs into tariffs (hvile-i-sig-selv, §§20–21), subject to regulatory scrutiny."],
    "Reading",
    [para([R("Two interventions at ", 13, INK), R("different corporate levels", 13, NAVY, b=True), R(": a municipal guarantee for one project and an intra-group cost allocation for another.", 13, INK)], spcAft=6),
     para([R("Vest shows a ", 13, INK), R("recovery problem", 13, NAVY, b=True), R(" attributed to a rule as then understood; the later correction shows why a company's appraisal must be read against the rule at the time.", 13, INK)], spcAft=0)])
N[7] = ("Aalborg is the cost-allocation case. The interesting point is where the obligation sits: the council guarantees Varme's borrowing for the network, "
        "while the boiler's costs are passed to Varme as purchaser through a self-cost supply agreement with Nordjyllandsværket. The mayor's 2024 interview "
        "frames green industry as a growth source, but his personal influence on individual investments is not confirmed by the records. Vest is the cautionary tale "
        "about reading a rule as understood at the time.")

case_slide(8, "S", YELLOW, NAVY,
    "Sysav and Malmö: support that keeps its conditions attached",
    "SkyZero carbon capture and the Siptex divestment (records S1–S4)",
    [("Jul 1973", "Consortium agreement proposal links shared finance to dependable waste flows: capital, guarantees, voting, delivery duties"),
     ("Apr 2024", "Board conditions for SkyZero: positive NPV on contracted revenues, sustainable financing, acceptable group risk, owner involvement"),
     ("Autumn 2024", "13 of 14 owner municipalities endorse the direction"),
     ("Dec 2025", "Sysav decides to divest Siptex textile sorting after weak demand for sorted outputs; sale reported July 2026"),
     ("26 Mar 2026", "Malmö authorises development and a guarantee ceiling ≈ SEK 1.2bn (45.98% of SEK 2.6bn)"),
     ("2027", "Final investment decision envisaged; purchase of carbon removals conditional on procurement compliance")],
    "What the records show",
    ["Projected revenues combine carbon-removal credits, avoided allowance costs and anticipated EU Innovation Fund and Industriklivet support; profitability described as dependent on support.",
     "Malmö's prospective purchase is tied to procurement law: in-house award (Teckal, Directive 2014/24 Art. 12) and the January 2026 CJEU judgment C-692/23 on group-level activity.",
     "A later reference (C-672/26) asks whether third-party sales count as entrusted activity; it is a question, not a ruling on SkyZero."],
    "Reading",
    [para([R("Malmö appears twice: as ", 13, INK), R("financier", 13, NAVY, b=True), R(" (guarantee ceiling) and as ", 13, INK), R("prospective customer", 13, NAVY, b=True), R(" (conditional purchase). Ownership did not remove the procurement condition.", 13, INK)], spcAft=6),
     para([R("Siptex is a ", 13, INK), R("change of activities", 13, NAVY, b=True), R(" after a weak market, distinct from abandoning a proposed investment.", 13, INK)], spcAft=0)])
N[8] = ("Sysav is the conditional-support case. Thirteen of fourteen owners back the direction, Malmö sets a guarantee ceiling, yet the business case still "
        "requires contracted revenues and a final decision in 2027. Malmö's dual role matters: borrowing support addresses exposure, while buying the output "
        "requires a procurement-compliant route, which Malmö explicitly linked to the CJEU judgment. Siptex shows revisability at the level of activities.")

# Slide 9 : comparative table ---------------------------------------------
hdr = ["Project (record)", "Owner action", "Financing basis", "Cost-recovery basis", "Stage / open condition"]
rows = [hdr,
    [cell([("Helsinki: decentralised heat and coal closures (H1–H5)", True)]),
     cell(["Incorporation with ≥50% equity target (2014); council chose decentralised option (2 Dec 2015)"]),
     cell(["Prepared equity capacity; EUR 360m vs EUR 960m CHP"]),
     cell(["Board: options commercially unattractive; limited tariff pass-through. 2025: fuel and allowance costs down"]),
     cell(["Pathway authorised 2015; plants closed 2023 and 2025. Price vs production effects not separated"])],
    [cell([("Aalborg: network and associated investments (A4)", True)]),
     cell(["Council guarantee DKK 385m (26 Jun 2025) with commission and risk assessment"]),
     cell(["Municipal guarantee for Varme borrowing; 17.3% of Varme assets"]),
     cell(["Statutory necessary-cost tariff route (§20); project costs' treatment not specified"]),
     cell(["Approved. Final risk assessment and commission not in cited minutes"])],
    [cell([("Aalborg: 150 MW electric boiler (A3)", True)]),
     cell(["No direct owner action documented; intra-group agreement"]),
     cell(["Not documented"]),
     cell(["Varme pays fixed capacity and variable costs at self-cost"]),
     cell(["Cold testing done; commissioning summer 2026. Duration and continuity obligations unspecified"])],
    [cell([("Aalborg: Vest 20 MW heat pump (A5–A6)", True)]),
     cell(["Not documented"]), cell(["Not documented"]),
     cell(["Abandonment attributed to excess-heat cap; regulator corrected 4 Dec 2023; cap abolished Jul 2025"]),
     cell(["Abandonment announced 9 Nov 2023. Later implementation not established"])],
    [cell([("Sysav: SkyZero carbon capture (S2, S4)", True)]),
     cell(["13 of 14 owners endorsed direction; Malmö conditional authorisation and guarantee ceiling ≈ SEK 1.2bn; prospective purchases"]),
     cell(["Internal funds, anticipated Innovation Fund / Industriklivet, borrowing with owner support; conditional"]),
     cell(["Positive NPV on contracted revenues required; Malmö purchase conditional on procurement compliance and final approval"]),
     cell(["Development; FID envisaged 2027. Entrusted-activity treatment pending (C-672/26); grants unconfirmed"])],
    [cell([("Sysav: Siptex textile sorting (S3)", True)]),
     cell(["Not documented"]), cell(["Not documented"]),
     cell(["Weak demand for sorted outputs"]),
     cell(["Divestment Dec 2025; sale reported Jul 2026. Post-sale financing not documented"])],
]
S[9] = title("Owner actions and the components of investability, by project", "Condensed from Table I of the paper") + [
    shape(0.5, 2.38, 12.3, 4.5, None, name="Table backdrop", fill=WHITE, alpha=90),
    table(0.55, 2.42, 12.2, [2.1, 2.6, 2.2, 2.9, 2.5], rows, sz=10.5, row_h=0.62, header_h=0.36),
]
N[9] = ("This is Table I condensed. Read it column by column rather than row by row: the owner-action column shows three different kinds of intervention; "
        "the financing and recovery columns show that support in one column does not settle the other; the last column shows what is still open. "
        "Point out the 'not documented' cells and repeat that they are limits of the record.")

# Slide 10 : finding 1 ------------------------------------------------------
S[10] = title("Finding 1: owner support coexisted with adverse appraisals and open recovery conditions") + [
    text(0.55, 2.05, 12.2, 0.5, [para([R("Financial support was given, yet recovery stayed under scrutiny in every arrangement. The three records show the same gap in three forms.", 14, GREY, i=True)], spcAft=0)], name="Lead"),
]
f1 = [
    ("H", NAVY, WHITE, "Helsinki", "Capital prepared", "Board: all alternatives commercially unattractive",
     "Council selected the board's conditional preference: a less burdensome, staged pathway. Political authorisation adjusted the means, not the purpose."),
    ("A", GREEN, WHITE, "Aalborg", "Guarantee DKK 385m for network", "Boiler costs assigned to the purchaser by agreement",
     "The group company accepted the producer's costs through a supply agreement. Support and payment obligations sit at different corporate levels."),
    ("S", YELLOW, NAVY, "Sysav / Malmö", "Guarantee ceiling ≈ SEK 1.2bn", "Contracted revenues and final approval still required",
     "Malmö qualified its support: borrowing ceiling now, purchase subject to procurement compliance, decision in 2027."),
]
for k, (ltr, col, tcol, nm, sup, cond, body) in enumerate(f1):
    x = 0.55 + k * (cw + gap)
    S[10] += [
        card(x, 2.65, cw, 3.4, [
            para([R(nm, 15, NAVY, b=True)], spcAft=8),
            para([R("Owner support", 11, SAGE, b=True)], spcAft=1),
            para([R(sup, 13.5, INK, b=True)], spcAft=8),
            para([R("Still outstanding / adverse", 11, SAGE, b=True)], spcAft=1),
            para([R(cond, 13.5, INK, b=True)], spcAft=10),
            para([R(body, 12.5, INK)], spcAft=0),
        ], inset=(0.2, 0.2, 0.2, 0.15)),
        badge(x + cw - 0.75, 2.8, 0.55, ltr, fill=col, color=tcol, sz=18),
    ]
S[10] += [card(0.55, 6.2, 12.2, 0.6, [para([R("Owner support advances a proposal; it does not by itself settle the business case.", 14.5, NAVY, b=True)], algn="ctr", spcAft=0)], fill="EEF3F8", alpha=96, anchor="ctr", shadow=False)]
N[10] = ("The central finding. In each case the owner did something financial, and in each case the recovery side was not thereby settled. Helsinki: capital "
         "prepared, board adverse. Aalborg: guarantee for the network, costs of the boiler assigned by contract. Malmö: guarantee ceiling, revenues and approval "
         "outstanding. Owner support advances a proposal; it does not make the business case.")

# Slide 11 : finding 2 ------------------------------------------------------
S[11] = title("Finding 2: institutions entered through appraisals and purchasing conditions")
chan = [
    ("Tariff and price rules", NAVY,
     ["Helen: limits on heat-price increases in the 2015 appraisal; Consumer Ombudsman required changes to price-rise procedures (KKV 2022).",
      "Denmark: necessary costs may enter tariffs (§§20–21); Vest abandoned on the excess-heat cap as then understood, corrected a month later."],
     "Appraised tariff capacity"),
    ("Supply agreements between companies", GREEN,
     ["Aalborg Varme covers Nordjyllandsværket's fixed capacity and variable operating costs at self-cost for the 150 MW boiler.",
      "Duration and continuity obligations are not specified in the cited reports; the agreement identifies the payer, not the full recovery."],
     "Contractual cost allocation"),
    ("Procurement rules", SAGE,
     ["Malmö's purchase of carbon removals conditional on procurement compliance, final investment approval and legally effective decisions.",
      "In-house award (Teckal) requires >80% entrusted activity; Malmö linked its condition to CJEU C-692/23 (Jan 2026)."],
     "Conditional purchasing"),
]
for k, (hd, col, items, tag) in enumerate(chan):
    x = 0.55 + k * (cw + gap)
    S[11] += [
        card(x, 2.15, cw, 3.55, [
            para([R(hd, 15, col, b=True)], spcAft=8),
            *bullets(items, sz=12.5, spcAft=6, bcolor=col),
        ], inset=(0.2, 0.2, 0.2, 0.15)),
        shape(x + 0.2, 5.25, cw - 0.4, 0.34, [para([R(tag, 11.5, WHITE, b=True)], algn="ctr", spcAft=0)], name="Tag", geom="roundRect", fill=col, anchor="ctr", radius=0.5),
    ]
S[11] += [
    card(0.55, 5.9, 12.2, 0.9, [
        para([R("Caveat: ", 14, NAVY, b=True),
              R("the presence of a rule does not establish that an investment's full costs will be recovered. The evidence of institutional influence lies in the specific assessment or condition, read against the rule as understood at the time.", 13.5, INK)], spcAft=0)],
        fill="EEF3F8", alpha=96, anchor="ctr", shadow=False),
]
N[11] = ("Second finding: institutions matter, but through concrete channels in specific decisions, not as 'Finland versus Denmark versus Sweden'. Tariff rules "
         "shaped Helen's appraisal and Aalborg Forsyning's Vest decision; a supply agreement allocated the boiler's costs; procurement law shaped how Malmö "
         "may buy output. Stress the caveat about rules as understood at the time, using Vest.")

# Slide 12 : contribution ---------------------------------------------------
S[12] = title("Contribution and implications") + [
    card(0.55, 2.15, 6.0, 2.2, [
        para([R("To the multiple-principals literature", 14, NAVY, b=True)], spcAft=6),
        para([R("Agreement on direction can coexist with different positions on the obligations needed to act (refining Voorn et al. 2019; Voorn & van Genugten 2022). Agreement among public actors must be examined at the level of expenditure and income, not only objectives. Even within one municipal group, common ownership leaves the payment obligation to be specified by contract.", 12.5, INK)], spcAft=0),
    ], inset=(0.2, 0.16, 0.2, 0.12)),
    card(6.75, 2.15, 6.0, 2.2, [
        para([R("To public service ecosystem research", 14, GREEN, b=True)], spcAft=6),
        para([R("Mediation of public values here takes a material and financial form: which production option a council accepts, which company pays for heat, on what terms a municipality may buy carbon removals. Agreement on an environmental purpose coexists with disagreement about acceptable cost. Service continuity (replacement heat before closure) is a condition of the transition pathway.", 12.5, INK)], spcAft=0),
    ], inset=(0.2, 0.16, 0.2, 0.12)),
    card(0.55, 4.55, 6.0, 2.25, [
        para([R("On investability", 14, SAGE, b=True)], spcAft=6),
        para([R("Its value lies less in classifying projects as investable or not than in separating the commitments that make investment possible. Owner support did not settle recovery in Helsinki or SkyZero; Aalborg's agreement shows a purchaser taking responsibility for production costs. Incomplete records do not license the inference that municipal ownership relaxes financing constraints.", 12.5, INK)], spcAft=0),
    ], inset=(0.2, 0.16, 0.2, 0.12)),
    card(6.75, 4.55, 6.0, 2.25, [
        para([R("Managerial implication: make the scope of a commitment explicit", 14, YELLOW, b=True)], spcAft=6),
        *bullets(["the alternatives considered", "the party accepting each cost", "the revenue conditions", "the decisions still outstanding"], sz=12.5, spcAft=2, bcolor=YELLOW),
        para([R("This allows scrutiny of exposure while a common purpose stays politically supported. Where rules affect recovery or purchasing, their interpretation belongs in the appraisal.", 12.5, INK)], spcAft=0, spcBef=4),
    ], fill="EEF3F8", alpha=96, inset=(0.2, 0.16, 0.2, 0.12)),
]
N[12] = ("Four takeaways. For multiple-principals research: agreement on direction is not agreement on obligations. For public service ecosystem research: "
         "the mediation is material, it concerns plants, contracts and prices. For the concept of investability: use it to separate commitments, not to score projects. "
         "For managers and councils: write down the alternatives, who pays each cost, the revenue conditions and the open decisions.")

# Slide 13 : limitations and next steps -------------------------------------
S[13] = title("Limitations and next steps") + [
    card(0.55, 2.15, 6.0, 4.3, [
        para([R("Limits of the evidence", 14, NAVY, b=True)], spcAft=6),
        *bullets([
            "Documentary accounts capture formal decisions and reported explanations more readily than informal negotiation.",
            "Evidence is uneven across projects and stages; several financing or recovery cells remain blank.",
            "National effects are not isolated; the comparison identifies institutional channels in specific decisions, not country outcomes.",
            "Reported operating costs (Helen 2025) need a different assessment from the return attributable to an earlier choice.",
            "Revisability claims (Helsinki staging, Siptex sale) would need revised costs, revenues and accepted exposure from the same project.",
        ], sz=12.5, spcAft=6),
    ], inset=(0.2, 0.16, 0.2, 0.12)),
    card(6.75, 2.15, 6.0, 4.3, [
        para([R("Next steps", 14, GREEN, b=True)], spcAft=6),
        *bullets([
            "Follow the same projects through implementation: boiler commissioning (summer 2026), SkyZero final decision (2027), post-closure accounts at Helen.",
            "Add the outstanding documents: Aalborg's final risk assessment and guarantee commission; the Varme–Nordjyllandsværket agreement's duration and continuity terms.",
            "Trace whether revisions preserved approved financing and credible recovery, and how accepted obligations changed during operation.",
            "Complement records with interviews on informal negotiation between councils, boards and group management.",
            "Test the framework on further arrangements (e.g. Finnish and German multi-utility groups) before any generalisation.",
        ], sz=12.5, spcAft=6, bcolor=GREEN),
    ], inset=(0.2, 0.16, 0.2, 0.12)),
]
N[13] = ("Be upfront about what documents cannot show. Then the follow-up plan: the three projects have dated milestones in 2026 and 2027, which makes a "
         "second wave of observation feasible. Invite suggestions on comparable arrangements and on access to board-level material.")

# Slide 14 : closing ---------------------------------------------------------
S[14] = title("Central finding") + [
    card(0.55, 2.15, 12.2, 1.95, [
        para([R("Owner support can advance a politically endorsed proposal while commercial appraisal or revenue conditions remain adverse or unsettled.", 19, NAVY, b=True)], spcAft=8, lnSpc=1.05),
        para([R("Public ownership matters through the authority and organisational relationships used to negotiate these obligations. Its contribution to green transition lies in coordinating acceptable expenditure, financing and recovery, with choices about staging, contracts and conditional participation. Financing intervention alone cannot be treated as a settled business case.", 13.5, INK)], spcAft=0),
    ], inset=(0.25, 0.2, 0.25, 0.18), anchor="ctr"),
    card(0.55, 4.3, 5.1, 2.5, [
        para([R("Thank you", 20, GREEN, b=True)], spcAft=8),
        para([R("[Presenter name]", 14, NAVY, b=True)], spcAft=2),
        para([R("[Affiliation]", 13, INK)], spcAft=2),
        para([R("[e-mail address]", 13, INK)], spcAft=10),
        para([R("Full paper, Table I and Appendix A (15 records) available on request.", 11.5, GREY, i=True)], spcAft=0),
    ], inset=(0.25, 0.2, 0.25, 0.15)),
    card(5.85, 4.3, 6.9, 2.5, [
        para([R("Selected references", 13, SAGE, b=True)], spcAft=4),
        para([R("Bruton, G. D. et al. (2015). State-owned enterprises around the world as hybrid organizations. Academy of Management Perspectives, 29(1).", 10.5, INK)], spcAft=3),
        para([R("Osborne, S. P. et al. (2022). Value creation in the public service ecosystem. Public Administration Review, 82(4).", 10.5, INK)], spcAft=3),
        para([R("Skålén, P., Engen, M., & Jenhaug, L. (2024). Conflicts over public value within public service ecosystems. Public Management Review, 26(11).", 10.5, INK)], spcAft=3),
        para([R("Voorn, B., & van Genugten, M. (2022). Dealing with multiple principals in at arm's length organizations. Public Administration, 100(2).", 10.5, INK)], spcAft=3),
        para([R("Primary records: Helsinki (2014, 2015); Helen (2023, 2025, 2026); Aalborg Kommune (2025a, b); Aalborg Varme A/S (2026); Forsyningstilsynet (2023); Sysav (2026a, b); Malmö stad (2026a, b); CJEU C-692/23, C-672/26.", 10.5, GREY)], spcAft=0),
    ], fill="EEF3F8", alpha=96, inset=(0.22, 0.16, 0.22, 0.12)),
]
N[14] = ("Close with the one-sentence finding and the policy point: financing support is not a settled business case. Thank the audience and invite questions, "
         "particularly from colleagues working on municipal companies in other countries.")

# ---------- write into the unpacked package ----------
def inject(slide_path, shapes_xml):
    x = open(slide_path, encoding="utf-8").read()
    x = x.replace("</p:spTree>", "".join(shapes_xml) + "</p:spTree>", 1)
    open(slide_path, "w", encoding="utf-8").write(x)

def notes_xml(text_, slide_rel="rId1"):
    # minimal notes slide
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
            '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>{escape(text_)}</a:t></a:r></a:p></p:txBody></p:sp>'
            '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>')

ct_path = os.path.join(UNP, "[Content_Types].xml")
ct = open(ct_path, encoding="utf-8").read()
notes_master_rel = None
# find notes master target from an existing notes slide rels
ex = open(os.path.join(UNP, "ppt/notesSlides/_rels/notesSlide1.xml.rels"), encoding="utf-8").read()
nm = re.search(r'Target="(\.\./notesMasters/notesMaster\d+\.xml)"', ex).group(1)

for n in sorted(S):
    sp = os.path.join(UNP, f"ppt/slides/slide{n}.xml")
    inject(sp, S[n])
    # notes
    rels_p = os.path.join(UNP, f"ppt/slides/_rels/slide{n}.xml.rels")
    rels = open(rels_p, encoding="utf-8").read()
    m = re.search(r'Target="\.\./notesSlides/(notesSlide\d+\.xml)"', rels)
    if m:
        np_ = os.path.join(UNP, "ppt/notesSlides", m.group(1))
        # replace notes text in existing notes slide: rewrite the file wholesale
        open(np_, "w", encoding="utf-8").write(notes_xml(N[n]))
    else:
        fname = f"notesSlide{n}.xml"
        np_ = os.path.join(UNP, "ppt/notesSlides", fname)
        open(np_, "w", encoding="utf-8").write(notes_xml(N[n]))
        open(os.path.join(UNP, "ppt/notesSlides/_rels", fname + ".rels"), "w", encoding="utf-8").write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="{nm}"/>'
            f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide{n}.xml"/></Relationships>')
        ids = [int(v) for v in re.findall(r'Id="rId(\d+)"', rels)]
        new = f'<Relationship Id="rId{max(ids)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/{fname}"/>'
        rels = rels.replace("</Relationships>", new + "</Relationships>")
        open(rels_p, "w", encoding="utf-8").write(rels)
        if f"/ppt/notesSlides/{fname}" not in ct:
            ct = ct.replace("</Types>", f'<Override PartName="/ppt/notesSlides/{fname}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>')
open(ct_path, "w", encoding="utf-8").write(ct)
print("filled", len(S), "slides")
