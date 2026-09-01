// Builds the Word versions of the two Malmö documents.
//   node build-docx.js
// Produces japan-transition-report.docx and kitakyushu-briefing.docx in this directory.

const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageBreak, PageOrientation, Header, Footer, PageNumber,
  LevelFormat, convertMillimetersToTwip,
} = require('docx');

// ---------------------------------------------------------------- constants

const ACCENT = '1D4E5F';
const ACCENT2 = '7A4A17';
const MUTED = '5B5B5B';
const RULE = 'C9C9C9';
const SOFT = 'EEF3F4';
const WARM = 'FAF4EA';
const GREY = 'F4F4F2';

const MARGIN = { top: convertMillimetersToTwip(20), bottom: convertMillimetersToTwip(18),
                 left: convertMillimetersToTwip(18), right: convertMillimetersToTwip(18) };
const W = 11906 - MARGIN.left - MARGIN.right;   // A4 width minus margins, in DXA

const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const hair = { style: BorderStyle.SINGLE, size: 4, color: RULE };
const solidRule = { style: BorderStyle.SINGLE, size: 8, color: '1A1A1A' };

// ---------------------------------------------------------------- inline markup

// **bold**, *italic*. Everything else is literal.
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ ...base, text: text.slice(last, m.index) }));
    const t = m[0];
    if (t.startsWith('**')) out.push(new TextRun({ ...base, text: t.slice(2, -2), bold: true }));
    else out.push(new TextRun({ ...base, text: t.slice(1, -1), italics: true }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ ...base, text: text.slice(last) }));
  return out;
}

// ---------------------------------------------------------------- blocks

const P = (text, o = {}) => new Paragraph({
  children: runs(text, o.run || {}),
  spacing: { after: o.after === undefined ? 120 : o.after, line: 264 },
  alignment: o.align,
  indent: o.indent,
  keepNext: o.keepNext,
  border: o.border,
});

const H1 = (text, sub) => [
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { after: sub ? 80 : 200 },
    children: [new TextRun({ text })],
  }),
  ...(sub ? [new Paragraph({
    spacing: { after: 240 },
    children: [new TextRun({ text: sub, size: 24, color: MUTED, font: 'Calibri' })],
  })] : []),
];

const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 320, after: 120 },
  keepNext: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 3 } },
  children: [new TextRun({ text })],
});

const H3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 200, after: 80 },
  keepNext: true,
  children: [new TextRun({ text })],
});

const H4 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_4,
  spacing: { before: 160, after: 60 },
  keepNext: true,
  children: [new TextRun({ text: text.toUpperCase(), characterSpacing: 12 })],
});

const KICKER = (text) => new Paragraph({
  spacing: { after: 120 },
  children: [new TextRun({
    text: text.toUpperCase(), font: 'Calibri', size: 15, bold: true,
    color: ACCENT, characterSpacing: 40,
  })],
});

const bullet = (text, level = 0) => new Paragraph({
  numbering: { reference: 'bullets', level },
  spacing: { after: 60, line: 264 },
  children: runs(text),
});

const numbered = (text) => new Paragraph({
  numbering: { reference: 'numbers', level: 0 },
  spacing: { after: 80, line: 264 },
  children: runs(text),
});

const srcItem = (text) => new Paragraph({
  numbering: { reference: 'sources', level: 0 },
  spacing: { after: 70, line: 240 },
  children: runs(text, { size: 17 }),
});

// A shaded single-cell table used for callouts.
function shadedBox(children, fill, barColor) {
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [W],
    borders: {
      top: noBorder, bottom: noBorder, right: noBorder,
      insideHorizontal: noBorder, insideVertical: noBorder,
      left: { style: BorderStyle.SINGLE, size: 18, color: barColor },
    },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: W, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill, color: 'auto' },
        margins: { top: 140, bottom: 100, left: 180, right: 180 },
        children,
      })],
    })],
  });
}

const box = (title, body, warm = false) => shadedBox(
  [
    new Paragraph({
      spacing: { after: 80 }, keepNext: true,
      children: [new TextRun({
        text: title.toUpperCase(), font: 'Calibri', size: 17, bold: true,
        color: warm ? ACCENT2 : ACCENT, characterSpacing: 12,
      })],
    }),
    ...body.map((t) => P(t, { after: 60 })),
  ],
  warm ? WARM : SOFT,
  warm ? ACCENT2 : ACCENT,
);

const keyfig = (text) => shadedBox(
  [new Paragraph({ spacing: { after: 0, line: 252 }, children: runs(text, { font: 'Calibri', size: 17 }) })],
  GREY, GREY,
);

// Table. `widths` are DXA and must sum to W. Cells are markup strings.
function tbl(widths, rows, opts = {}) {
  const { header = false, size = 18 } = opts;
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: widths,
    borders: {
      top: noBorder, bottom: hair, left: noBorder, right: noBorder,
      insideHorizontal: hair, insideVertical: noBorder,
    },
    rows: rows.map((cells, r) => new TableRow({
      tableHeader: header && r === 0,
      children: cells.map((c, i) => {
        const isHead = header && r === 0;
        return new TableCell({
          width: { size: widths[i], type: WidthType.DXA },
          margins: { top: 90, bottom: 90, left: 0, right: 160 },
          borders: isHead ? { bottom: solidRule } : {},
          children: [new Paragraph({
            spacing: { after: 0, line: 252 },
            children: runs(String(c), isHead
              ? { font: 'Calibri', size: 15, color: MUTED, characterSpacing: 8 }
              : { size }),
          })],
        });
      }),
    })),
  });
}

// Two-column "label / value" table, used for the fact lists.
const facts = (rows, labelW = 3000) => tbl([labelW, W - labelW],
  rows.map(([k, v]) => [`**${k}**`, v]));

const flow = (lines) => new Table({
  width: { size: W, type: WidthType.DXA },
  columnWidths: [W],
  borders: {
    top: noBorder, bottom: noBorder, right: noBorder,
    insideHorizontal: noBorder, insideVertical: noBorder,
    left: { style: BorderStyle.SINGLE, size: 6, color: RULE },
  },
  rows: [new TableRow({
    children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      margins: { top: 100, bottom: 100, left: 180, right: 0 },
      children: lines.map((l, i) => new Paragraph({
        spacing: { after: i === lines.length - 1 ? 0 : 40, line: 264 },
        children: runs(l, { font: 'Calibri', size: 17 }),
      })),
    })],
  })],
});

const caveat = (text) => new Paragraph({
  spacing: { before: 100, after: 120, line: 252 },
  border: { top: { style: BorderStyle.DOTTED, size: 4, color: RULE, space: 6 } },
  children: runs(text, { italics: true, size: 17, color: MUTED }),
});

const partHead = (part, title, strap) => [
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({
    spacing: { before: 0, after: 100 },
    border: { top: { style: BorderStyle.SINGLE, size: 18, color: ACCENT, space: 6 } },
    children: [new TextRun({
      text: part.toUpperCase(), font: 'Calibri', size: 15, bold: true,
      color: ACCENT, characterSpacing: 40,
    })],
  }),
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: title, size: 30 })],
  }),
  P(strap, { run: { color: MUTED, size: 19 }, after: 200 }),
];

const caseHead = (label, title, strap) => [
  new Paragraph({
    spacing: { before: 280, after: 60 }, keepNext: true,
    border: { top: { style: BorderStyle.SINGLE, size: 10, color: '1A1A1A', space: 5 } },
    children: [new TextRun({
      text: label.toUpperCase(), font: 'Calibri', size: 15, bold: true,
      color: ACCENT2, characterSpacing: 24,
    })],
  }),
  new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 0, after: 50 }, keepNext: true,
    children: [new TextRun({ text: title, size: 24 })],
  }),
  P(strap, { run: { color: MUTED, size: 19 }, after: 140, keepNext: true }),
];

const metaBlock = (rows) => new Table({
  width: { size: W, type: WidthType.DXA },
  columnWidths: [2400, W - 2400],
  borders: {
    top: hair, bottom: hair, left: noBorder, right: noBorder,
    insideHorizontal: noBorder, insideVertical: noBorder,
  },
  rows: rows.map(([k, v]) => new TableRow({
    children: [
      new TableCell({
        width: { size: 2400, type: WidthType.DXA },
        margins: { top: 70, bottom: 70, left: 0, right: 120 },
        children: [new Paragraph({
          spacing: { after: 0 },
          children: [new TextRun({ text: k, font: 'Calibri', size: 17, bold: true })],
        })],
      }),
      new TableCell({
        width: { size: W - 2400, type: WidthType.DXA },
        margins: { top: 70, bottom: 70, left: 0, right: 0 },
        children: [new Paragraph({
          spacing: { after: 0, line: 252 },
          children: runs(v, { font: 'Calibri', size: 17, color: MUTED }),
        })],
      }),
    ],
  })),
});

const NAME_PLACEHOLDER =
  '[name], [affiliation]';

// ---------------------------------------------------------------- doc scaffold

function makeDoc(title, sections) {
  return new Document({
    title,
    creator: 'Nordic field study 2026',
    description: title,
    styles: {
      default: {
        document: { run: { font: 'Cambria', size: 21, color: '1A1A1A' },
                    paragraph: { spacing: { line: 264 } } },
        heading1: { run: { font: 'Calibri', size: 36, bold: true, color: '1A1A1A' },
                    paragraph: { spacing: { before: 200, after: 120 } } },
        heading2: { run: { font: 'Calibri', size: 25, bold: true, color: ACCENT },
                    paragraph: { spacing: { before: 320, after: 120 } } },
        heading3: { run: { font: 'Calibri', size: 21, bold: true, color: '1A1A1A' },
                    paragraph: { spacing: { before: 200, after: 80 } } },
        heading4: { run: { font: 'Calibri', size: 17, bold: true, color: MUTED },
                    paragraph: { spacing: { before: 160, after: 60 } } },
      },
    },
    numbering: {
      config: [
        { reference: 'bullets', levels: [
          { level: 0, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 360, hanging: 200 } } } },
          { level: 1, format: LevelFormat.BULLET, text: '·', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 200 } } } },
        ] },
        { reference: 'numbers', levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 400, hanging: 260 } } } },
        ] },
        { reference: 'sources', levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 400, hanging: 260 } } } },
        ] },
      ],
    },
    sections: [{
      properties: {
        page: { size: { orientation: PageOrientation.PORTRAIT }, margin: MARGIN },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 120 },
            children: [new TextRun({ children: [PageNumber.CURRENT], font: 'Calibri', size: 16, color: MUTED })],
          })],
        }),
      },
      children: sections,
    }],
  });
}

// ================================================================ REPORT

const report = [];
const R = (...x) => report.push(...x.flat());

R(KICKER('Working report · for the Malmö dialogue'));
R(H1('Financing the transition',
  'Nine Japanese cases in industrial decarbonisation, and what they reveal about who owns, '
  + 'who pays and who carries the risk — with questions for the City of Malmö'));

R(metaBlock([
  ['Prepared for', 'Per-Arne Nilsson, Senior Strategist, Environmental Department, City of Malmö'],
  ['Meeting', 'Malmö, 9 September 2026'],
  ['Prepared by', NAME_PLACEHOLDER],
  ['Evidence base', 'A study visit to Kitakyushu on 22 September 2023 by the Public Finance Study '
    + 'Group, Kyoto University, together with the city and company documents listed in Part V; and '
    + "the author's own research on Kitakyushu's industrial structure"],
  ['Context', 'Nordic field study, 7–11 September 2026 — Danish Energy Agency · Copenhagen Business '
    + 'School · Malmö · Lolland'],
]));
R(P('', { after: 200 }));

R(H2('Executive summary'));
R(P('Kitakyushu is Japan’s clearest example of a heavy-industry city that rebuilt its identity '
  + 'around the environment. It is also, on the economic evidence, a warning. This report sets out '
  + 'nine cases drawn from the city and its region, examines each for its *ownership and financing '
  + 'structure* rather than its targets, and arrives at a single argument I would like to test '
  + 'against Malmö’s experience.'));
R(P('**The argument.** Japanese municipalities have weak fiscal autonomy and, since electricity '
  + 'liberalisation, nothing resembling a Nordic municipal utility. Kitakyushu has therefore solved '
  + 'decarbonisation by *buying services rather than owning assets* — third-party ownership of solar, '
  + 'batteries and even air conditioning on public buildings, under a rule that the resulting tariff '
  + 'must not exceed the electricity price the city already pays. It works: the rollout is fast, the '
  + 'municipal balance sheet is untouched, and 626 public buildings across eighteen municipalities '
  + 'were on renewable electricity by mid-2023. But the city ends up with contracts rather than '
  + 'capital, and the pattern repeats one level up: the region’s flagship offshore wind farm is '
  + 'owned by a consortium of five national utilities and contractors, in which the municipality’s '
  + 'role was to run the tender and provide the port.'));
R(P('**The counter-evidence.** Two cases cut against that reading. The reuse-panel demonstration at '
  + 'the Hibikinada Biotope shows the same third-party model being extended to second-life assets, '
  + 'with a leasing company inside the structure — the beginnings of a financing industry rather than '
  + 'a procurement trick. And Gotō, a small island city in Nagasaki, has produced something Kitakyushu '
  + 'has not: a locally-owned maintenance firm, a citizens’ electricity retailer, and a training '
  + 'association that deliberately creates competitors. Value is retained locally there in a way that '
  + '940,000-person Kitakyushu has never achieved.'));
R(P('**The long record.** Kitakyushu has run an environment-and-knowledge-led growth strategy for '
  + 'three decades through its Science and Research Park and the agency FAIS. Gross product has '
  + 'nonetheless grown slowly; what actually expanded was health care, real estate and public '
  + 'services, not the targeted high-value manufacturing. My tentative explanation is that the '
  + 'programme’s success criterion was commercialisation rather than value added, that no private '
  + 'capital layer ever formed, and that the business services a technology cluster needs never '
  + 'agglomerated — so the city substituted for them itself, and thereby capped the scale at what a '
  + 'shrinking municipal budget can fund.'));
R(P('**What I want from Malmö.** Malmö took the opposite route on almost every one of these choices. '
  + 'Part IV puts fourteen questions, grouped by theme. The single one I care most about: does a city '
  + 'that keeps ownership end up with more capability than a city that buys services — or only with '
  + 'more debt?'));

R(H2('Contents'));
R(tbl([1100, W - 1100], [
  ['Part I', '**Framing** — the question, the city, the four phases'],
  ['Part II', '**Nine cases**, each read for ownership, financing and risk'],
  ['', '1  The municipal estate: third-party ownership at scale'],
  ['', '2  Second-life assets: the reuse-panel demonstration'],
  ['', '3  Eco-Town: circular industry as an export'],
  ['', '4  Hibiki offshore wind farm: a municipally-tendered consortium'],
  ['', '5  Green Energy Port Hibiki: the port as industrial policy'],
  ['', '6  Floating wind capability: a state-subsidised vessel'],
  ['', '7  Ørsted in Japan: the ownership model Japan does not have'],
  ['', '8  Gotō: the community counter-case'],
  ['', '9  FAIS and the Science Park: thirty years, weak result'],
  ['Part III', '**Synthesis** — comparison, five findings, the missing layers'],
  ['Part IV', '**Fourteen questions** for the City of Malmö'],
  ['Part V', '**Sources and method**'],
], { size: 18 }));

// ---- Part I
R(partHead('Part I', 'Framing', 'Why this comparison, what Kitakyushu is, and how it got here.'));

R(H2('1  The question'));
R(P('Malmö and Kitakyushu are not obvious twins, but their problem was the same. Each was a second '
  + 'city built on a single industry — shipbuilding in Malmö, integrated steelmaking in Kitakyushu — '
  + 'and each lost the certainty that industry provided. Both answered by redefining themselves as '
  + 'environmental cities. Both are now judged on whether that redefinition produced an economy or '
  + 'only a reputation.'));
R(P('I have studied Kitakyushu since my undergraduate work, and my interest is deliberately narrow. '
  + 'I am not asking what the city declares; declarations are cheap and Kitakyushu has made many good '
  + 'ones. I am asking three questions of every project: **who owns the asset, who supplies the '
  + 'capital, and who is left holding the risk.** Those three answers determine whether a transition '
  + 'builds durable municipal capability or merely rents it.'));
R(P('The reason to ask them of Japan specifically is that the Japanese municipal toolkit is unusually '
  + 'thin. Local governments have limited own-source revenue and limited borrowing discretion. Since '
  + 'electricity liberalisation there is no municipal utility in the Nordic sense — Kitakyushu holds a '
  + 'stake in a local retailer, Kitakyushu Power, but not in generation, networks or heat. What a '
  + 'Japanese city has to work with is its buildings, its land, its port, its convening power, and its '
  + 'procurement. Everything in Part II is a variation on making those five things do the work that a '
  + 'balance sheet does elsewhere.'));

R(H2('2  Kitakyushu at a glance'));
R(facts([
  ['Established', '1963, by the merger of five cities of equal standing (Moji, Kokura, Wakamatsu, '
    + 'Yahata, Tobata) — an unusual origin that left the city polycentric and without a single '
    + 'dominant centre'],
  ['Area', '491.71 km²'],
  ['Population', 'c. 940,000 (January 2022); c. 490,000 households. A designated city in long-term '
    + 'decline from a peak above 1.06 million in 1979'],
  ['Principal industries', 'Iron and steel, chemicals, cement, ceramics, automotive'],
  ['Greenhouse gas emissions', '14.78 Mt CO₂e (2019)'],
  ['Industrial share of emissions', '**64%** — against 32% for Japan as a whole (national total '
    + '1.21 Gt CO₂e, 2019)'],
  ['Headline target', '−47% GHG by FY2030 against FY2013; net zero by 2050'],
  ['Installed renewables (city)', 'Wind 24,852 kW and solar 255,757 kW at end-March 2017; target '
    + '1,302–1,402 MW by FY2030'],
]));
R(P('That 64% is the most important number in this report, and it recurs in every case that follows. '
  + 'Two-thirds of Kitakyushu’s emissions sit inside a handful of private industrial installations '
  + 'that the city does not own, cannot regulate directly, and cannot afford to lose. Every instrument '
  + 'in Part II is shaped by that fact: the city works intensively on the third of emissions it can '
  + 'reach, and tries to influence the rest through industrial policy rather than climate policy. When '
  + 'Kitakyushu talks about decarbonisation it is, in a precise sense, talking about industrial '
  + 'strategy.'));

R(H2('3  Four phases'));
R(tbl([1500, 2300, W - 3800], [
  ['Period', 'Phase', 'What happened'],
  ['1901–1963', '**Building the steel city**',
    'The state-owned Yawata Steel Works opens in 1901 and becomes the core of Japanese '
    + 'industrialisation. Five towns grow together around it and merge in 1963.'],
  ['1960s–1970s', '**Pollution, and the citizens’ response**',
    'Dokai Bay is declared a "sea of death"; the sky over Yahata is described as having seven '
    + 'colours. Local women’s associations document the damage and force the issue onto the '
    + 'political agenda, producing pollution-control agreements between the city and industry ahead '
    + 'of national law. The environmental identity was imposed from below, not chosen.'],
  ['1980s–2000s', '**Environment as an industry**',
    'The clean-up is converted into an export. The Eco-Town project (from 1997) turns the Hibikinada '
    + 'waterfront into a recycling industrial estate; the Science and Research Park and its agency '
    + 'FAIS attempt a knowledge economy; the city becomes a hub for environmental cooperation with '
    + 'Asian cities. Eco-Model City (2008), SDGs Future City (2018).'],
  ['2010s–', '**Energy as an industry**',
    'Green Energy Port Hibiki (from 2011) reorients the port towards offshore wind. Zero Carbon City '
    + 'declaration (October 2020); climate emergency declaration (June 2021); revised climate plan '
    + '(August 2021); Green Growth Strategy (February 2022); selection as a national *Decarbonisation '
    + 'Leading Area* (April 2022).'],
], { header: true }));

R(box('The national policy backdrop, compressed', [
  'The Japanese offshore wind sector is very young, and the Kitakyushu cases only make sense against '
  + 'it: feed-in tariff introduced 2012; Port and Harbour Act amended 2016 to allow competitive '
  + 'tendering of port water areas; 5th Basic Energy Plan 2018 (wind at 1.7%); the *Renewable Sea '
  + 'Area Utilisation Act* in force 2019 — widely called Japan’s "offshore wind year one"; July '
  + '2020, a policy to retire around 90% of inefficient coal capacity; September 2020, four base ports '
  + 'designated (Kitakyushu, Akita, Noshiro, Kashima); December 2020, a national target of up to 45 GW '
  + 'of offshore wind by 2040; 2020 Port Act amendment extending port occupancy rights from 20 to 30 '
  + 'years. Kitakyushu began its port programme in 2011, ahead of all of it.',
]));

// ---- Part II
R(partHead('Part II', 'Nine cases',
  'Each case is read the same way: what it is, the numbers, the ownership and financing structure, '
  + 'what it demonstrates, and what remains unresolved.'));

// Case 1
R(caseHead('Case 1 · the municipal estate', 'Third-party ownership at scale',
  'Decarbonisation Leading Area and the "100% Renewable Kitakyushu Model" — how a city with no '
  + 'capital decarbonises 3,600 buildings.'));
R(H4('What it is'));
R(P('In April 2022 Kitakyushu was selected in the first cohort of the national *Decarbonisation '
  + 'Leading Area* programme, which requires net-zero electricity in the residential and commercial '
  + 'sector by FY2030 and is intended to produce a replicable model — the government’s phrase is a '
  + '"decarbonisation domino". Kitakyushu applied jointly with **17 neighbouring municipalities**, '
  + 'eighteen in total.'));
R(keyfig('**Scope** approximately **3,600 public facilities** across the eighteen municipalities, plus '
  + 'the recycling companies of Eco-Town  ·  **New renewable capacity** 37 MW  ·  **PV sites** c. 290 '
  + 'in Kitakyushu, c. 130 in the 17 neighbours  ·  **Renewable electricity share** 100% for the '
  + 'target estate  ·  **Funding** the national Regional Decarbonisation Transition and Renewable '
  + 'Energy Promotion Grant'));
R(H4('The structure'));
R(P('The city does not buy solar panels. It buys electricity from panels someone else owns that '
  + 'happen to sit on its roofs. Its own slogan is a shift **"from ownership to use"**.'));
R(flow([
  '**Equipment company** owns and installs the PV, battery or air-conditioning unit',
  '→ **JV / consortium** operates and manages it, with IoT monitoring',
  '→ **Retailer** (Kitakyushu Power and others) holds the electricity supply contract',
  '→ **Public or private facility** pays an electricity tariff only — no capital expenditure',
]));
R(P('Above this sits a three-step ladder that any building can be moved along. **Step 1** is 100% '
  + 'renewable electricity supplied from generation *inside the city*, including waste-to-energy '
  + 'plants, through Kitakyushu Power. **Step 2** adds PV plus a battery under third-party ownership, '
  + 'charging when power is cheap and discharging when it is expensive. **Step 3** adds '
  + 'energy-efficient equipment, also third-party owned, reducing total consumption and therefore the '
  + 'volume of renewable electricity that must be procured.'));
R(H4('Results to mid-2023'));
R(tbl([W - 4200, 1500, 2700], [
  ['Indicator', 'Date', 'Result'],
  ['Public facilities on 100% renewable electricity (18-municipality area)', '1 Jul 2023',
    '**626 facilities** (586 in Kitakyushu itself)'],
  ['CO₂ avoided by those facilities', '1 Jul 2023', '**10,936.5 t/year**'],
  ['Firms certified under the Decarbonised Electricity Certification scheme', '1 May 2023',
    '**36 companies, 48 sites**; c. 44,550 t CO₂/yr'],
  ['PV installed under the Leading Area programme', 'FY2022',
    '16 public facilities (350 kW); 1 Eco-Town firm (200 kW)'],
  ['Batteries installed', 'FY2022', '14 public facilities (230 kWh); 1 private site (10 kWh)'],
  ['Air conditioning installed', 'FY2022', '28 public facilities; grant spend ¥100m and ¥370m'],
], { header: true, size: 17 }));
R(H4('Three design choices that matter'));
R(bullet('**A hard price rule.** The city will only sign a PPA where the resulting tariff is at or '
  + 'below the electricity price it already pays. Decarbonisation is not permitted to cost more. This '
  + 'is what has kept the programme politically durable — and also what limits how fast it can spread, '
  + 'since it excludes every building where the arithmetic does not work.'));
R(bullet('**Speed, not just cost.** The clearest case is air conditioning for school kitchens, '
  + 'installed for heat-stress protection of staff. Conventional procurement would have taken about 13 '
  + 'years at roughly 10 schools a year, at c. ¥48m/year. Under third-party ownership the rollout '
  + 'takes about 2 years at c. ¥44m/year, and IoT monitoring extends equipment life from 13 to 15 '
  + 'years. The saving is modest; the eleven years are the point.'));
R(bullet('**Scale by aggregation.** A city of 940,000 cannot make a low-cost PPA model bankable on its '
  + 'own estate. Bundling eighteen municipalities and roughly 3,600 buildings creates a portfolio '
  + 'large enough for private counterparties to bid on. *This is the single most transferable idea in '
  + 'the report.*'));
R(H4('A worked example — Hiagari Citizens’ Centre (Step 2)'));
R(P('12 kW of PV and one 16.4 kWh stationary battery on a building consuming about 40,000 kWh a year. '
  + 'Roughly 13,000 kWh a year is self-consumed, giving about **33% electricity self-sufficiency**. '
  + 'The building is also a designated evacuation shelter: on a full charge the battery can run office '
  + 'lighting for 12 hours and charge the mobile phones of 50 evacuees for 18 hours. In Japan, '
  + 'resilience is frequently the argument that actually unlocks the budget, with carbon as the '
  + 'secondary benefit.'));
R(H4('The certification scheme — recognition converted into money'));
R(P('A company on a decarbonised electricity tariff receives a sticker, a logo and publication on the '
  + 'city’s channels. It also receives eligibility for a municipal finance facility and additional '
  + 'points in the scoring of four separate municipal grant and business-selection schemes. The city '
  + 'spends almost nothing and converts reputation into access to money it already controls. The first '
  + 'hundred firms are designated "decarbonisation pioneer companies".'));
R(H4('Unresolved'));
R(P('A long service contract sits awkwardly with annual municipal budgeting; the city ends the period '
  + 'with contracts rather than assets; and the model reaches only buildings where the price rule '
  + 'holds. Nobody has yet shown what happens at contract expiry.'));

// Case 2
R(caseHead('Case 2 · second-life assets', 'The reuse-panel demonstration at Hibikinada Biotope',
  'Third-party ownership extended to used equipment — and a leasing company brought inside the '
  + 'structure.'));
R(H4('What it is'));
R(P('A small demonstration beside the nature centre at the Hibikinada Biotope: a **5 kW array of '
  + 'reused solar panels**, roughly ten years into their service life, generating for on-site '
  + 'self-consumption. The site is itself one of the facilities covered by the 100% Renewable '
  + 'Kitakyushu Model.'));
R(H4('Why it exists'));
R(P('Japan’s first wave of feed-in-tariff solar is now reaching the end of its 20-year contracts, '
  + 'and a very large volume of panels will come off roofs and fields over the next decade. The '
  + 'demonstration has three stated purposes: to find productive uses for megasolar assets after FIT '
  + 'expiry; to build and stress-test a *third-party ownership model for reused panels*; and to work '
  + 'out the rules such a market would need. It measures two things — the degradation and output '
  + 'efficiency of decade-old panels, and their cost.'));
R(H4('The structure — and why it is the most interesting line in this report'));
R(flow([
  '**Shinryo** (Eco-Town recycler, Mitsubishi Chemical group) supplies and processes the panels',
  '→ **Tokyo Century** (leasing and finance company) — the asset owner and financier',
  '→ **Kitakyushu Power** (local retailer) — the supply relationship',
  '→ **City of Kitakyushu** — host site, convenor, and the party defining the public purpose',
]));
R(P('Case 1 is a procurement technique. This is the same technique with a *financial institution '
  + 'inside it*. A leasing company taking ownership of second-life assets is the beginning of an asset '
  + 'class: if used panels can be underwritten, they can be financed, and if they can be financed the '
  + 'circular-economy claims in Case 3 acquire a balance sheet. Kitakyushu’s weakness across thirty '
  + 'years, as Case 9 argues, has been precisely the absence of private capital. This is the smallest '
  + 'case in the report and possibly the most consequential.'));
R(H4('Unresolved'));
R(P('5 kW is a demonstration, not a market. Whether reused panels can be insured, warranted and '
  + 'depreciated on terms a lessor will accept at scale is exactly what the experiment is meant to '
  + 'find out, and it had not reported when these materials were produced.'));

// Case 3
R(caseHead('Case 3 · circular industry', 'Eco-Town: Shinryo and Recycle Tech',
  'The clean-up converted into an industry — and the industry now recycling the energy '
  + 'transition’s own waste.'));
R(H4('The firms'));
R(facts([
  ['Shinryo Co., Ltd.', 'Founded October 1964. Head office in Yahatanishi-ku, Kitakyushu. Capital '
    + '¥500m, wholly owned by Mitsubishi Chemical. Sales ¥25.9bn (FY2021), c. 1,254 staff (March '
    + '2022). Neutralises and converts waste acid, alkali and oil into fuel and cement feedstock; '
    + 'precision cleaning of semiconductor and display equipment parts; wafer reclamation; fine '
    + 'chemicals.'],
  ['Recycle Tech Co., Ltd.', 'Founded December 1998; capital ¥30m; owned by Shinryo and Ricoh. '
    + 'Operating in the Hibikinada Eco-Town since 1998 under the national Eco-Town designation. Annual '
    + 'capacity 5,400 t of office equipment — c. 3,600 copiers and 4,500 PCs and printers per month — '
    + 'dismantled by hand to reach **99% resource recovery**. Employment of people with disabilities '
    + 'is an explicit part of the operating concept.'],
], 2600));
R(H4('Solar panel recycling'));
R(P('The newer line is photovoltaic panel recycling, planned at **90,000 panels a year**, with a '
  + 'stated recycling rate above **99%**. Panels are sorted into reuse (sold into the second-hand '
  + 'market) and recycling; the high-grade line separates the aluminium frame and sheet glass, '
  + 'thermally decomposes the EVA encapsulant, and recovers silver, copper, glass cullet and silicon '
  + 'cells.'));
R(keyfig('Per 1 MW of decommissioned panels (c. 86 t): aluminium **c. 13 t**, glass **c. 54 t**, '
  + 'silver and copper **c. 1 t**  ·  recycling rather than landfill or road-base use avoids '
  + '**c. 200 t CO₂**  ·  the same 1 MW of megasolar avoids **c. 18,000 t CO₂** over 20 years against '
  + 'coal generation'));
R(H4('What it demonstrates'));
R(P('Eco-Town is the one part of Kitakyushu’s strategy that indisputably produced an industry, and '
  + 'it did so by the least fashionable route: a national designation, a designated site, and large '
  + 'incumbent chemical and electronics firms willing to put subsidiaries there. It is not a start-up '
  + 'story. It is also now closing a loop with Case 1 — the panels the city is installing under '
  + 'third-party ownership will be handled at end of life by firms three kilometres away, and '
  + 'second-life panels and EV batteries from those firms are being cascaded back into the '
  + 'programme.'));
R(H4('Unresolved'));
R(P('The recovered materials are commodities with thin margins, and the volumes depend entirely on a '
  + 'decommissioning wave whose timing is set by national FIT policy, not by the city.'));

// Case 4
R(caseHead('Case 4 · offshore wind', 'Kitakyushu Hibikinada Offshore Wind Farm',
  'The largest offshore wind farm in Japan at completion — tendered by the municipality, owned by '
  + 'five national companies.'));
R(keyfig('**25 turbines of 9,600 kW**, maximum output **220,000 kW**  ·  sea area c. **2,700 ha**, '
  + '1–10 km north–south by 11 km east–west, water depth 8–30 m  ·  rotor diameter **174 m**, blade '
  + 'tip c. **200 m** above sea level  ·  jacket (fixed-bottom) foundations  ·  output c. '
  + '**500 GWh/year** ≈ **170,000 households** ≈ **40% of the city’s households**  ·  **20-year** '
  + 'generation business  ·  construction began **13 March 2023**, commissioning within FY2025'));
R(H4('How it came about'));
R(P('Not through the national auction system, which did not yet exist. In **February 2017** the *City '
  + 'of Kitakyushu* ran a public tender for the port water area — possible because of the 2016 Port '
  + 'and Harbour Act amendment — and selected the consortium that became Hibiki Wind Energy K.K., '
  + 'incorporated in April 2017. The company then carried out wind and marine surveys and '
  + 'environmental impact assessment, obtained wind farm certification, and started construction in '
  + 'March 2023. Six years from tender to first steel.'));
R(H4('Who owns it'));
R(tbl([3600, 1100, W - 4700], [
  ['Shareholder', 'Share', 'What they are'],
  ['J-POWER (Electric Power Development Co.)', '40%', 'National wholesale generator'],
  ['Kyuden Mirai Energy', '30%', 'Renewables arm of Kyushu Electric Power'],
  ['Hokutaku', '10%', 'Wind turbine O&M specialist'],
  ['Saibu Gas', '10%', 'Regional gas utility'],
  ['Kyudenko', '10%', 'Electrical engineering contractor'],
], { header: true }));
R(P('The construction chain is similarly national: Vestas Japan supplies the turbines; a Penta-Ocean '
  + '/ Nippon Steel Engineering joint venture builds and installs the foundations; J-POWER Hitec does '
  + 'the onshore electrical work; a Penta-Ocean / Wakachiku joint venture builds the O&M port; Tokyo '
  + 'Kisen operates the crew transfer vessels.'));
R(H4('The municipality’s actual role'));
R(P('Kitakyushu’s contribution was the water area, the tender, the base port, and legitimacy — the '
  + 'fisheries negotiations, the citizen briefings, the local-firm matching sessions with '
  + 'manufacturers, the seminars and the links to local education. The company describes its ambition '
  + 'as a wind farm "that walks with the region". **The city holds no equity.** It captured the '
  + 'project by owning the seabed access and the port, not by investing.'));
R(box('The comparison I would like to draw', [
  'This is, I think, the sharpest single contrast with Denmark and Sweden. A Japanese municipality can '
  + 'convene a 220 MW project, shape it, and host it — and end up with property tax, some local '
  + 'contracts and a photograph of the mayor at the groundbreaking. Whether that is a failure or '
  + 'simply a different and defensible division of labour is a question I cannot answer from the '
  + 'Japanese side alone. Note also the capacity arithmetic: 25 × 9.6 MW is 240 MW of turbines '
  + 'declared at a maximum output of 220 MW — the grid connection, not the wind, sets the ceiling.',
], true));

// Case 5
R(caseHead('Case 5 · industrial policy', 'Green Energy Port Hibiki',
  'A port used as the instrument a municipality actually possesses.'));
R(H4('What it is'));
R(P('A programme begun by the city in **2011** — before the feed-in tariff, eight years before '
  + 'Japan’s offshore wind legislation — to convert the Hibikinada district into a comprehensive '
  + 'base for the wind industry, built on four functions: import/export and domestic transhipment; '
  + 'turbine marshalling and load-out; operations and maintenance; and manufacturing of turbines, '
  + 'components and structures.'));
R(P('In **September 2020** Kitakyushu Port was designated a base port under the Port and Harbour Act — '
  + 'one of four nationally (with Akita, Noshiro and Kashima) and **the only one in western Japan**. '
  + 'Designation matters physically as well as legally: turbine components now exceed 60 m in length '
  + 'and 400 t in weight, and require reinforced quays and large back-up storage yards that only a '
  + 'designated port will receive public investment to build.'));
R(H4('Phases'));
R(tbl([2100, W - 2100], [
  ['**Phase 1 (2013–)**', 'A demonstration zone: onshore testing of offshore turbines, an O&M base, '
    + 'training facilities, and a solar–wind hybrid demonstration, procured by inviting proposals that '
    + 'would lead to industrial agglomeration.'],
  ['**Phase 2 (2016–)**', 'Attracting Japan’s first full-scale offshore wind farm into the port '
    + 'area — Case 4 — explicitly to create domestic demand and give the agglomeration something to '
    + 'serve.'],
  ['**Phase 3**', 'Building the base port and the surrounding industrial land, and recruiting firms '
    + 'and vessels to it.'],
  ['**Phase X**', 'Adapting to floating foundations and to turbines in the 12–16 MW class, whose blade '
    + 'tips exceed 300 m.'],
]));
R(H4('Turbine scale as a proxy for the whole story'));
R(P('Ten 1.5 MW GE turbines installed by NS Wind Power Hibiki in March 2003 reached 100 m at the tip. '
  + 'Two 3.3 MW Vestas machines installed by Wind Energy Research Park in January 2018 reached 140 m. '
  + 'The 9.5 MW class planned for the wind farm reaches 200 m. The 12–16 MW machines now in '
  + 'development exceed 300 m. Every one of those steps re-specifies the quay, the vessels and the '
  + 'factories — which is why a port strategy has to be a rolling one.'));
R(H4('The supply chain the port assembled'));
R(P('Nippon Steel Engineering (jacket foundations, Wakamatsu); thyssenkrupp rothe erde Japan (large '
  + 'slewing bearings, Hibikinada); Furukawa Electric Industrial Cable (tower and nacelle cables, '
  + 'Moji); Yaskawa Electric (generators and electrical products, Yukuhashi); Regency Steel Japan '
  + '(steel structures, Tobata); Ishibashi Manufacturing (gearboxes, Nōgata); Penta-Ocean (SEP jack-up '
  + 'vessel); Tokyo Kisen (crew transfer vessels); Hokutaku (O&M). Distances from Hibikinada to the '
  + 'other base ports — Noshiro 630, Kashima 620, Akita 580 nautical miles — and to Nagasaki 155 and '
  + 'Taipei 750, define the market it can actually serve.'));
R(H4('What it demonstrates, and what is unresolved'));
R(P('This is the clearest case of a Japanese municipality using an asset it genuinely owns to shape '
  + 'private investment, and doing so a decade ahead of national policy. Whether it pays depends on '
  + 'something Kitakyushu does not control: how much offshore wind Japan actually builds in the '
  + 'western sea areas, and whether floating foundations shift the industrial centre of gravity '
  + 'elsewhere.'));

// Case 6
R(caseHead('Case 6 · technology capability', 'FLOAT RAISER — a state-subsidised construction vessel',
  'How Japan builds an industrial capability when no private party will fund the first unit.'));
R(H4('What it is'));
R(P('A semi-submersible spud barge built from FY2016 with a subsidy from the Ministry of the '
  + 'Environment’s programme for cost reduction and diffusion of low-carbon floating offshore wind. '
  + 'Owned and operated by Offshore Wind Farm Construction Co., Ltd., a vehicle of Toda Construction '
  + 'and Yoshida-gumi. The name means the vessel (*raiser*) that stands a float upright.'));
R(keyfig('**110.0 m × 43.0 m**, depth 6.8 m, draught 4.7 m, gross tonnage **12,300 t**  ·  deck area '
  + 'c. **3,890 m²**, deadweight c. **13,500 t**  ·  roll-on capability for loads up to c. **5,000 t**, '
  + 'from either the long or the short side  ·  submerges to **7.4 m** below deck level, and can do so '
  + 'trimmed fore or aft  ·  four 40 m spuds; ballast pumping 1.2–6.7 m/h'));
R(H4('What it does'));
R(P('A hybrid spar float for a floating wind turbine is built on land, rolled onto the barge on '
  + 'multi-axle trailers with the ballast adjusted as it comes aboard, towed offshore, and then '
  + 'floated off as the barge submerges. The same vessel can launch caissons, jackets and work '
  + 'vessels, so it serves fixed-bottom construction and general marine works as well.'));
R(H4('What it demonstrates'));
R(P('Floating offshore wind is where Japan’s resource actually is — the shelf is narrow and deep '
  + 'water begins close inshore — but no contractor will build the first specialised vessel for a '
  + 'market that does not yet exist. The state broke that deadlock by subsidising the asset rather '
  + 'than the electricity. It is a different instrument from everything else in this report: not a '
  + 'tariff, not a grant to a municipality, but public money placed directly into a private industrial '
  + 'capability. In a Nordic setting that same gap has often been closed by a foundation or a state '
  + 'holding company — which is one reason I want to understand the Danish and Swedish ownership '
  + 'vehicles properly.'));
R(H4('Unresolved'));
R(P('One subsidised vessel is not a supply chain. Whether Japanese yards capture floating wind, or '
  + 'whether the work migrates to lower-cost yards abroad while Japan keeps only installation, is '
  + 'undecided.'));

// Case 7
R(caseHead('Case 7 · the foreign operator', 'Ørsted in Japan',
  'A Danish company as the direct bridge — and as the demonstration of a corporate transformation '
  + 'Japan has no equivalent of.'));
R(H4('The company'));
R(P('Ørsted built the world’s first offshore wind farm in Denmark in 1991 and holds roughly **26%** '
  + 'of the global offshore wind market. From 2008 it began shifting from black to green energy: since '
  + 'then it has cut coal consumption by **73%** and halved its carbon emissions. It has **28 offshore '
  + 'wind farms in operation** across Denmark, the UK, Germany, the Netherlands, the US and Taiwan, '
  + 'having installed **1,600 turbines** totalling **7.6 GW**, supplying about **18 million people**, '
  + 'with a further **3.4 GW** under construction and around **3,300** dedicated staff. Roughly **one '
  + 'in four** offshore turbines worldwide was installed by the company. It intends to quadruple its '
  + 'renewable capacity to **50 GW by 2030**.'));
R(H4('Its own decarbonisation timetable'));
R(tbl([1200, W - 1200], [
  ['**2023**', 'Complete exit from coal-fired generation'],
  ['**2025**', 'Carbon neutrality in its own operations'],
  ['**2032**', '50% of the supply chain decarbonised'],
  ['**2040**', 'Carbon neutrality across the entire business including the supply chain'],
]));
R(H4('In Japan'));
R(P('Ørsted entered the Japanese market in 2019 and opened a Tokyo office in May 2019, in order to '
  + 'pursue offshore wind in the Asia-Pacific. Its case for Japan is straightforward: over 70% of the '
  + 'country is mountainous, and about three quarters of 127 million people live in coastal urban '
  + 'areas — so offshore wind is close to the ideal form of large-scale domestic renewable '
  + 'generation.'));
R(H4('Why it is in this report'));
R(P('Three reasons. First, it is a literal Denmark–Japan bridge: the questions I want to ask in '
  + 'Copenhagen and Malmö about Danish energy governance have a direct Japanese footprint. Second, it '
  + 'is the counter-example to Case 4 — a single company that took the whole transformation onto its '
  + 'own balance sheet, rather than a consortium assembled per project. Third, and most relevant to '
  + 'our conversation, its history is inseparable from Danish ownership arrangements: a company with a '
  + 'majority state shareholding, in a country with a strong tradition of foundation ownership and '
  + 'cooperative wind, was able to sustain a twenty-year strategic reorientation. No Japanese utility '
  + 'has done anything comparable, and I suspect the reason is structural rather than managerial.'));
R(caveat('The Ørsted figures above are taken from a Japanese-language corporate profile collected in '
  + '2023 and reflect the company’s position at that time; they should be refreshed before '
  + 'citation.'));

// Case 8
R(caseHead('Case 8 · the counter-case',
  'Gotō, Nagasaki — E-WIND, citizens’ power, and a training association',
  'A small island city that retained locally what a city of 940,000 did not.'));
R(H4('The firm'));
R(P('E-WIND Ltd. is based in Tomie-chō, Gotō City, in the Gotō islands off Nagasaki. Founded December '
  + '1990; capital ¥15m; **50 staff** at February 2022, average age **39**. It entered wind turbine '
  + 'maintenance in 2006 and consolidated entirely into wind O&M in 2008, when it took its present '
  + 'name. It now runs branch offices in Hokkaido, Muroran, Wakayama, Kagoshima and Fukue, and '
  + 'monitors **more than 100 onshore and offshore turbines** from Hokkaido to Okinawa, 24 hours a '
  + 'day, year-round, from its own information management centre — able to shut a turbine down '
  + 'remotely. Technicians receive about three months of training in the company’s own facility '
  + 'followed by a year accompanying senior engineers. The company developed its own tablet-based '
  + 'maintenance management system. It was recognised by METI as a *Regional Future Driving '
  + 'Enterprise* (2017) and among the *300 SMEs taking flight* for regional contribution (2016), and '
  + 'received a special jury prize in the New Energy Awards (2020).'));
R(H4('What it built around itself — the part that matters'));
R(bullet('**Gotō Citizens’ Power (2018).** A local retailer, established so that residents and '
  + 'local businesses would experience the benefit of the offshore wind and solar generation around '
  + 'them. E-WIND was closely involved in founding it and supports its product design, marketing and '
  + 'agency network. The retail brand is "Gotō no Denki" — Gotō’s electricity.'));
R(bullet('**Nagasaki Wind Service Group (2018).** A voluntary association whose explicit purpose is to '
  + 'help *other* firms enter turbine maintenance: a common training curriculum, and visualisation of '
  + 'member technicians’ skill levels, to relieve the technician shortage created by the '
  + 'acceleration of the renewables market. A company deliberately creating its own competitors, '
  + 'because the constraint is the labour pool rather than the order book.'));
R(bullet('**Its own building.** The head office runs on 100% Gotō-produced renewable energy, with '
  + 'rooftop PV and batteries under an **on-site PPA** and the balance bought from Gotō no Denki. '
  + 'Company EVs double as emergency power and the stored electricity is opened to the community in a '
  + 'disaster. The arrangement is certified under a local "Gotō-version RE100".'));
R(bullet('**Farmland regeneration.** Abandoned agricultural land is restored to produce rice, '
  + 'vegetables and camellia oil from Gotō’s camellia woods, returned to residents.'));
R(box('Why I am putting a 50-person company beside a 220 MW wind farm', [
  'Because on the question this report asks, Gotō scores better than Kitakyushu. The turbines around '
  + 'Gotō are owned by outside investors, exactly as at Hibikinada. But the *maintenance* — the '
  + 'twenty-year revenue stream, the skilled employment, the training system and the retail '
  + 'relationship with residents — is owned locally, by a company that then built a citizens’ '
  + 'retailer and a training association around itself. Kitakyushu has the port, the consortium and '
  + 'the industrial estate; Gotō has the recurring local income. The Japanese offshore wind industry '
  + 'has, so far, tended to award the construction to Tokyo and the O&M to the region, and Gotō is '
  + 'what it looks like when a region takes that seriously.',
], true));
R(H4('Unresolved'));
R(P('Whether the model scales. Gotō’s advantage is partly that it is small, remote and cohesive '
  + 'enough for one firm to matter; the same strategy in a city of 940,000 would face incumbents Gotō '
  + 'does not have.'));

// Case 9
R(caseHead('Case 9 · the long record', 'FAIS and the Science and Research Park',
  'Thirty years of environment-and-knowledge-led growth strategy, and a disappointing economic '
  + 'result.'));
R(P('Everything above is what the city presents, and most of it is genuinely working. My own research '
  + 'concerns the longer record, and it is less flattering. Kitakyushu has run an '
  + 'environment-and-knowledge-led growth strategy for three decades through the Science and Research '
  + 'Park and its agency FAIS, and the city’s gross product has nonetheless grown slowly relative '
  + 'to other designated cities. The transition has been real as environmental policy and weak as '
  + 'economic policy. I examined this through a growth-contribution analysis of Kitakyushu against '
  + 'Fukuoka City from 1980 to 2020, and through the record of FAIS’s own programmes.'));
R(H4('The industries targeted were not the industries that grew'));
R(tbl([1800, W - 1800], [
  ['Period', 'Industries the city and FAIS set out to grow'],
  ['1994–1998', 'Imaging and video, aerospace, airport-related industry'],
  ['1999–2003', 'Environment (recycling), information technology, welfare, housing, lifestyle and '
    + 'culture, biotechnology'],
  ['2004–2008', 'Environment (recycling, eco-premium products, biomass), information technology, '
    + 'welfare devices, biotechnology, next-generation robotics, semiconductors'],
  ['2019', 'Semiconductors, car electronics, information and communication (system LSI), robotics'],
], { header: true }));
R(P('What in fact contributed most to the city’s product in the most recent period was health care '
  + 'and social work, real estate, and miscellaneous manufacturing. The target list was rewritten '
  + 'roughly every five years, largely tracking whichever industry national programmes were funding at '
  + 'the time, while the productive structure underneath moved towards lower-value, locally-consumed '
  + 'services.'));
R(H4('Three explanations I am testing'));
R(numbered('**The success criterion was commercialisation, not value added.** FAIS assessed projects '
  + 'on technical feasibility, the prospect of reaching a product, and employment — not on revenue or '
  + 'contribution to the city’s economy. A project that reached market counted as a success even if '
  + 'it never grew. National schemes attached to some of the same projects *did* carry revenue '
  + 'targets, and the robotics programme, which used them, is one of the areas that actually '
  + 'expanded.'));
R(numbered('**There was never enough private capital.** No meaningful venture capital layer formed '
  + 'locally. Regional banks fund company formation rather than R&D, so FAIS financed research from '
  + 'its own budget — bounded by a shrinking municipal balance sheet. The comparison I use is '
  + 'Pittsburgh, where university-led renewal was carried by roughly a billion dollars a year of '
  + 'investment, and where venture investors brought a commercial judgement into the decision to fund '
  + 'at all.'));
R(numbered('**The supporting services never agglomerated.** A technology cluster needs accountants, '
  + 'patent and corporate lawyers, marketing and investors within reach. Around the Science and '
  + 'Research Park these are thin, and the city substituted by providing them itself — which keeps the '
  + 'cluster running but caps it at the scale a municipality can fund, and removes the competition '
  + 'that would improve the services.'));
R(caveat('This case reports work in progress. The contribution analysis is not yet published and the '
  + 'figures behind it are still being revised; I present the structural argument rather than the '
  + 'numbers, and would rather be corrected than believed.'));

// ---- Part III
R(partHead('Part III', 'Synthesis',
  'The nine cases compared on the three questions, five findings, and what is structurally missing.'));

R(H2('4  The nine cases compared'));
R(tbl([2000, 1700, 1700, W - 5400], [
  ['Case', 'Who owns the asset', 'Who supplies the capital',
    'Who carries the risk — and what the municipality ends up holding'],
  ['**1 Municipal estate PPA**', 'Equipment company', 'Private, via consortium',
    'Provider carries technical risk; city carries a long tariff obligation. City holds *contracts*.'],
  ['**2 Reuse-panel demonstration**', 'Leasing company (Tokyo Century)', 'Private finance',
    'Lessor carries residual-value risk on second-life assets — the first case here where a financial '
    + 'institution takes it. City holds a *method*.'],
  ['**3 Eco-Town recyclers**', 'Private firms (Mitsubishi Chemical, Ricoh groups)',
    'Corporate parents; national Eco-Town designation',
    'Firms carry commodity-price risk. City holds *land, designation and an industry*.'],
  ['**4 Hibiki offshore wind**', 'SPC owned by five national companies',
    'Corporate equity and project finance',
    'Consortium carries all of it. City holds *tax, a tender record and a port* — no equity.'],
  ['**5 Green Energy Port**', 'Public port; private facilities on it',
    'National and municipal port investment',
    'Public sector carries demand risk on the port; firms carry their own. City holds *infrastructure '
    + 'and locational advantage*.'],
  ['**6 FLOAT RAISER**', 'Private contractors (Toda, Yoshida-gumi)',
    'Ministry of the Environment subsidy plus private',
    'State bought down first-mover risk directly. Nation holds a *capability*.'],
  ['**7 Ørsted**', 'The company itself', 'Corporate balance sheet',
    'One firm carries a twenty-year reorientation — possible under Danish ownership arrangements.'],
  ['**8 Gotō / E-WIND**', 'Generation: outside investors. O&M and retail: local',
    'Small local firm; community retailer',
    'Local firm carries operating risk and keeps the recurring income. Region holds *skills, jobs and '
    + 'a customer relationship*.'],
  ['**9 FAIS / Science Park**', 'Public agency', 'Municipal budget and national grants',
    'City carries everything, at a scale its budget allows. City holds *projects that reached market '
    + 'but did not grow*.'],
], { header: true, size: 16 }));

R(H2('5  Five findings'));
R(H3('5.1  Kitakyushu buys services; it does not accumulate assets'));
R(P('Read down the ownership column and only Case 5 leaves the municipality owning anything. This is '
  + 'not incompetence — it is the rational response to weak fiscal autonomy and no municipal utility. '
  + 'But it means the city’s capability is contractual. When a contract ends, or a provider leaves, '
  + 'or the national grant regime changes, there is little accumulated capital to fall back on. A '
  + 'Nordic city that spent thirty years owning heat networks, housing and water has, by construction, '
  + 'a different set of options in 2026.'));
R(H3('5.2  The binding constraint is the absence of a private capital layer'));
R(P('Case 9 says it directly: no venture capital, regional banks that finance company formation rather '
  + 'than R&D, and an agency forced to fund research from a municipal budget. Case 6 shows the state '
  + 'stepping in where private capital would not. Case 2 is the only place where a financial '
  + 'institution voluntarily takes a novel risk — and it is a 5 kW demonstration. Every large project '
  + 'in this report was financed by national corporations, national subsidy, or the city itself. '
  + 'Nothing was financed by a local investor.'));
R(H3('5.3  Aggregation is the one instrument that clearly works'));
R(P('Bundling eighteen municipalities and 3,600 buildings turned an unbankable municipal estate into a '
  + 'portfolio private counterparties would bid for. The same logic explains why the port strategy '
  + 'worked and the science park did not: the port aggregated demand for a whole region and a whole '
  + 'industry, while FAIS supported projects one at a time. Where Kitakyushu has created scale it has '
  + 'attracted capital; where it has acted retail, it has had to pay for everything itself.'));
R(H3('5.4  The value that stays local is operations, not construction'));
R(P('Construction contracts go to national firms — Penta-Ocean, Nippon Steel Engineering, Vestas, '
  + 'J-POWER. What can stay is the twenty-year maintenance relationship, and Gotō shows a region '
  + 'capturing exactly that, together with the training system that makes it defensible. Kitakyushu '
  + 'has begun to do the same through Hokutaku’s shareholding and the O&M port, but Gotō went '
  + 'further by building a citizens’ retailer and an association that trains its own competitors.'));
R(H3('5.5  Resilience, not carbon, is what unlocks Japanese budgets'));
R(P('The battery at Hiagari Citizens’ Centre is justified by evacuation-shelter duty; E-WIND’s '
  + 'EVs are justified by disaster power supply; the Kitakyushu energy park literature leads with the '
  + 'city’s distance from plate boundaries. Carbon is real but it is rarely the argument that '
  + 'carries a committee. Any transfer of Japanese practice to Europe, or European practice to Japan, '
  + 'has to survive that translation.'));

R(H2('6  The missing layers'));
R(P('Setting the nine cases side by side, four things are absent from the Japanese municipal '
  + 'transition that a Nordic reader would expect to find:'));
R(facts([
  ['A municipal utility', 'No ownership of generation, networks or heat; a minority position in a '
    + 'retailer is the whole of it. The city cannot cross-subsidise, cannot borrow against an asset '
    + 'base, and cannot use tariffs as policy.'],
  ['Patient equity', 'No foundation, holding company or public investment vehicle able to hold an '
    + 'asset for twenty years on a non-commercial mandate. Case 6 shows the state substituting with a '
    + 'one-off subsidy.'],
  ['A local investor base', 'No venture capital; regional banks lending against company formation '
    + 'rather than technology. Case 2’s leasing company is the first sign of one.'],
  ['Business services agglomeration', 'The accountants, patent lawyers and marketing capability a '
    + 'cluster needs are thin, and the city has substituted for them itself — capping scale and '
    + 'removing competition.'],
], 3200));
R(P('Those four absences are, I think, one absence described four ways: **nobody in the region owns '
  + 'capital with a long horizon and a public purpose.** Denmark answers that with foundation '
  + 'ownership and municipal holding companies; Sweden answers it differently again. That is the '
  + 'substance of what I want to learn.'));

// ---- Part IV
R(partHead('Part IV', 'Fourteen questions for the City of Malmö',
  'Grouped by theme. I have deliberately made no claims about Malmö’s own figures — I would rather '
  + 'ask than assume.'));

const Q = (n, head, body) => [
  new Paragraph({
    spacing: { before: 160, after: 40 }, keepNext: true,
    children: [new TextRun({ text: `${n}. ${head}`, font: 'Calibri', size: 19, bold: true })],
  }),
  P(body, { run: { color: MUTED, size: 19 }, after: 100 }),
];

R(H3('A.  The shipyard and the aftermath'));
R(Q(1, 'Who absorbed the cost?', 'When Kockums wound down, who carried it — the municipal balance '
  + 'sheet, the state, or private developers? Kitakyushu never had a single closure event; its steel '
  + 'industry shrank slowly and is still there, which may have made the transition politically easier '
  + 'and economically slower.'));
R(Q(2, 'What happened to the workforce?', 'Japanese transitions are often cushioned by redeployment '
  + 'inside the same corporate group, which protects individuals but slows structural change. Did '
  + 'Sweden’s separation of employment protection from the individual firm produce a faster or a '
  + 'more painful adjustment — and which would you choose again?'));
R(Q(3, 'Was Västra Hamnen an economic project or a symbolic one?', 'Kitakyushu’s equivalents — '
  + 'Eco-Town, the Science Park — succeeded and failed respectively, and the difference seems to be '
  + 'whether incumbent industry had a reason to be there. What made the difference in Malmö?'));

R(H3('B.  Ownership and money'));
R(Q(4, 'What does the City of Malmö still own?', 'Across district heating, water, housing and any '
  + 'energy assets — and how much of your climate programme depends on that ownership rather than on '
  + 'regulation or procurement?'));
R(Q(5, 'Municipal capital, regulation, or purchasing power?', 'Roughly what share of your targets is '
  + 'delivered by each? I would like to compare the mix rather than the targets.'));
R(Q(6, 'Is there patient capital with a public purpose?', 'I am also visiting Lolland and its holding '
  + 'company LOKE, and studying Danish foundation-owned enterprises. Does Malmö have any vehicle that '
  + 'lets public purposes be pursued with private capital on a long horizon — and if not, what fills '
  + 'the gap?'));
R(Q(7, 'Did you ever have to buy down first-mover risk directly?', 'Case 6 in this report is a '
  + 'construction vessel the Japanese state paid for because no contractor would build the first one. '
  + 'Has Malmö or Region Skåne ever had to fund an asset rather than an outcome?'));

R(H3('C.  Mechanisms I would like to test against yours'));
R(Q(8, 'Third-party ownership on municipal buildings.', 'Do you use PPA or energy-service models, and '
  + 'did they survive procurement law? In Japan the weak point is that a long service contract sits '
  + 'badly with annual budgeting; I would expect EU procurement rules to create a different but '
  + 'comparable friction.'));
R(Q(9, 'The price rule.', 'Kitakyushu will only do a PPA where the tariff is at or below the price it '
  + 'already pays — decarbonisation is not permitted to cost more. Does Malmö operate any equivalent '
  + 'discipline, and has it helped or constrained you?'));
R(Q(10, 'Scale through your neighbours.', 'Kitakyushu bundled eighteen municipalities and 3,600 '
  + 'buildings to make its programme large enough to attract bidders. How does Malmö achieve scale '
  + 'with Lund, Burlöv and Region Skåne — and does joint procurement work in practice or mainly on '
  + 'paper?'));
R(Q(11, 'Second-life assets.', 'Case 2 is an attempt to make used solar panels financeable. '
  + 'Sweden’s first large wave of renewables will decommission sooner than Japan’s. Is anyone '
  + 'building a market in second-life equipment, and will a lessor touch it?'));

R(H3('D.  The emissions you do not control'));
R(Q(12, 'The 64% problem.', 'Two-thirds of Kitakyushu’s emissions belong to private industry the '
  + 'city cannot regulate and cannot afford to lose. Does Malmö face the same asymmetry — and if so, '
  + 'what has actually worked: agreements, planning conditions, procurement, or public pressure?'));
R(Q(13, 'Keeping the operations work local.', 'In Japanese offshore wind, the construction goes to '
  + 'national firms and only maintenance can stay in the region. Gotō built a local O&M company, a '
  + 'citizens’ retailer and a training association around that fact. Has Malmö deliberately '
  + 'targeted the recurring, operational share of any industry — and did it hold?'));

R(H3('E.  The question underneath all of it'));
R(Q(14, 'Does owning beat buying?', 'Kitakyushu buys services and ends thirty years with contracts. '
  + 'Malmö, Lolland and the Danish foundation-owned firms kept ownership. Has that left you with more '
  + 'capability — or only with more debt and slower decisions? An honest answer, including the parts '
  + 'that did not work, is worth more to me than a successful case study.'));

// ---- Part V
R(partHead('Part V', 'Sources and method', 'What this report rests on, and what it does not.'));

R(H2('7  Method note'));
R(P('The primary evidence is a set of documents collected during a study visit to Kitakyushu on '
  + '**22 September 2023** by the Public Finance Study Group of Kyoto University — city presentations, '
  + 'company profiles, project brochures and site materials, several of them prepared for that visit. '
  + 'These are promotional and official documents: they are reliable on structure, dates, capacities '
  + 'and ownership, and they are advocacy on outcomes. I have used them for the former and treated the '
  + 'latter sceptically.'));
R(P('Figures are reported as the source gives them, with the source’s date attached, and have not '
  + 'been updated to 2026 — several will have moved, particularly the offshore wind farm’s status '
  + 'and Ørsted’s portfolio. Case 9 draws on my own unpublished research and is flagged as such in '
  + 'the text. Historical background on the pollution period and the 1963 merger comes from the '
  + 'standard literature on Kitakyushu, not from these documents. **No claims are made about Malmö, '
  + 'Lolland or Swedish and Danish practice**; where I would otherwise have asserted something, '
  + 'Part IV asks it instead.'));

R(H2('8  Sources'));
R(srcItem('City of Kitakyushu, Environment Bureau, Renewable Energy Promotion Division, '
  + '「脱炭素先行地域にかかる北九州市の取り組み」 (*Kitakyushu’s initiatives concerning the '
  + 'Decarbonisation Leading Area*), presentation, 22 September 2023. — Cases 1 and 3; Part I.'));
R(srcItem('City of Kitakyushu, Shinryo, Kitakyushu Power and Tokyo Century, '
  + '「リユースパネルを活用した実証事業の概要について」 (*Outline of the demonstration using reused '
  + 'panels*), site material prepared for the Kyoto University Public Finance Study Group visit, '
  + '22 September 2023. — Case 2.'));
R(srcItem('Shinryo Co., Ltd., *Company Profile*, October 2022; Recycle Tech Co., Ltd., company profile '
  + 'and *PV Panel Recycling (PVR)* brochure. — Case 3.'));
R(srcItem('Hibiki Wind Energy K.K., 「北九州響灘洋上ウインドファーム建設工事の概要」 (*Outline of the '
  + 'construction of the Kitakyushu Hibikinada Offshore Wind Farm*), c. June 2023. — Case 4.'));
R(srcItem('City of Kitakyushu, 「グリーンエネルギーポートひびき」 (*Green Energy Port Hibiki*), project '
  + 'brochure. — Case 5; the national policy timeline.'));
R(srcItem('City of Kitakyushu, *Next Generation Energy Kitakyushu*, energy park guide and municipal '
  + 'renewable energy map. — Part I; Case 5.'));
R(srcItem('Offshore Wind Farm Construction Co., Ltd. (Toda Construction, Yoshida-gumi), *FLOAT RAISER* '
  + 'vessel brochure, December 2022. — Case 6.'));
R(srcItem('Ørsted Japan K.K., *Company Profile* (Japanese edition), c. 2022–23. — Case 7.'));
R(srcItem('E-WIND Ltd., *Corporate Information*, c. 2022. — Case 8.'));
R(srcItem('Author’s own research on Kitakyushu’s industrial structure and the Science and '
  + 'Research Park (FAIS), drawing on *Kitakyushu City Economic Accounts* (1996–2020) and FAIS '
  + 'programme documents; with Pittsburgh as the comparator. Unpublished. — Case 9.'));

R(H2('9  A note on what is missing'));
R(P('Three gaps a reader should hold against this report. It is written almost entirely from Japanese '
  + 'sources and therefore reproduces a Japanese framing of what counts as success. It says nothing '
  + 'about the residents of Kitakyushu, who appear here only as evacuees sheltering beside a battery — '
  + 'the citizens’ movement of the 1960s is the last point at which they act. And it treats the '
  + '64% of emissions that belong to heavy industry as a constraint rather than a subject, because the '
  + 'documents I have do so. If our conversation has time for only one thing beyond the questions in '
  + 'Part IV, I would rather it were the third of those.'));

// ================================================================ BRIEFING

const brief = [];
const B = (...x) => brief.push(...x.flat());

B(KICKER('Briefing note · for discussion'));
B(H1('Kitakyushu: from steel and smog to a decarbonisation front-runner',
  'What a Japanese heavy-industry city can and cannot do with its own hands — and why I would like '
  + 'to compare it with Malmö'));
B(metaBlock([
  ['Prepared for', 'Per-Arne Nilsson, Senior Strategist, Environmental Department, City of Malmö'],
  ['Meeting', 'Malmö, 9 September 2026'],
  ['Prepared by', NAME_PLACEHOLDER + '  ·  Nordic field study, 7–11 September 2026 (Danish Energy '
    + 'Agency · Copenhagen Business School · Malmö · Lolland)'],
  ['Subject', 'Municipal-level industrial transition and the financing of decarbonisation'],
]));
B(P('', { after: 160 }));

B(H2('1  Why this case'));
B(P('Malmö and Kitakyushu are not obvious twins, but their problem was the same. Each was a second '
  + 'city built on a single industry — shipbuilding in Malmö, integrated steelmaking in Kitakyushu — '
  + 'and each lost the certainty that industry provided. Both answered by redefining themselves as '
  + 'environmental cities, and both are now judged on whether that redefinition produced an economy or '
  + 'only a reputation.'));
B(P('I have studied Kitakyushu since my undergraduate work, and my interest is deliberately narrow: '
  + '*not* what the city declares, but who carries the capital and who carries the risk. Japanese '
  + 'municipalities have weak fiscal autonomy and, since liberalisation, no municipal utility in the '
  + 'Nordic sense. That constraint has pushed Kitakyushu towards a specific answer — buying energy '
  + 'services rather than owning energy assets — which I suspect is the mirror image of how Malmö, '
  + 'Lolland and the Danish foundation-owned firms have solved it. This note sets out the Japanese '
  + 'side so that our conversation can start from a shared factual base.'));

B(H2('2  Kitakyushu at a glance'));
B(facts([
  ['Established', '1963, by the merger of five cities of equal standing (Moji, Kokura, Wakamatsu, '
    + 'Yahata, Tobata)'],
  ['Area', '491.71 km²'],
  ['Population', 'c. 940,000 (Jan 2022); c. 490,000 households. A designated city in long-term '
    + 'decline from a 1979 peak of over 1.06 million'],
  ['Principal industries', 'Iron and steel, chemicals, cement, ceramics, automotive'],
  ['Greenhouse gas emissions', '14.78 Mt CO₂e (2019)'],
  ['Share from the industrial sector', '**64%** — against 32% for Japan as a whole (national total '
    + '1.21 Gt CO₂e, 2019)'],
  ['Headline target', '−47% GHG by FY2030 against FY2013; net zero by 2050'],
]));
B(P('That 64% is the single most important number in this note. Two-thirds of Kitakyushu’s '
  + 'emissions sit inside a handful of private industrial installations that the city does not own, '
  + 'cannot regulate directly, and cannot afford to lose. Every instrument described below is shaped '
  + 'by that fact — the city works on the third of emissions it can reach, and tries to influence the '
  + 'rest through industrial policy rather than climate policy.'));

B(H2('3  Four phases of the transition'));
B(tbl([1500, 2300, W - 3800], [
  ['Period', 'Phase', 'What happened'],
  ['1901–1963', '**Building the steel city**',
    'The state-owned Yawata Steel Works opens in 1901 and becomes the core of Japanese '
    + 'industrialisation. Five towns grow together around it and merge in 1963.'],
  ['1960s–1970s', '**Pollution and the citizens’ response**',
    'Dokai Bay is declared a "sea of death"; the sky over Yahata is described as having seven '
    + 'colours. Local women’s associations document the damage and force the issue onto the '
    + 'political agenda, producing pollution-control agreements between the city and industry well '
    + 'ahead of national law. This is the origin of the city’s environmental identity — it was '
    + 'imposed from below, not chosen.'],
  ['1980s–2000s', '**Environment as an industry**',
    'The clean-up is converted into an export: the Eco-Town project (from 1997) turns the Hibikinada '
    + 'waterfront into a recycling industrial estate; the Science and Research Park and its agency '
    + 'FAIS (from the 1990s) attempt to build a knowledge economy; the city becomes a hub for '
    + 'environmental cooperation with Asian cities. Selected as an Eco-Model City (2008) and an SDGs '
    + 'Future City (2018).'],
  ['2010s–', '**Energy as an industry**',
    'Green Energy Port Hibiki (from 2011) reorients the port towards offshore wind. Zero Carbon City '
    + 'declaration (Oct 2020); climate emergency declaration (Jun 2021); Green Growth Strategy '
    + '(Feb 2022); selection as a national *Decarbonisation Leading Area* (Apr 2022).'],
], { header: true }));
B(box('The structural parallel worth testing', [
  'In both cities the environmental turn began as damage control and only later became a growth '
  + 'story. The interesting question is whether the growth story is real. Malmö has Västra Hamnen and '
  + 'a university; Kitakyushu has Eco-Town and a science park. Section 6 sets out why I am sceptical '
  + 'about the Japanese half of that claim.',
]));

B(H2('4  What Kitakyushu is doing now'));
B(H3('4.1  Green Growth Strategy (February 2022)'));
B(P('An action plan for a "virtuous circle of environment and economy", built on four pillars: '
  + '(i) decarbonised power and hydrogen supply, (ii) business support for innovation, '
  + '(iii) decarbonised urban development through planning and transport, (iv) export of environmental '
  + 'business to Asia. Stated expected effects: **¥590–680 billion** of economic impact and '
  + '**c. 930,000 t** of CO₂ reduction. Targets for FY2030: renewable capacity of **1,302–1,402 MW** '
  + 'within the city (roughly three times the current level) and hydrogen demand of '
  + '**5,700 t/year**.'));
B(H3('4.2  Decarbonisation Leading Area (April 2022)'));
B(P('A national programme requiring net-zero electricity in the residential and commercial sector by '
  + 'FY2030, intended as a replicable model — the government’s phrase is a "decarbonisation '
  + 'domino". Kitakyushu applied *jointly with 17 neighbouring municipalities* (18 in total) and was '
  + 'selected in the first cohort. Scope and figures:'));
B(bullet('Target estate: approximately **3,600 public facilities** across the 18 municipalities, plus '
  + 'the cluster of recycling companies in Eco-Town.'));
B(bullet('New renewable capacity: **37 MW**; PV at approximately **290 sites** in Kitakyushu and '
  + '**130 sites** in the 17 neighbouring municipalities; **100%** renewable electricity for the '
  + 'target facilities.'));
B(bullet('Cascade use of second-life PV panels and EV batteries, developed with Eco-Town firms and '
  + 'vehicle manufacturers, to lower the cost further.'));
B(bullet('Funded through the national Regional Decarbonisation Transition and Renewable Energy '
  + 'Promotion Grant.'));
B(H3('4.3  The "100% Renewable Kitakyushu Model" (announced February 2021)'));
B(P('A three-step standard roadmap that any facility can be moved along:'));
B(tbl([1400, W - 1400], [
  ['**Step 1**', '100% renewable electricity, supplied from generation *inside the city* (including '
    + 'waste-to-energy plants) through Kitakyushu Power and other retailers.'],
  ['**Step 2**', 'Self-sufficient facility: PV plus battery, installed under third-party ownership. '
    + 'The battery charges when power is cheap and discharges when it is expensive, absorbing surplus '
    + 'local renewable output and cutting the bill.'],
  ['**Step 3**', 'Step 2 plus energy-efficient equipment, also third-party owned, reducing total '
    + 'consumption and therefore the volume of renewable electricity that must be procured.'],
]));
B(H3('4.4  Progress on the ground'));
B(tbl([W - 4200, 1500, 2700], [
  ['Indicator', 'Date', 'Result'],
  ['Public facilities on 100% renewable electricity (18-municipality area)', '1 July 2023',
    '**626 facilities** (586 in Kitakyushu itself)'],
  ['CO₂ avoided by those facilities', '1 July 2023', '**10,936.5 t/year**'],
  ['Firms certified under the city’s Decarbonised Electricity Certification scheme', '1 May 2023',
    '**36 companies, 48 sites**; c. 44,550 t CO₂/year'],
  ['PV installed under the Leading Area programme, FY2022', 'FY2022',
    '16 public facilities (350 kW) and 1 Eco-Town firm (200 kW)'],
  ['Batteries installed, FY2022', 'FY2022',
    '14 public facilities (230 kWh) and 1 private site (10 kWh)'],
  ['Air conditioning installed, FY2022', 'FY2022',
    '28 public facilities; grant spend ¥100m and ¥370m respectively'],
], { header: true, size: 17 }));
B(P('The certification scheme is worth a moment. It costs the city almost nothing: a company on a '
  + 'decarbonised electricity tariff gets a sticker, a logo, publication on the city’s channels, '
  + 'and — the part that bites — eligibility for a municipal finance facility and extra points in the '
  + 'scoring of four separate municipal grant and business-selection schemes. Recognition is converted '
  + 'into access to money the city already controls.'));

B(H2('5  The mechanism most likely to interest you: third-party ownership'));
B(P('Kitakyushu does not buy solar panels. It buys electricity from panels that someone else owns and '
  + 'that happen to sit on its roofs. The city’s own slogan for this is a shift **"from ownership '
  + 'to use"**. The structure:'));
B(flow([
  '**Equipment company** owns and installs the PV, battery or air-conditioning unit',
  '→ **JV / consortium** operates and manages the equipment, with IoT monitoring',
  '→ **Retailer** (Kitakyushu Power and others) holds the electricity supply contract',
  '→ **Public or private facility** pays only an electricity tariff — no capital expenditure',
]));
B(P('Three design choices make this more than a financing trick:'));
B(bullet('**A hard price rule.** The city will only do a PPA at a facility where the resulting tariff '
  + 'is at or below the electricity price it already pays. Decarbonisation is not permitted to cost '
  + 'more. This is the discipline that has kept the programme politically durable — and also what '
  + 'limits how fast it can spread.'));
B(bullet('**Speed, not just cost.** The clearest case is air conditioning for school kitchens, '
  + 'installed for heat-stress protection of staff. Conventional municipal procurement would have '
  + 'taken about 13 years at roughly 10 schools a year, at c. ¥48m/year. Under third-party ownership '
  + 'the rollout takes about 2 years at c. ¥44m/year, and IoT monitoring extends equipment life from '
  + '13 to 15 years. The saving is real but modest; the eleven years are the point.'));
B(bullet('**Scale by aggregation.** A city of 940,000 cannot make a low-cost PPA model bankable on its '
  + 'own estate alone. Bundling 18 municipalities and roughly 3,600 buildings is what creates a '
  + 'portfolio large enough for private counterparties to bid on.'));
B(H3('A worked example — Hiagari Citizens’ Centre (Step 2)'));
B(P('12 kW of PV and one 16.4 kWh stationary battery on a building consuming about 40,000 kWh a year. '
  + 'Roughly 13,000 kWh a year is self-consumed, giving about **33% electricity self-sufficiency**. '
  + 'The building is also a designated evacuation shelter: on a full charge the battery can run the '
  + 'office lighting for 12 hours and charge the mobile phones of 50 evacuees for 18 hours. In Japan, '
  + 'resilience is frequently the argument that actually unlocks the budget, with carbon as the '
  + 'secondary benefit — I would be interested to know whether Malmö has an equivalent "second '
  + 'reason".'));
B(H3('The wider industrial build-out'));
B(bullet('**Offshore wind.** Kitakyushu Port was designated in September 2020 as a base port under the '
  + 'Port and Harbour Act — the only one in western Japan. The Hibikinada offshore wind farm is '
  + 'planned at approximately **220 MW** (25 turbines of 9.6 MW), developed by Hibiki Wind Energy. '
  + 'Since 2011 the Green Energy Port Hibiki project has been assembling a four-function cluster: '
  + 'import/export, turbine marshalling, O&M, and related manufacturing.'));
B(bullet('**Hydrogen.** The Fukuoka Prefecture Hydrogen Hub Council brings together around 30 '
  + 'organisations, chaired by the Governor of Fukuoka, with the Mayor of Kitakyushu, Kyushu Electric '
  + 'Power, Saibu Gas and Nippon Steel as vice-chairs — aiming at a commercial hydrogen supply chain '
  + 'centred on the Hibikinada waterfront.'));
B(bullet('**Circular industry.** Eco-Town firms recycle PV panels at a recovery rate above 99% and '
  + 'return aluminium, glass, silver and copper to the market; second-life EV batteries are being '
  + 'brought into the same system.'));
B(bullet('**Partnership agreements as the city’s main instrument.** April 2022 with Izutsuya, '
  + 'Toyota Motor Kyushu, Kyushu Electric Power and Kyushu Institute of Technology; October 2022 with '
  + 'Daikin, Denso Kyushu and Yanekara (a University of Tokyo spin-out). The city’s stated '
  + 'ambition is to act as a *hub* — it convenes, demonstrates and certifies, because it cannot invest '
  + 'at scale.'));

B(H2('6  A critical reading — where I think the case actually stands'));
B(P('Everything above is what the city presents, and most of it is genuinely working. My own research '
  + 'concerns the longer record, and it is less flattering. Kitakyushu has run an '
  + 'environment-and-knowledge-led growth strategy for three decades, and the city’s gross product '
  + 'has nonetheless grown slowly relative to other designated cities. The transition has been real as '
  + 'environmental policy and weak as economic policy.'));
B(P('I examined this through a growth-contribution analysis of Kitakyushu against Fukuoka City from '
  + '1980 to 2020, and through the record of the Science and Research Park and its agency FAIS. Two '
  + 'findings shape how I read the current decarbonisation programme.'));
B(H3('6.1  The industries that were targeted are not the industries that grew'));
B(tbl([1800, W - 1800], [
  ['Period', 'Industries the city and FAIS set out to grow'],
  ['1994–1998', 'Imaging and video, aerospace, airport-related industry'],
  ['1999–2003', 'Environment (recycling), information technology, welfare, housing, lifestyle and '
    + 'culture, biotechnology'],
  ['2004–2008', 'Environment (recycling, eco-premium products, biomass), information technology, '
    + 'welfare devices, biotechnology, next-generation robotics, semiconductors'],
  ['2019', 'Semiconductors, car electronics, information and communication (system LSI), robotics'],
], { header: true }));
B(P('What in fact contributed most to the city’s product in the most recent period was health care '
  + 'and social work, real estate, and miscellaneous manufacturing. The target list was rewritten '
  + 'roughly every five years, largely tracking whichever industry national programmes were funding at '
  + 'the time; the productive structure underneath it moved towards lower-value, locally-consumed '
  + 'services.'));
B(H3('6.2  Three explanations I am testing'));
B(numbered('**The success criterion was commercialisation, not value added.** FAIS assessed projects '
  + 'on technical feasibility, the prospect of getting to a product, and employment — not on revenue '
  + 'or contribution to the city’s economy. A project that reached market counted as a success '
  + 'even if it never grew. National schemes attached to the same projects *did* carry revenue '
  + 'targets, and the robotics programme, which used them, is one of the areas that actually '
  + 'expanded.'));
B(numbered('**There was never enough private capital.** No meaningful venture capital layer formed '
  + 'locally. Regional banks fund company formation rather than R&D, so FAIS ended up financing '
  + 'research from its own budget — which is bounded by a shrinking municipal balance sheet. The '
  + 'comparison I use is Pittsburgh, where university-led renewal was carried by roughly a billion '
  + 'dollars a year of investment and where venture investors brought a commercial judgement into the '
  + 'decision to fund.'));
B(numbered('**The supporting services never agglomerated.** A technology cluster needs accountants, '
  + 'patent and corporate lawyers, marketing, and investors within reach. Around the Science and '
  + 'Research Park these are thin, and the city has substituted for them by providing the services '
  + 'itself. That keeps the cluster running but caps it at the scale a municipality can fund, and '
  + 'removes the competition that would improve the services themselves.'));
B(box('Why this matters for our conversation', [
  'If that diagnosis is right, the third-party-ownership model in Section 5 is the most interesting '
  + 'thing Kitakyushu has done in thirty years — not because of the megawatts, which are small, but '
  + 'because it is the first instrument that puts *private balance sheets* behind municipal '
  + 'decarbonisation instead of asking a declining municipal budget to carry it. The unresolved '
  + 'question is whether buying services rather than owning assets leaves the city with any lasting '
  + 'capability, or whether it has simply outsourced its energy transition. That is exactly the '
  + 'question I want to put to a city that took the opposite route.',
]));
B(caveat('Section 6 reports work in progress. The contribution analysis is not yet published and the '
  + 'figures behind it are still being revised; I present the structural argument here rather than the '
  + 'numbers, and I would rather be corrected than believed.'));

B(H2('7  What I would like to ask you'));
B(Q(1, 'The shipyard.', 'When Kockums wound down, who absorbed the cost — the municipal balance sheet, '
  + 'the state, or private developers? Kitakyushu never had a single closure event; its steel industry '
  + 'shrank slowly and is still there, which may have made the transition politically easier and '
  + 'economically slower.'));
B(Q(2, 'What does the City of Malmö still own?', 'Kitakyushu has no municipal utility. It has a stake '
  + 'in a local electricity retailer and it owns buildings — that is nearly the whole lever. What is '
  + 'Malmö’s ownership position today across district heating, water, and housing, and how much of '
  + 'your climate programme depends on it?'));
B(Q(3, 'Municipal capital, regulation, or procurement?', 'Of Malmö’s climate targets, roughly what '
  + 'share is delivered by the city spending its own money, by regulating others, and by using its '
  + 'purchasing power? I would like to compare the mix rather than the targets.'));
B(Q(4, 'The emissions you do not control.', 'Two-thirds of Kitakyushu’s emissions belong to private '
  + 'industry. Does Malmö face the same asymmetry — and if so, what has actually worked: agreements, '
  + 'planning conditions, procurement, or public pressure?'));
B(Q(5, 'Third-party ownership under EU procurement law.', 'Do you use PPA or energy-service models on '
  + 'municipal buildings? In Japan the model’s weak point is that a long service contract is hard '
  + 'to reconcile with annual budgeting; I would expect EU procurement rules to create a different but '
  + 'comparable friction.'));
B(Q(6, 'Cooperating with your neighbours.', 'Kitakyushu bundled 18 municipalities to make its '
  + 'programme large enough to attract private bidders. How does Malmö achieve scale with Lund, Burlöv '
  + 'and Region Skåne — and does joint procurement work in practice or mainly on paper?'));
B(Q(7, 'The workforce.', 'What became of the shipyard workers? Japanese transitions are often '
  + 'cushioned by redeployment inside the same corporate group, which protects individuals but slows '
  + 'structural change. I would like to know whether Sweden’s separation of employment protection '
  + 'from the individual firm produced a faster or a more painful adjustment — and which you would '
  + 'choose again.'));
B(Q(8, 'Ownership forms.', 'I am also visiting Lolland and its holding company LOKE, and studying '
  + 'Danish foundation-owned enterprises. Does Malmö have anything comparable — a vehicle that lets '
  + 'public purposes be pursued with private capital and a long time horizon?'));

B(H2('8  Sources'));
B(srcItem('City of Kitakyushu, Environment Bureau, Renewable Energy Promotion Division, '
  + '「脱炭素先行地域にかかる北九州市の取り組み」 (*Kitakyushu’s initiatives concerning the '
  + 'Decarbonisation Leading Area*), presentation, 22 September 2023. — Principal source for '
  + 'Sections 4 and 5.'));
B(srcItem('City of Kitakyushu, *Green Energy Port Hibiki* — 「グリーンエネルギーポートひびき」, project '
  + 'brochure. — Offshore wind cluster and base port designation.'));
B(srcItem('City of Kitakyushu, *Next Generation Energy Kitakyushu*, energy park and municipal '
  + 'renewable energy map. — Installed generation in the Hibikinada district.'));
B(srcItem('Shinryo Co., Ltd. and Recycle Tech Co., Ltd., company profiles and PV panel recycling '
  + 'materials, 2022. — Eco-Town circular industry.'));
B(srcItem('Ørsted Japan K.K., company profile. — Offshore wind developer active in the Japanese '
  + 'market.'));
B(srcItem('Author’s own research on Kitakyushu’s industrial structure and the Science and '
  + 'Research Park (FAIS), drawing on *Kitakyushu City Economic Accounts* (1996–2020) and FAIS '
  + 'programme documents. Unpublished; Section 6 only.'));
B(srcItem('Background on the pollution period and the 1963 municipal merger is drawn from the standard '
  + 'literature on Kitakyushu’s post-war history and is not attributable to the documents above.'));

// ---------------------------------------------------------------- write

(async () => {
  for (const [name, children, title] of [
    ['japan-transition-report.docx', report, 'Financing the transition — nine Japanese cases'],
    ['kitakyushu-briefing.docx', brief, 'Kitakyushu: from steel and smog to a decarbonisation front-runner'],
  ]) {
    const buf = await Packer.toBuffer(makeDoc(title, children));
    fs.writeFileSync(name, buf);
    console.log(`${name}  ${(buf.length / 1024).toFixed(0)} KB`);
  }
})();
