const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, BorderStyle, HeadingLevel,
  VerticalAlign, Footer, PageNumber, convertMillimetersToTwip,
} = require('docx');

const MINCHO = { name: 'ＭＳ 明朝', eastAsia: 'ＭＳ 明朝', ascii: 'Times New Roman', hAnsi: 'Times New Roman' };
const GOTHIC = { name: 'ＭＳ ゴシック', eastAsia: 'ＭＳ ゴシック', ascii: 'Arial', hAnsi: 'Arial' };

const SIZE = 21;        // 10.5pt
const CONTENT = 9638;   // A4 minus 20mm margins each side
const HEAD_FILL = 'EDEDED';

// ---------- helpers ----------

// Body paragraph, indented one full-width character.
const body = (text, opts = {}) => new Paragraph({
  spacing: { after: 100, line: 290, lineRule: 'auto' },
  indent: { firstLine: SIZE * 10 },
  ...opts,
  children: [new TextRun({ text, font: MINCHO, size: SIZE })],
});

// Numbered section heading with a rule underneath.
const h1 = (text) => new Paragraph({
  spacing: { before: 280, after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '808080', space: 4 } },
  children: [new TextRun({ text, font: GOTHIC, size: 24, bold: true })],
});

// (1)-style sub-heading.
const h2 = (text) => new Paragraph({
  spacing: { before: 170, after: 70 },
  children: [new TextRun({ text, font: GOTHIC, size: SIZE, bold: true })],
});

const spacer = (after = 120) => new Paragraph({ spacing: { after }, children: [] });

// One paragraph inside a table cell.
const cellText = (text, { bold = false, align = AlignmentType.LEFT, font = MINCHO } = {}) =>
  new Paragraph({
    alignment: align,
    spacing: { before: 30, after: 30, line: 250, lineRule: 'auto' },
    children: [new TextRun({ text, font, size: 19, bold })],
  });

const cell = (text, width, opts = {}) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined,
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 60, bottom: 60, left: 108, right: 108 },
  children: (Array.isArray(text) ? text : [text]).map((t) =>
    cellText(t, { bold: opts.bold, align: opts.align, font: opts.font })),
});

// Build a table from a header row plus body rows.
const table = (widths, header, rows, opts = {}) => new Table({
  columnWidths: widths,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  rows: [
    ...(header ? [new TableRow({
      tableHeader: true,
      cantSplit: true,
      children: header.map((t, i) =>
        cell(t, widths[i], { bold: true, fill: HEAD_FILL, align: AlignmentType.CENTER, font: GOTHIC })),
    })] : []),
    ...rows.map((r, ri) => {
      const isTotal = opts.totalRow && ri === rows.length - 1;
      return new TableRow({
        cantSplit: true,
        children: r.map((t, i) =>
          cell(t, widths[i], {
            fill: (opts.labelFill && i === 0) || isTotal ? HEAD_FILL : undefined,
            bold: (opts.labelBold && i === 0) || isTotal,
            align: (opts.centerCols || []).includes(i) ? AlignmentType.CENTER : undefined,
            font: (opts.labelFill && i === 0) || isTotal ? GOTHIC : undefined,
          })),
      });
    }),
  ],
});

// ---------- content ----------

const children = [];

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 80 },
  children: [new TextRun({ text: '物品購入理由書', font: GOTHIC, size: 32, bold: true })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 320 },
  children: [new TextRun({
    text: 'ネットワーク接続ストレージ（NAS）　UGREEN NASync DXP4800 Plus',
    font: GOTHIC, size: 21,
  })],
}));

children.push(table([2400, 7238], null, [
  ['申請日', '　　　　年　　月　　日'],
  ['所属・職名', '　'],
  ['氏名', '　'],
  ['経費区分', '学内個人研究費'],
  ['品名', 'ネットワーク接続ストレージ（NAS）'],
  ['型番', 'UGREEN NASync DXP4800 Plus（4ベイ／ディスクレス）'],
  ['数量', '1台（本体のみ）'],
  ['見積金額', '　　　　　　　円（本体のみ。内訳は「7. 経費内訳」参照）'],
], { labelFill: true }));

// 1
children.push(h1('1. 購入の目的'));
children.push(body('研究資料・研究データを一元的に集約し、AIツールが継続的に参照・処理できる常時稼働型の研究データ基盤を構築するため、ネットワーク接続ストレージ（NAS）を導入する。'));
children.push(new Paragraph({
  spacing: { after: 100, line: 290, lineRule: 'auto' },
  indent: { firstLine: SIZE * 10 },
  children: [
    new TextRun({ text: '本申請は単独のストレージ設備の追加ではなく、', font: MINCHO, size: SIZE }),
    new TextRun({ text: '既に学内個人研究費によって導入しているAIツール（Claude Code、Codex等）と研究データを接続し、これまでの研究環境整備への投資を有効に活用するための情報基盤整備', font: MINCHO, size: SIZE, bold: true }),
    new TextRun({ text: 'として位置づけられる。', font: MINCHO, size: SIZE }),
  ],
}));

// 2
children.push(h1('2. 現状と課題'));
children.push(body('現在、学内個人研究費を利用してClaude Code、Codex等のAIツールを導入し、研究資料の整理、データ処理、プログラム作成、分析作業等への活用を進めている。しかし、運用上、以下の課題が生じている。'));
children.push(h2('(1) 処理能力が個々の研究用PCの環境に依存している'));
children.push(body('現在はAIツールを個々の研究用PC上で利用しているため、処理可能なデータ量や稼働時間が、各端末の保存容量・稼働状況に依存する。長時間を要する処理は、PCの起動状態やスリープ状態に左右され、継続的な実行が難しい。'));
children.push(h2('(2) AIエージェントが継続的に参照できる研究データ領域が確保できない'));
children.push(body('研究資料がPC本体・外付けHDD・クラウドストレージ等に分散しているため、AIエージェントが一貫して参照できる研究データ領域を確保することが難しい。同一資料の複数版が各所に存在し、参照先の同定にも支障が生じている。'));
children.push(h2('(3) 研究データの保全体制が個別端末任せになっている'));
children.push(body('現状では、研究データの複製・世代管理が各端末および各サービスに委ねられており、機器故障や誤操作に対する冗長性が確保されていない。'));

// 3
children.push(h1('3. 導入により実現すること'));
children.push(body('NASに研究資料を体系的に集約するとともに、AIエージェントが継続的に参照・処理できる常時稼働型の環境を構築する。これにより、以下が実現される。'));
children.push(h2('(1) AIツールの「その都度利用」から「継続的な研究支援環境」への発展'));
children.push(body('既に導入しているClaude Code、Codex等を、個々のPC上でその都度利用するツールから、蓄積された研究資料を継続的に整理・検索・処理する研究支援環境へと発展させることができる。具体的には、資料の分類・索引化、横断検索、定型的なデータ整形・集計処理等を、PCの稼働状況に依存せず継続的に実行できる。'));
children.push(h2('(2) 研究データの参照先の一元化'));
children.push(body('分散していた研究資料を単一の保存領域に集約することで、AIツールおよび研究者自身がアクセスする参照先が一元化され、版の同定や資料探索に要する時間を削減できる。'));
children.push(h2('(3) 処理内容に応じた外部サービスとローカル環境の使い分け'));
children.push(body('ローカルLLMを併用することで、用途や資料の性質に応じて、外部のAIサービスを利用する処理と、研究者の管理下にあるローカル環境内で完結させる処理を使い分けることが可能となる。外部送信が適切でない資料については、ネットワーク外部にデータを送出せずに処理を行うことができ、研究データの管理上も有益である。'));
children.push(h2('(4) 研究データの保全'));
children.push(body('RAID構成による冗長化と定期的な世代バックアップにより、機器故障・誤削除に対する研究データの保全体制を確立する。'));

// 4
children.push(h1('4. 本機種を選定した理由'));
children.push(new Paragraph({
  spacing: { after: 140, line: 290, lineRule: 'auto' },
  indent: { firstLine: SIZE * 10 },
  children: [
    new TextRun({ text: '本用途では、単なる保存容量の確保ではなく、', font: MINCHO, size: SIZE }),
    new TextRun({ text: 'AI処理を継続的に実行できる計算能力・拡張性・転送速度を備えた常時稼働機', font: MINCHO, size: SIZE, bold: true }),
    new TextRun({ text: 'であることが要件となる。UGREEN NASync DXP4800 Plusは、以下の点でこれらの要件を満たす。', font: MINCHO, size: SIZE }),
  ],
}));
children.push(table([2100, 2400, 5138],
  ['要件', '本機種の仕様', '選定理由'],
  [
    ['AI処理・コンテナ実行に耐えるCPU', '第12世代 Intel Pentium Gold 8505（5コア）', '一般的な低性能NAS用SoCと異なり、x86系プロセッサを搭載しており、AIエージェントの常駐実行、コンテナ（Docker）による処理環境の構築、ローカルLLMの小規模モデル推論に対応できる。'],
    ['大容量メモリへの拡張', 'DDR5 8GB搭載、最大64GBまで増設可能', 'ローカルLLMの実行および複数処理の並行実行にはメモリ容量が制約となるため、後日の増設余地があることが必須要件である。'],
    ['高速なネットワーク転送', '10ギガビットEthernet（10GbE）＋2.5GbE の2系統', '大容量の研究データを研究用PCとNAS間でやり取りする際、通常の1GbE接続では転送が処理時間の律速となる。10GbE対応により、実質的にローカルディスクに近い速度で作業できる。'],
    ['高速キャッシュ／作業領域', 'M.2 NVMe SSDスロット×2', 'AIによる索引化・検索処理では多数の小容量ファイルへの高速アクセスが必要となるため、HDDとは別にSSD領域を確保できることが有効である。'],
    ['保存容量と冗長性', '3.5インチドライブベイ×4、RAID構成に対応（最大96TB）', '4ベイ構成によりRAID5等の冗長構成が組めるため、1台のドライブ故障時にも研究データを保全できる。2ベイ機ではミラーリングに限られ、容量効率が低い。'],
    ['常時稼働に適した消費電力', '低消費電力設計（24時間稼働を前提とした構成）', 'AIエージェントによる継続的な処理を行うため、常時稼働が前提となる。研究用PCを常時起動する運用と比較して、消費電力・機器寿命の面で合理的である。'],
  ]));
children.push(spacer(160));
children.push(h2('本機種の価格上の妥当性'));
children.push(new Paragraph({
  spacing: { after: 100, line: 290, lineRule: 'auto' },
  indent: { firstLine: SIZE * 10 },
  children: [
    new TextRun({ text: '同等仕様（x86系CPU・10GbE・4ベイ・メモリ増設可）の他社製NASは、本体のみで15万円〜20万円程度となる場合が多い。本機種の国内実勢価格は9万円台〜10万円程度であり、', font: MINCHO, size: SIZE }),
    new TextRun({ text: '必要な仕様を満たす選択肢の中で最も低廉な部類に属する', font: MINCHO, size: SIZE, bold: true }),
    new TextRun({ text: '。', font: MINCHO, size: SIZE }),
  ],
}));

// 5
children.push(h1('5. 代替手段との比較検討'));
children.push(body('本申請に先立ち、以下の代替手段を検討したが、いずれも本研究用途の要件を満たさないため、NASの導入が必要と判断した。'));
children.push(table([2600, 7038],
  ['代替手段', '不採用の理由'],
  [
    ['外付けHDD／SSDの増設', 'PCに接続された端末に処理が依存するため、常時稼働型の処理環境を構築できない。複数端末からの同時参照、RAIDによる冗長化にも対応しない。'],
    ['クラウドストレージの容量追加', 'AIエージェントが継続的にファイルシステムとして参照する用途には適さず、大容量データの反復的な読み書きで転送速度・API制限の制約を受ける。また、年額の継続費用が発生し、長期的には本機器の購入費を上回る。外部サービスであるため、外部送信が適切でない資料の取り扱いにも制約がある。'],
    ['学内共用ストレージの利用', '保存領域としては利用可能だが、利用者側で常駐プロセス（AIエージェント、コンテナ環境）を実行することができず、本申請の目的である「継続的な処理環境」を構築できない。'],
    ['高性能な研究用PCの追加購入', '常時稼働を前提とすると消費電力・設置スペースの面で不利であり、RAIDによるデータ保全機能も標準では備えない。同等の目的を達成するには本機器より高額となる。'],
  ]));

// 6
children.push(h1('6. 研究費支出の妥当性'));
[
  ['既存投資の有効活用', '既に学内個人研究費で導入済みのAIツールの利用効率を高めるための基盤整備であり、追加投資に対する効果が既存資産に及ぶ。'],
  ['費用対効果', '資料整理・データ処理・集計作業等に要していた作業時間の削減が見込まれ、研究の実施時間を確保できる。クラウドストレージの継続課金と比較しても、複数年の使用を前提とすれば費用面で優位である。'],
  ['継続的な使用', '購入後は研究期間を通じて継続的に使用する設備であり、単年度の消耗的支出ではない。'],
  ['研究データの適正管理', '研究データの冗長化保存および外部送信の抑制により、研究データ管理上の要請にも資する。'],
].forEach(([label, text], i) => {
  children.push(new Paragraph({
    spacing: { after: 100, line: 290, lineRule: 'auto' },
    indent: { left: SIZE * 10, hanging: SIZE * 10 },
    children: [
      new TextRun({ text: `(${i + 1}) `, font: MINCHO, size: SIZE }),
      new TextRun({ text: `${label}：`, font: GOTHIC, size: SIZE, bold: true }),
      new TextRun({ text, font: MINCHO, size: SIZE }),
    ],
  }));
});

// 7
children.push(h1('7. 経費内訳'));
children.push(table([1900, 3138, 900, 1850, 1850],
  ['品目', '仕様', '数量', '単価', '金額'],
  [
    ['NAS本体', 'UGREEN NASync DXP4800 Plus（4ベイ）', '1台', '', ''],
    ['合計', '', '', '', ''],
  ], { centerCols: [2, 3, 4], totalRow: true }));
children.push(spacer(120));
children.push(new Paragraph({
  spacing: { after: 60, line: 280, lineRule: 'auto' },
  indent: { left: SIZE * 10, hanging: SIZE * 10 },
  children: [new TextRun({ text: '※ 本申請は本体のみの購入であり、内蔵HDD等の記録媒体は本申請に含まない。', font: MINCHO, size: 19 })],
}));
children.push(new Paragraph({
  spacing: { after: 60, line: 280, lineRule: 'auto' },
  indent: { left: SIZE * 10, hanging: SIZE * 10 },
  children: [new TextRun({ text: '※ 記録媒体については既存の資産を用いて運用する。', font: MINCHO, size: 19 })],
}));

// 8
children.push(h1('8. 設置・運用・管理計画'));
children.push(table([2200, 7438], null, [
  ['設置場所', '研究室内（施錠管理下）'],
  ['管理責任者', '　'],
  ['ネットワーク接続', '学内ネットワークに接続し、外部からの直接アクセスは行わない設定とする。'],
  ['アクセス制御', '利用者アカウントおよびアクセス権限を設定し、研究データへのアクセスを管理する。'],
  ['バックアップ', 'RAID構成による冗長化に加え、重要データについては別媒体への定期バックアップを行う。'],
  ['データ管理', '本学の研究データ管理および情報セキュリティに関する規程に従って運用する。個人情報を含む資料については、外部AIサービスへ送信せず、本機器内のローカル環境で処理する運用とする。'],
  ['使用期間', '導入後、研究期間を通じて継続使用する。'],
], { labelFill: true }));

children.push(spacer(160));
children.push(new Paragraph({
  spacing: { before: 200, line: 290, lineRule: 'auto' },
  indent: { firstLine: SIZE * 10 },
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: '808080', space: 8 } },
  children: [new TextRun({
    text: '以上のとおり、本機器の導入は、既に導入済みのAIツールと研究データを接続し、研究環境整備への既存投資を有効に活用するために必要な情報基盤整備であるため、購入を申請する。',
    font: MINCHO, size: SIZE,
  })],
}));

// ---------- document ----------

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: MINCHO, size: SIZE }, paragraph: { spacing: { line: 290, lineRule: 'auto' } } },
    },
  },
  sections: [{
    properties: {
      page: {
        margin: {
          top: convertMillimetersToTwip(22),
          bottom: convertMillimetersToTwip(20),
          left: convertMillimetersToTwip(20),
          right: convertMillimetersToTwip(20),
        },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: MINCHO, size: 18 })],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = process.argv[2];
  fs.writeFileSync(out, buf);
  console.log('wrote', out, buf.length, 'bytes');
});
