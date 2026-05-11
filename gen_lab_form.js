const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign
} = require('C:/Users/user/AppData/Roaming/npm/node_modules/docx');
const fs = require('fs');

// A4, 0.4in (576 DXA) margins → content width = 11906 - 1152 = 10754
const CW = 10754;
const LW = 1300, FW = 4077;  // 4-col: (LW+FW)*2 = 10754
// Equipment cols (8): sum must = 10754
const EQ = [300, 3280, 1480, 620, 1180, 780, 780, 2334];
// 300+3280+1480+620+1180+780+780+2334 = 10754 ✓
const SG = [2688, 2689, 2688, 2689]; // 4-equal sign cols, sum=10754

const F = 'Microsoft JhengHei';
const NAVY = '0f1f3d', NAV2 = '1e3a6e', GOLD = 'c9a84c';
const LBLC = 'd5cdb8', STRC = 'f0ece2', WHTC = 'FFFFFF';

const bs = { style: BorderStyle.SINGLE, size: 4, color: '9a9a9a' };
const bn = { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' };
const AB = { top: bs, bottom: bs, left: bs, right: bs };
const NB = { top: bn, bottom: bn, left: bn, right: bn };

const T = (text, sz = 16, bold = false, color = '222222') =>
  new TextRun({ text, font: F, size: sz, bold, color });

const P = (runs, align) => new Paragraph({
  spacing: { before: 0, after: 0, line: 240 },
  ...(align ? { alignment: align } : {}),
  children: Array.isArray(runs) ? runs : [runs]
});

const TC = (w, children, { fill = WHTC, cs, mt = 44, mb = 44, ml = 88, mr = 72, noB = false } = {}) =>
  new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: noB ? NB : AB,
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: mt, bottom: mb, left: ml, right: mr },
    verticalAlign: VerticalAlign.CENTER,
    ...(cs && cs > 1 ? { columnSpan: cs } : {}),
    children: Array.isArray(children) ? children : [children]
  });

// Label cell
const lbl = (text, w, req = false) =>
  TC(w, P([T(text, 15, true, '1a2a4a'), ...(req ? [T('*', 14, true, 'c0392b')] : [])], AlignmentType.CENTER), { fill: LBLC });

// Input cell
const inp = (w, opts = {}) =>
  TC(w, P(opts.hint ? [T(opts.hint, 13, false, 'bbbbbb')] : [T('')]), opts);

// Section header (spans all cols via cs)
const hdr = (text, badge, nCols) =>
  new TableRow({ children: [
    TC(CW, P([T('▏ ', 17, true, GOLD), T(text, 17, true, WHTC), ...(badge ? [T(`  ${badge}`, 13, false, '99bbdd')] : [])]),
      { fill: NAV2, cs: nCols, ml: 120, mt: 58, mb: 58 })
  ]});

// Checkbox row (spans all cols)
const chkRow = (items, nCols) => {
  const runs = [];
  items.forEach((item, i) => {
    if (i > 0) runs.push(T('    '));
    runs.push(T('□ ', 16, false, '666666'));
    runs.push(T(item, 15, false, '333333'));
  });
  return new TableRow({ children: [TC(CW, P(runs), { fill: STRC, cs: nCols, ml: 120 })] });
};

const TR = (...cells) => new TableRow({ children: cells });
const SP = () => new Paragraph({ spacing: { before: 0, after: 26 }, children: [T('', 2)] });

// ── TITLE ──────────────────────────────────────────────────────────────────
const tTitle = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
  rows: [new TableRow({ children: [
    TC(CW, [
      P([T('○○大學　○○學系', 14, false, '88aacc')], AlignmentType.CENTER),
      P([T('實 驗 室 借 用 暨 器 材 申 請 表', 24, true, WHTC)], AlignmentType.CENTER),
      P([T('表單編號：LAB-_______________　　申請日期：____年____月____日', 13, false, '888888')], AlignmentType.CENTER),
    ], { fill: NAVY, mt: 90, mb: 90 })
  ]})]
});

// ── 壹、申請人 + 指導老師 ──────────────────────────────────────────────────
const tInfo = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: [LW, FW, LW, FW],
  rows: [
    hdr('壹、申請人與指導老師資料', '（＊必填）', 4),
    TR(lbl('姓　　名', LW, true), inp(FW), lbl('學　　號', LW, true), inp(FW)),
    TR(lbl('系所 / 班級', LW, true), inp(FW), lbl('年　　級', LW), inp(FW)),
    TR(lbl('聯絡電話', LW, true), inp(FW), lbl('指導老師', LW, true), inp(FW)),
    TR(lbl('電子信箱', LW), inp(FW), lbl('老師電話', LW), inp(FW)),
  ]
});

// ── 貳、實驗室借用 ───────────────────────────────────────────────────────
const tBorrow = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: [LW, FW, LW, FW],
  rows: [
    hdr('貳、實驗室借用申請', '（學生填寫）', 4),
    TR(lbl('實驗室名稱', LW, true), inp(FW), lbl('實驗室編號', LW, true), inp(FW)),
    TR(lbl('借　用　日', LW, true), inp(FW), lbl('預計人　數', LW, true), inp(FW)),
    TR(lbl('借用時　段', LW, true), inp(FW, { hint: '自　　：　　　至　　：　　' }), lbl('課程 / 計畫', LW), inp(FW)),
    TR(lbl('使 用 目 的', LW, true), inp(FW + LW + FW, { cs: 3 })),
  ]
});

// ── 參、器材清單 ─────────────────────────────────────────────────────────
const eH = (text, w) => TC(w, P([T(text, 13, true, WHTC)], AlignmentType.CENTER), { fill: NAVY, ml: 30, mr: 30, mt: 52, mb: 52 });
const eC = (w, text = '') => TC(w, P([T(text, 14, false, '555555')], AlignmentType.CENTER), { ml: 28, mr: 28, mt: 36, mb: 36 });
const eRow = n => new TableRow({ children: EQ.map((w, i) => eC(w, i === 0 ? String(n).padStart(2, '0') : '')) });

const tEq = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: EQ,
  rows: [
    hdr('參、器材借用清單', '（借出 / 歸還欄由設備組核簽）', 8),
    new TableRow({ children: [
      eH('項次', EQ[0]), eH('器材 / 設備名稱', EQ[1]), eH('型號規格', EQ[2]),
      eH('數量', EQ[3]), eH('設備號', EQ[4]),
      eH('借出 ✓', EQ[5]), eH('歸還 ✓', EQ[6]), eH('備　　註', EQ[7]),
    ]}),
    ...[1, 2, 3, 4, 5, 6].map(eRow),
  ]
});

// ── 肆、設備組確認（借用前）──────────────────────────────────────────────
const sigTC = (w) => TC(w, P([T('')]), { mt: 86, mb: 86 });

const tOffice = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: [LW, FW, LW, FW],
  rows: [
    hdr('肆、設備組確認（借用前）', '（設備組填寫）', 4),
    chkRow(['確認申請時段實驗室無人使用', '固定設備運作正常', '安全設備齊全', '已登記排程系統'], 4),
    TR(lbl('確認人員', LW), inp(FW), lbl('確認日期時間', LW), inp(FW)),
    TR(lbl('設備組簽章', LW), sigTC(FW), lbl('主 管 核 章', LW), sigTC(FW)),
  ]
});

// ── 伍、借出點交確認 ──────────────────────────────────────────────────────
const tLend = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: [LW, FW, LW, FW],
  rows: [
    hdr('伍、借出點交確認', '（借用當日）', 4),
    chkRow(['門窗電力狀況正常', '器材清單核對完畢', '鑰匙已交付（No.______）', '安全規則已告知'], 4),
    TR(lbl('點交時間', LW), inp(FW), lbl('點交人員', LW), inp(FW)),
    TR(lbl('設備組簽章', LW), sigTC(FW), lbl('學生確認簽章', LW), sigTC(FW)),
  ]
});

// ── 陸、歸還確認（指導老師執行）──────────────────────────────────────────
const tReturn = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: [LW, FW, LW, FW],
  rows: [
    hdr('陸、歸還確認（指導老師執行）', '（實驗完畢）', 4),
    chkRow(['器材全數歸還', '無損壞缺件', '實驗室清潔整理', '水電已關閉', '廢棄物依規定處置'], 4),
    TR(lbl('實際結束時間', LW), inp(FW), lbl('損壞 / 遺失說明', LW), inp(FW)),
    TR(lbl('指導老師簽章', LW), sigTC(FW), lbl('學 生 簽 章', LW), sigTC(FW)),
    TR(lbl('設備組人員簽章', LW), sigTC(FW), lbl('鑰匙歸還確認', LW), sigTC(FW)),
  ]
});

// ── 注意事項 + 核章 ──────────────────────────────────────────────────────
const notes = [
  '申請須於借用日三個工作天前提出，取得指導老師核簽後送設備組審核。',
  '借用前須親洽設備組確認實驗室無人使用，未確認者不得擅自進入。',
  '指導老師須全程可聯繫；實驗結束後親自清點設備並檢查實驗室。',
  '器材損壞依本系賠償標準辦理；廢液依環保規定分類處置。',
];

const tNotes = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
  rows: [
    hdr('注意事項', '', 1),
    new TableRow({ children: [
      TC(CW,
        notes.map((n, i) => P([T(`${i + 1}. `, 13, true, '555555'), T(n, 13, false, '333333')])),
        { fill: STRC, ml: 130, mr: 130, mt: 66, mb: 66 }
      )
    ]})
  ]
});

const signC = (label, w) => TC(w, [
  P([T(label, 14, true, '1a2a4a')], AlignmentType.CENTER),
  P([T('')]),
  P([T('─ 簽  章 ─', 12, false, 'cccccc')], AlignmentType.CENTER),
], { fill: STRC, mt: 58, mb: 58 });

const tSign = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: SG,
  rows: [
    hdr('核　　章', '', 4),
    new TableRow({ children: [signC('申請人', SG[0]), signC('指導老師', SG[1]), signC('設備組組長', SG[2]), signC('系　主　任', SG[3])] })
  ]
});

// ── DOCUMENT ─────────────────────────────────────────────────────────────
const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 576, right: 576, bottom: 576, left: 576 }
      }
    },
    children: [
      tTitle, SP(),
      tInfo, SP(),
      tBorrow, SP(),
      tEq, SP(),
      tOffice, SP(),
      tLend, SP(),
      tReturn, SP(),
      tNotes, SP(),
      tSign
    ]
  }]
});

Packer.toBuffer(doc)
  .then(buf => {
    fs.writeFileSync('C:\\Users\\user\\allen\\lab-form.docx', buf);
    console.log('Done: lab-form.docx');
  })
  .catch(err => { console.error(err); process.exit(1); });
