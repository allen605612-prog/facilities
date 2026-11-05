// 第5章 水與陸地 綜合練習（B4，版式仿正心中學段考卷）：題庫在 _ch5_combined_data.js
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, PageBreak, ImageRun, Footer, PageNumber } = require('docx');
// 用法：node _gen_ch5_combined.js 輸出.docx [題庫.js] [標題]（省略時為第5章合併卷）
const parts = require('./' + (process.argv[3] || '_ch5_combined_data.js'));
const TITLE = process.argv[4] || '正心中學115學年度國三上第5章水與陸地綜合練習';

const FONT = { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '標楷體' };
const LINE = 320;
const t = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 24, bold: o.bold });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [t(runs)], spacing: { before: o.before || 0, after: o.after ?? 0, line: LINE }, alignment: o.align ?? AlignmentType.JUSTIFIED, indent: o.indent, keepNext: o.keepNext, keepLines: true });
const FIG = 'C:\\Users\\user\\allen\\_exam_fig\\';
const img = (f, w, h) => new Paragraph({ indent: { left: IND }, spacing: { before: 40, after: 40 }, keepNext: true, children: [new ImageRun({ type: 'png', data: fs.readFileSync(FIG + f), transformation: { width: w, height: h } })] });

const IND = 480;                       // 題號懸掛縮排
const HANG = { left: IND, hanging: IND }, BODY = { left: IND };
const L = ['A', 'B', 'C', 'D'];
const cb = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const borders = { top: cb, bottom: cb, left: cb, right: cb };
// 表格：無底色，首列與首欄粗體，靠左對齊題目內文
const mkTable = (rows, widths, center) => new Table({ alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT, indent: center ? undefined : { size: IND, type: WidthType.DXA },
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
  rows: rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, ci) => new TableCell({
    borders, width: { size: widths[ci], type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ keepNext: true, spacing: { after: 0, line: 280 }, alignment: widths.length === 2 && ci === 1 && ri > 0 ? AlignmentType.LEFT : AlignmentType.CENTER, children: [t(c, { bold: ri === 0 || ci === 0 })] })] })) })) });
const optRuns = o => o.map((s, k) => t(`${k ? '  ' : ''}(${L[k]}) ${s}`));

const body = [];
body.push(p([t(TITLE, { size: 36 })], { align: AlignmentType.LEFT, after: 60 }));
body.push(p([t('單一選擇題（共50題，每題2分，共100分）')], { align: AlignmentType.LEFT, after: 120 }));

let n = 0; const all = [];
parts.forEach(part => {
  const off = n;
  part.start = n + 1; part.end = n + part.items.length;
  body.push(p([t(`${part.name}（${part.start}-${part.end}題）`, { size: 27, bold: true })], { align: AlignmentType.LEFT, before: off ? 160 : 0, after: 80, keepNext: true }));
  part.items.forEach(x => {
    n++; all.push(x);
    if (x.sec) body.push(p([t(x.sec, { bold: true })], { align: AlignmentType.LEFT, before: 60, after: 40, keepNext: true }));
    if (x.intro) body.push(p([t(x.intro.replace(/第(\d+)～(\d+)題/, (_, a, b) => `第${+a + off}～${+b + off}題`))], { before: 60, after: 40, keepNext: true }));
    if (x.fig) body.push(img(...x.fig));
    if (x.table && x.tableFirst) { body.push(mkTable(x.table, x.widths)); body.push(p('', { keepNext: true, after: 40 })); }
    const lines = x.q.split('\n');
    const stem = [t(`${n}.\t`), t(lines[0])];
    if (x.u) { stem.push(t(x.u, { bold: true })); stem.push(t(x.q2)); }
    const inline = lines.length === 1 && !(x.table && !x.tableFirst);
    if (inline) {                         // 選項直接接在題幹後面
      body.push(new Paragraph({ children: [...stem, t(' '), ...optRuns(x.o)], spacing: { after: 80, line: LINE }, alignment: AlignmentType.JUSTIFIED, indent: HANG, tabStops: [{ type: 'left', position: IND }], keepLines: true }));
      return;
    }
    body.push(new Paragraph({ children: stem, spacing: { after: 20, line: LINE }, alignment: AlignmentType.JUSTIFIED, indent: HANG, tabStops: [{ type: 'left', position: IND }], keepNext: true, keepLines: true }));
    lines.slice(1).forEach(s => body.push(p([t(s)], { indent: { left: IND + 240 }, keepNext: true })));
    if (x.table && !x.tableFirst) { body.push(p('', { keepNext: true, after: 20 })); body.push(mkTable(x.table, x.widths)); body.push(p('', { keepNext: true, after: 40 })); }
    body.push(p(optRuns(x.o), { indent: BODY, after: 80 }));
  });
});

// 答案與解析（教師用，另起一頁）
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(p([t('答案與解析（教師用）', { size: 36 })], { align: AlignmentType.LEFT, after: 120 }));
const rows = [];
for (let r = 0; r * 10 < all.length; r++) {
  const idx = [...Array(10).keys()].map(c => r * 10 + c);
  rows.push(['題號', ...idx.map(i => i < all.length ? String(i + 1) : '')], ['答案', ...idx.map(i => i < all.length ? all[i].a : '')]);
}
body.push(mkTable(rows, [1000, ...Array(10).fill(900)], true));
body.push(p('', { after: 80 }));
parts.forEach(part => {
  body.push(p([t(`${part.name}（${part.start}-${part.end}題）`, { size: 27, bold: true })], { align: AlignmentType.LEFT, before: 100, after: 40, keepNext: true }));
  for (let i = part.start - 1; i < part.end; i++)
    body.push(p([t(`${i + 1}.（${all[i].a}）`, { bold: true }), t(all[i].e)], { indent: { left: 800, hanging: 800 }, after: 40 }));
});

const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
  new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 }), new TextRun({ text: '(共', font: FONT, size: 20 }),
  new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 20 }), new TextRun({ text: '頁)', font: FONT, size: 20 })] })] });

const cnt = { A: 0, B: 0, C: 0, D: 0 }; all.forEach(x => cnt[x.a]++);
const doc = new Document({ styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  sections: [{ properties: { page: { size: { width: 14570, height: 20636 }, margin: { top: 1440, bottom: 1300, left: 1560, right: 1440, footer: 700 } } },  // B4（257×364 mm）
    footers: { default: footer }, children: body }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('ok', n, JSON.stringify(cnt)); });
