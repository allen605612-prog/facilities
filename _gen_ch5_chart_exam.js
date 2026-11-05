const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, PageBreak, TabStopType, ImageRun } = require('docx');
const FONT = { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '標楷體' };
const t = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 24, bold: o.bold });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [t(runs)], spacing: { after: o.after ?? 40, line: 340 }, alignment: o.align, indent: o.indent, tabStops: o.tabs, keepNext: o.keepNext });
const FIG = 'C:\\Users\\user\\allen\\_exam_fig\\';
const img = (f, w, h) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, keepNext: true, children: [new ImageRun({ type: 'png', data: fs.readFileSync(FIG + f), transformation: { width: w, height: h } })] });

// intro：題組說明（含表或圖），qs：題目
const G = [
  { intro: '下表為地球上各種水體所占的比例：', table: [['水體', '海水', '冰川、冰帽及永久雪', '地下水', '河流、湖泊', '其他'], ['比例', '97.2%', '2.15%', '0.63%', '0.009%', '0.011%']], widths: [1000, 1300, 2600, 1300, 1400, 1300],
    qs: [{ q: '若將海水以外的水都視為淡水，則冰川、冰帽及永久雪約占地球「淡水」總量的多少？', o: ['約2%', '約25%', '約77%', '約97%'], a: 'C', e: '淡水約占100%－97.2%＝2.8%，冰占淡水比例約2.15÷2.8≈77%。' }] },
  { intro: '下表為某海域海水中各種鹽類占總鹽類的質量百分比，已知此海域海水鹽度為35‰：', table: [['鹽類', '氯化鈉', '氯化鎂', '硫酸鎂', '其他'], ['占總鹽類質量比例', '78%', '10%', '6%', '6%']], widths: [2600, 1500, 1500, 1500, 1500],
    qs: [{ q: '取此海域海水10公斤完全蒸發，所得的氯化鎂約為多少公克？', o: ['3.5', '35', '273', '350'], a: 'B', e: '10公斤＝10000公克海水含鹽類10000×35/1000＝350公克，其中氯化鎂占10%，即35公克。' }] },
  { intro: '下圖為某盆地1951～2007年間地下水位與累計地層下陷量的變化，試回答第3～4題：', fig: ['fig1.png', 520, 276],
    qs: [
      { q: '政府為解決地層下陷，開始嚴格管制抽取地下水。根據圖中資料判斷，管制措施最可能從哪一個時間點開始實施？', o: ['西元1955年前後', '西元1965年前後', '西元1975年前後', '西元1995年前後'], a: 'C', e: '約1975年起地下水位由最低點開始回升，表示抽水量減少，最可能是開始管制的時間點。' },
      { q: '根據圖中資料，下列敘述何者正確？', o: ['地下水位回升後，地面也回升至原來的高度', '地下水位最低時，地層下陷的速度最慢', '西元1991年以後的下陷速度大於1961～1971年', '地下水位回升後，地層下陷趨緩，但已下陷的地面無法恢復'], a: 'D', e: '1975年後累計下陷量曲線趨於平緩但沒有回升，表示被壓縮的岩層孔隙無法恢復原狀。1961～1971年下陷約110公分，是下陷最快的時期。' }] },
  { intro: '下表為同一條河流甲、乙、丙三個測站的觀測資料：', table: [['測站', '平均流速（公尺／秒）', '河道寬度（公尺）', '河床主要沉積物'], ['甲', '0.4', '300', '泥'], ['乙', '2.5', '20', '礫石'], ['丙', '1.2', '80', '沙']], widths: [1200, 2800, 2400, 2400],
    qs: [{ q: '三個測站由上游至下游的排列順序，最可能為下列何者？', o: ['甲、丙、乙', '丙、乙、甲', '甲、乙、丙', '乙、丙、甲'], a: 'D', e: '上游流速快、河道窄、沉積物顆粒大；下游流速慢、河道寬、沉積物顆粒小。故順序為乙（上游）→丙→甲（下游）。' }] },
  { intro: '下圖為某河流出海口附近的沉積物剖面，甲層在最上方，各地層未經翻轉，試回答第6～7題：', fig: ['fig2.png', 300, 229],
    qs: [
      { q: '沉積哪一層時，河水的流速最快？', o: ['甲', '乙', '丙', '丁'], a: 'B', e: '流速越快，搬運能力越強，能沉積的顆粒越大。乙層為礫石，顆粒最大，表示當時流速最快。' },
      { q: '由最早沉積的地層到最晚沉積的地層，此處河水流速的變化依序為何？', o: ['變慢→變快→變慢', '變快→變慢→變快', '持續變快', '持續變慢'], a: 'A', e: '未經翻轉時下層較老。丁（沙）→丙（泥）：流速變慢；丙（泥）→乙（礫石）：變快；乙（礫石）→甲（沙）：變慢。' }] },
  { intro: '小華將甲、乙、丙、丁四種礦物兩兩互相刮磨，結果如下表：', table: [['刮磨的組合', '甲與乙', '乙與丙', '甲與丙', '甲與丁', '乙與丁'], ['受損的礦物', '乙', '乙', '甲', '丁', '乙']], widths: [1800, 1400, 1400, 1400, 1400, 1400],
    qs: [{ q: '四種礦物的硬度由大到小排列，應為下列何者？', o: ['甲＞丙＞丁＞乙', '丙＞丁＞甲＞乙', '丙＞甲＞丁＞乙', '甲＞丙＞乙＞丁'], a: 'C', e: '互相刮磨時硬度小者受損：甲＞乙、丙＞乙、丙＞甲、甲＞丁、丁＞乙，故丙＞甲＞丁＞乙。' }] },
  { intro: '下表為甲～丁四種礦物的性質：', table: [['礦物', '硬度', '條痕顏色', '滴加稀鹽酸', '解理'], ['甲', '7', '白色', '無反應', '無'], ['乙', '3', '白色', '冒氣泡', '三組'], ['丙', '1', '白色', '無反應', '一組'], ['丁', '2.5', '白色', '無反應', '一組（片狀）']], widths: [1200, 1200, 1800, 2000, 2400],
    qs: [{ q: '根據表中資料，下列敘述何者正確？', o: ['甲可以在乙的表面刮出刻痕', '丙是四種礦物中最不容易被刮傷的', '丁受力後容易沿三組平行的解理面裂開', '甲的主要成分為碳酸鈣'], a: 'A', e: '甲硬度7大於乙的3，可刮傷乙；丙硬度最小，最容易被刮傷；丁只有一組解理；會與稀鹽酸反應冒泡的是含碳酸鈣的乙（方解石），甲應為石英。' }] },
  { intro: '下表為甲～丁四種岩石的觀察紀錄：', table: [['岩石', '觀察紀錄'], ['甲', '由肉眼可見的大顆粒結晶緊密鑲嵌而成，沒有層理'], ['乙', '顆粒極細，具有明顯層理，岩層中可見貝殼化石'], ['丙', '結晶顆粒細小，肉眼難以分辨，整體呈灰黑色'], ['丁', '由大小不一的礫石與沙粒膠結而成']], widths: [1200, 7400],
    qs: [{ q: '四種岩石中，屬於火成岩的有哪些？', o: ['甲、乙', '甲、丙', '乙、丁', '丙、丁'], a: 'B', e: '甲結晶大（深成岩，如花岡岩）、丙結晶細小灰黑（火山岩，如玄武岩）皆為火成岩；乙（頁岩）、丁（礫岩）為沉積岩。' }] },
];

const L = ['A', 'B', 'C', 'D'];
const cb = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const borders = { top: cb, bottom: cb, left: cb, right: cb };
const mkTable = (rows, widths) => new Table({ alignment: AlignmentType.CENTER,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
  rows: rows.map((r, ri) => new TableRow({ children: r.map((c, ci) => new TableCell({
    borders, width: { size: widths[ci], type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 80, right: 80 },
    shading: (ri === 0 || (rows.length === 2 && ci === 0)) ? { fill: 'E7E6E6', type: 'clear', color: 'auto' } : undefined,
    children: [new Paragraph({ keepNext: true, alignment: widths.length === 2 && ci === 1 && ri > 0 ? AlignmentType.LEFT : AlignmentType.CENTER, children: [t(c, { size: 22, bold: ri === 0 })] })] })) })) });

const HANG = { left: 1100, hanging: 1100 }, OPT = { left: 1100 };
const body = [];
body.push(p([t('○○國民中學　115學年度第一學期　九年級自然科　第5章　圖表題', { size: 28, bold: true })], { align: AlignmentType.CENTER }));
body.push(p([t('範圍：第5章　水與陸地（5-1～5-3）', { size: 22 })], { align: AlignmentType.CENTER }));
body.push(p([t('　　年　　班　　號　姓名：＿＿＿＿＿＿＿　　得分：＿＿＿＿', { size: 24 })], { align: AlignmentType.RIGHT, after: 160 }));
body.push(p([t('圖表題（每題10分，共100分）', { bold: true, size: 26 })], { after: 100 }));

let n = 0; const all = [];
G.forEach(g => {
  body.push(p([t(g.intro, { bold: true })], { after: 60, keepNext: true }));
  if (g.table) { body.push(mkTable(g.table, g.widths)); body.push(p('', { after: 60, keepNext: true })); }
  if (g.fig) body.push(img(...g.fig));
  g.qs.forEach(x => {
    n++; all.push(x);
    body.push(p([t(`（　）${n}.`), t(x.q)], { indent: HANG, keepNext: true }));
    const maxLen = Math.max(...x.o.map(s => s.length));
    if (maxLen <= 5)
      body.push(p(x.o.map((s, k) => t(`(${L[k]})${s}${k < 3 ? '\t' : ''}`)), { indent: OPT, tabs: [2800, 4500, 6200].map(pos => ({ type: TabStopType.LEFT, position: pos })), after: 160 }));
    else if (maxLen <= 14)
      [0, 2].forEach(k => body.push(p([t(`(${L[k]})${x.o[k]}\t(${L[k + 1]})${x.o[k + 1]}`)], { indent: OPT, tabs: [{ type: TabStopType.LEFT, position: 5200 }], after: k === 2 ? 160 : 20, keepNext: k === 0 })));
    else x.o.forEach((s, k) => body.push(p([t(`(${L[k]})${s}`)], { indent: OPT, after: k === 3 ? 160 : 20, keepNext: k < 3 })));
  });
});

body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(p([t('答案與解析（教師用）', { size: 28, bold: true })], { align: AlignmentType.CENTER, after: 120 }));
body.push(mkTable([['題號', ...all.map((_, i) => String(i + 1))], ['答案', ...all.map(x => x.a)]], [900, ...Array(10).fill(820)]));
body.push(p('', { after: 120 }));
all.forEach((x, i) => body.push(p([t(`${i + 1}.（${x.a}）`, { bold: true, size: 22 }), t(x.e, { size: 22 })], { indent: { left: 600, hanging: 600 }, after: 60 })));

const doc = new Document({ styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } }, children: body }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('ok', n); });
