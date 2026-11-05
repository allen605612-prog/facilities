const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, PageBreak, TabStopType, ImageRun } = require('docx');

const FONT = { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '標楷體' };
const t = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 24, bold: o.bold });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [t(runs)], spacing: { after: o.after ?? 40, line: 340 }, alignment: o.align, indent: o.indent, tabStops: o.tabs, keepNext: o.keepNext, keepLines: true });
const FIG = 'C:\\Users\\user\\allen\\_exam_fig\\';
const img = (f, w, h) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, keepNext: true, children: [new ImageRun({ type: 'png', data: fs.readFileSync(FIG + f), transformation: { width: w, height: h } })] });

const Q = [
  { intro: '下表為某集水區一年的水收支資料（單位：毫米），降水除了蒸發散與形成地表逕流外，其餘皆入滲成為地下水。試回答第1～2題：',
    table: [['項目', '降水量', '蒸發散量', '地表逕流量', '入滲量'], ['水量', '2500', '1000', '1000', '？']], widths: [1600, 1500, 1500, 1700, 1500],
    q: '該年入滲成為地下水的水量，約占降水量的多少？', o: ['10％', '20％', '40％', '60％'], a: 'B',
    e: '入滲量＝2500－1000－1000＝500毫米，500÷2500＝20％。' },
  { q: '承上題，若該集水區大規模開發，原本的森林改建為住宅區與柏油道路，則水收支最可能發生下列哪一種變化？',
    o: ['蒸發散量增加，入滲量增加', '地表逕流量減少，入滲量增加', '降水量減少，入滲量不變', '地表逕流量增加，入滲量減少'], a: 'D',
    e: '柏油路面與建築物使雨水不易滲入地下，大部分成為地表逕流，因此逕流增加、入滲（地下水補注）減少，也較容易淹水。' },
  { q: '小安在三個相同的漏斗中，分別裝入等體積的乾燥礫石、粗砂、黏土，再各倒入200毫升的水，記錄5分鐘內漏斗下方收集到的水量，結果如下表。根據實驗結果，下列推論何者最合理？',
    table: [['漏斗內物質', '礫石', '粗砂', '黏土'], ['收集到的水量（毫升）', '180', '150', '20']], widths: [2800, 1500, 1500, 1500],
    o: ['黏土顆粒間孔隙大，水最容易通過', '由黏土固結形成的頁岩，最適合作為含水層', '顆粒較大的沉積物透水性較佳，由其形成的岩層較可能成為含水層', '礫石留住的水最多，因此礫岩層中的地下水最豐富'], a: 'C',
    e: '水通過礫石、粗砂較快，表示顆粒大、孔隙連通好，透水性佳，其形成的礫岩、砂岩常為含水層；黏土透水性差，形成的頁岩常為不透水層。礫石留住的水只有20毫升，最少。' },
  { q: '某平原上由北往南依序有丁、甲、乙、丙四口水井，一座垃圾掩埋場位於甲、乙兩井之間。已知地下水會由地下水面較高處往較低處流動，各井測得的地下水面海拔高度如下表。若掩埋場的汙水滲入地下，則哪兩口井的井水最可能受到汙染？',
    table: [['水井', '丁（最北）', '甲', '乙', '丙（最南）'], ['地下水面海拔（公尺）', '35', '32', '28', '24']], widths: [2800, 1500, 1200, 1200, 1500],
    o: ['乙、丙', '甲、丁', '甲、乙', '丁、丙'], a: 'A',
    e: '地下水面由北向南逐漸降低，地下水由北往南流。掩埋場位於甲、乙之間，汙水隨地下水往南流，會先到乙井，再到丙井。' },
  { q: '某古老墓園位於多雨且酸雨嚴重的地區，園中有兩塊同年代設立的墓碑：甲碑以花岡岩製成，刻字仍然清晰；乙碑以大理岩製成，刻字已模糊、表面粗糙。造成兩者差異的最主要原因為下列何者？',
    o: ['大理岩結晶顆粒較大，容易因熱脹冷縮而崩解', '大理岩主要成分為碳酸鈣，容易與酸性雨水反應而發生化學風化', '花岡岩的硬度比大理岩小，較不易被破壞', '乙碑位置較低，長期受到河流侵蝕'], a: 'B',
    e: '大理岩主要成分為碳酸鈣，遇酸會起化學反應而被溶蝕，屬化學風化；花岡岩主要由石英、長石等抗風化能力強的礦物組成。' },
  { q: '下列為四種自然現象的描述：\n甲、高山岩石裂縫中的水反覆結冰、融化，使岩石裂開\n乙、強風挾帶沙粒不斷撞擊岩柱，使岩柱底部變細\n丙、海浪不斷拍打海崖，在崖壁上形成海蝕洞\n丁、岩石表面長出地衣，分泌酸性物質使岩石表面變得鬆軟\n上述現象中，屬於「侵蝕作用」的是哪些？',
    o: ['甲、乙', '甲、丁', '丙、丁', '乙、丙'], a: 'D',
    e: '侵蝕作用是風、流水、海浪等外營力將岩石碎屑自原處剝離帶走的作用，乙（風蝕）、丙（海蝕）屬之；甲為物理風化，丁為化學風化。' },
  { q: '小宇在一條河流的彎道處，由左岸往右岸每隔5公尺測量一次水深，結果如下表。根據此資料，下列推論何者最合理？',
    table: [['與左岸的距離（公尺）', '0', '5', '10', '15', '20（右岸）'], ['水深（公尺）', '0.3', '0.8', '1.5', '2.6', '3.2']], widths: [2800, 900, 900, 900, 900, 1500],
    o: ['左岸為彎道外側，受到的側向侵蝕較強', '右岸流速較慢，容易堆積形成沙洲', '左岸水最深，最適合設置戲水區', '右岸為彎道外側，側向侵蝕較強，不宜在岸邊興建房屋'], a: 'D',
    e: '彎道外側流速快、側向侵蝕強，河道較深；內側流速慢，以沉積為主，河道較淺。右岸水深最大，應為外側，岸邊易被掏空。' },
  { intro: '下圖為某河川出海口附近海岸線位置的長期變化（以1960年的位置為0，向海為正），該河川上游的大型水庫於1978年完工。試回答第8～9題：',
    fig: ['h5_coast.png', 500, 265],
    q: '根據上圖，下列敘述何者正確？',
    o: ['1960～1978年間，河川帶到出海口的泥沙量少於海浪帶走的泥沙量', '1978年後海岸線持續向海擴張，陸地面積增加', '1978年後海岸線向陸地退縮，可能是水庫攔截泥沙，使河川輸沙量減少', '海岸線的位置只受海浪影響，與上游水庫無關'], a: 'C',
    e: '1978年前海岸線向海推進，表示泥沙供應量大於海浪帶走量；水庫完工後泥沙被攔截，供應量小於帶走量，海岸線向陸地退縮。' },
  { q: '承上題，下列哪一項措施最有助於減緩1978年以後的海岸線變化？',
    o: ['水庫定期排砂或設置繞庫排砂道，讓泥沙回到下游河道', '在河口大量抽取砂石作為建材', '在上游支流再加建數座攔沙壩', '大量抽取沿海地區的地下水'], a: 'A',
    e: '讓泥沙回到下游可增加出海口的泥沙供應；抽砂與攔沙壩會減少泥沙，超抽地下水則造成地層下陷，使問題更嚴重。' },
  { q: '在一條可調整坡度的水槽中，鋪上大小不同的顆粒，並保持每次流入的水量相同，觀察水流能帶走的最大顆粒，結果如下表。有關此實驗的敘述，下列何者正確？',
    table: [['水槽坡度', '5°', '10°', '20°', '30°'], ['能帶走的最大顆粒', '細砂', '粗砂', '小礫石', '大礫石']], widths: [2400, 1400, 1400, 1400, 1400],
    o: ['操縱變因為水量，水量越大，搬運能力越強', '應變變因為水槽坡度', '操縱變因為顆粒大小，顆粒越大越容易被搬運', '操縱變因為坡度，坡度越大流速越快，能搬運的顆粒越大'], a: 'D',
    e: '實驗中改變的是坡度（操縱變因），觀察的是能帶走的最大顆粒（應變變因），水量為控制變因。坡度大→流速快→搬運能力強，這也是河川上游能搬運大石塊的原因。' },
  { q: '下列為某校自然科教室的岩石檢索表：\n1. 滴加稀鹽酸會冒泡→到2；不會冒泡→到3\n2. 可見貝殼碎片等化石→甲；由緊密的方解石晶體組成，無化石→乙\n3. 可見大小不一的圓滑碎屑，以膠結物黏合→丙；由肉眼可見的礦物晶體緊密鑲嵌→丁\n甲、乙、丙、丁依序最可能為下列何者？',
    o: ['石灰岩、大理岩、礫岩、花岡岩', '大理岩、石灰岩、礫岩、玄武岩', '石灰岩、大理岩、花岡岩、礫岩', '大理岩、石灰岩、花岡岩、礫岩'], a: 'A',
    e: '含碳酸鈣者遇稀鹽酸冒泡：有化石的是石灰岩（沉積岩），重新結晶、無化石的是大理岩（變質岩）。不冒泡者中，由圓滑碎屑膠結的是礫岩；晶體肉眼可見且緊密鑲嵌的是花岡岩（玄武岩晶粒細小，肉眼不易見）。' },
  { q: '小芸暑假旅行時記錄了三處的岩石：\n甲地（澎湖）：黑色岩石，晶粒細小，呈柱狀排列\n乙地（金門）：淺色岩石，晶粒粗大，被用來建造碉堡\n丙地（花蓮太魯閣）：白色岩壁，滴加稀鹽酸會冒泡\n下列敘述何者正確？',
    o: ['甲地岩石為花岡岩，由岩漿在地表附近快速冷卻形成', '丙地岩石為火山噴發形成的安山岩', '乙地岩石晶粒粗大，是岩漿在地底深處緩慢冷卻形成的', '甲、乙、丙三地的岩石都屬於火成岩'], a: 'C',
    e: '甲為玄武岩（地表附近快速冷卻，晶粒細小）；乙為花岡岩（地底深處緩慢冷卻，晶粒粗大）；丙為大理岩，屬變質岩，主要成分碳酸鈣，遇稀鹽酸冒泡。' },
  { q: '登山隊在臺灣海拔三千多公尺的高山岩層中，發現貝類等海洋生物的化石。下列何者是最合理的解釋？',
    o: ['古代曾發生巨大海嘯，將貝殼沖到高山上', '這些岩層原本在海底沉積形成，之後受板塊擠壓等內營力作用而被抬升', '岩漿噴發時將海底生物帶到山頂，冷卻後形成化石', '鳥類將貝類帶到山上食用，遺骸長期累積形成化石'], a: 'B',
    e: '化石多保存在沉積岩中，海洋生物化石出現在高山，表示該岩層原在海底形成，後因板塊擠壓而抬升，是臺灣造山運動的證據。' },
  { q: '取等體積的乾燥岩石甲、乙、丙，浸泡在水中24小時後取出，擦乾表面再測量質量，結果如下表（吸水量越多，表示岩石孔隙越多）。若這三種岩石分別構成地底的三個岩層，則哪一個岩層最可能成為良好的含水層？',
    table: [['岩石', '甲', '乙', '丙'], ['浸水前質量（公克）', '250', '270', '280'], ['浸水後質量（公克）', '255', '270.5', '298']], widths: [2800, 1500, 1500, 1500],
    o: ['甲岩層，因為浸水前的質量最小', '乙岩層，因為吸水後質量幾乎不變', '丙岩層，因為吸水量最多，孔隙最多', '三個岩層都不能儲存地下水'], a: 'C',
    e: '吸水量：甲5公克、乙0.5公克、丙18公克。丙的孔隙最多，能儲存較多地下水，最可能成為含水層；乙質地緻密，接近不透水層。' },
  { q: '某山區將大片原始森林砍伐，改種植淺根性的高冷蔬菜，並開闢多條產業道路。數年後，當地及下游最可能出現下列哪些現象？\n甲、下雨時地表逕流增加\n乙、土壤侵蝕加劇，坡地易發生崩塌\n丙、下游河床因泥沙淤積而抬高\n丁、地下水的補注量增加',
    o: ['甲、乙、丙', '甲、乙、丁', '乙、丙、丁', '甲、丙、丁'], a: 'A',
    e: '森林根系可固定土壤、涵養水源。砍伐後雨水不易入滲，地表逕流增加、地下水補注減少（丁錯誤）；表土被大量沖刷，泥沙被帶到下游淤積，使河床抬高。' },
];

const L = ['A', 'B', 'C', 'D'];
const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };
const mkTable = (rows, widths) => new Table({
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, alignment: AlignmentType.CENTER,
  rows: rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, ci) => new TableCell({
    borders, width: { size: widths[ci], type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 80, right: 80 },
    shading: ri === 0 ? { fill: 'E7E6E6', type: 'clear', color: 'auto' } : undefined,
    children: [new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, children: [t(c, { size: 22, bold: ri === 0 })] })] })) })) });

const HANG = { left: 1100, hanging: 1100 };
const OPT_IND = { left: 1100 };
const body = [];
body.push(p([t('地表環境與資源　會考類似題', { size: 30, bold: true })], { align: AlignmentType.CENTER }));
body.push(p([t('範圍：地球上的水、地下水、地表營力與地貌、岩石', { size: 22 })], { align: AlignmentType.CENTER }));
body.push(p([t('　　年　　班　　號　姓名：＿＿＿＿＿＿＿　　得分：＿＿＿＿', { size: 24 })], { align: AlignmentType.RIGHT, after: 160 }));
body.push(p([t('單選題（共15題）', { bold: true, size: 26 })], { after: 100 }));

Q.forEach((x, i) => {
  if (x.intro) body.push(p([t(x.intro)], { after: 60, keepNext: true }));
  if (x.fig) body.push(img(...x.fig));
  const lines = x.q.split('\n');
  body.push(p([t(`（　）${i + 1}.`), t(lines[0])], { indent: HANG, keepNext: true }));
  lines.slice(1).forEach(s => body.push(p([t(s)], { indent: OPT_IND, keepNext: true })));
  if (x.table) { body.push(mkTable(x.table, x.widths)); body.push(p('', { after: 20, keepNext: true })); }
  const short = x.o.every(s => s.length <= 8);
  if (short) {
    body.push(p(x.o.map((s, k) => t(`(${L[k]})${s}${k < 3 ? '\t' : ''}`)), { indent: OPT_IND, tabs: [2800, 4500, 6200].map(pos => ({ type: TabStopType.LEFT, position: pos })), after: 140 }));
  } else {
    x.o.forEach((s, k) => body.push(p([t(`(${L[k]})${s}`)], { indent: OPT_IND, after: k === 3 ? 140 : 20, keepNext: k < 3 })));
  }
});

body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(p([t('答案與解析（教師用）', { size: 28, bold: true })], { align: AlignmentType.CENTER, after: 120 }));
const nums = Q.map((_, i) => String(i + 1)), ans = Q.map(x => x.a);
body.push(mkTable([['題號', ...nums.slice(0, 8)], ['答案', ...ans.slice(0, 8)], ['題號', ...nums.slice(8), ''], ['答案', ...ans.slice(8), '']], [900, ...Array(8).fill(900)]));
body.push(p('', { after: 120 }));
Q.forEach((x, i) => body.push(p([t(`${i + 1}.（${x.a}）`, { bold: true, size: 22 }), t(x.e, { size: 22 })], { indent: { left: 600, hanging: 600 }, after: 60 })));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } }, children: body }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('ok', Q.length, ans.join('')); });
