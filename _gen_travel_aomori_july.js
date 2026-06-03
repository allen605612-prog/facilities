const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak
} = require("docx");
const fs = require("fs");

// 顏色定義
const DEEP_BLUE  = "1A3557";
const MID_BLUE   = "2E6DA4";
const LIGHT_BLUE = "D6E8F7";
const PALE_BLUE  = "EEF5FB";
const GOLD       = "D4A843";
const WHITE_COL  = "FFFFFF";
const DARK_GRAY  = "333333";
const MID_GRAY   = "888888";
const RED_ACCENT = "C0392B";

const CONTENT_W = 9746;

// ---- 輔助函數 ----

function hr() {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: LIGHT_BLUE, space: 1 } },
    children: []
  });
}

function spacer(pts = 6) {
  return new Paragraph({ spacing: { before: pts * 20, after: 0 }, children: [] });
}

function coverTitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, font: "Microsoft JhengHei", size: 64, bold: true, color: DEEP_BLUE })]
  });
}

function coverSubtitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: "Microsoft JhengHei", size: 32, color: MID_BLUE })]
  });
}

function coverMeta(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 40 },
    children: [new TextRun({ text, font: "Microsoft JhengHei", size: 24, color: DARK_GRAY })]
  });
}

function coverRoute(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 1 },
      top:    { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 1 }
    },
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: "Microsoft JhengHei", size: 22, color: MID_BLUE })]
  });
}

function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 300, after: 100 },
    shading: { fill: DEEP_BLUE, type: ShadingType.CLEAR },
    indent: { left: -200, right: -200 },
    children: [new TextRun({ text: "  " + text, font: "Microsoft JhengHei", size: 26, bold: true, color: WHITE_COL })]
  });
}

function dayHeading(text) {
  return new Paragraph({
    spacing: { before: 240, after: 60 },
    shading: { fill: MID_BLUE, type: ShadingType.CLEAR },
    children: [new TextRun({ text: "  " + text, font: "Microsoft JhengHei", size: 24, bold: true, color: WHITE_COL })]
  });
}

function bullet(text, isStar = false) {
  const color = isStar ? RED_ACCENT : DARK_GRAY;
  return new Paragraph({
    spacing: { before: 30, after: 30 },
    indent: { left: 360, hanging: 360 },
    children: [new TextRun({ text: (isStar ? "★ " : "・") + text, font: "Microsoft JhengHei", size: 20, color })]
  });
}

function stayLine(text) {
  return new Paragraph({
    spacing: { before: 40, after: 60 },
    indent: { left: 360 },
    children: [
      new TextRun({ text: "▶ 住宿推薦：", font: "Microsoft JhengHei", size: 20, bold: true, color: MID_BLUE }),
      new TextRun({ text, font: "Microsoft JhengHei", size: 20, color: MID_BLUE })
    ]
  });
}

function noteItem(text) {
  return new Paragraph({
    spacing: { before: 30, after: 30 },
    indent: { left: 200 },
    children: [new TextRun({ text, font: "Microsoft JhengHei", size: 19, color: MID_GRAY })]
  });
}

function normalText(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: "Microsoft JhengHei", size: opts.size || 20, color: opts.color || DARK_GRAY, bold: opts.bold })]
  });
}

const borderDef = { style: BorderStyle.SINGLE, size: 4, color: "AACCEE" };
const borders   = { top: borderDef, bottom: borderDef, left: borderDef, right: borderDef, insideHorizontal: borderDef, insideVertical: borderDef };
const cellM     = { top: 80, bottom: 80, left: 120, right: 120 };

function makeTable(headers, rows, colWidths) {
  const headerRow = new TableRow({
    children: headers.map((h, i) =>
      new TableCell({
        borders, width: { size: colWidths[i], type: WidthType.DXA }, margins: cellM,
        shading: { fill: MID_BLUE, type: ShadingType.CLEAR },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: h, font: "Microsoft JhengHei", size: 20, bold: true, color: WHITE_COL })]
        })]
      })
    )
  });
  const dataRows = rows.map((row, ri) =>
    new TableRow({
      children: row.map((cell, ci) =>
        new TableCell({
          borders, width: { size: colWidths[ci], type: WidthType.DXA }, margins: cellM,
          shading: { fill: ri % 2 === 0 ? PALE_BLUE : WHITE_COL, type: ShadingType.CLEAR },
          children: [new Paragraph({
            children: [new TextRun({ text: String(cell), font: "Microsoft JhengHei", size: 19, color: DARK_GRAY })]
          })]
        })
      )
    })
  );
  return new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: colWidths, rows: [headerRow, ...dataRows] });
}

function makeCostTable(rows, totalLabel, totalAmount) {
  const cw = [Math.round(CONTENT_W * 0.62), Math.round(CONTENT_W * 0.38)];
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: cw,
    rows: [
      new TableRow({
        children: ["項目", "金額（新台幣）"].map((h, i) =>
          new TableCell({
            borders, width: { size: cw[i], type: WidthType.DXA }, margins: cellM,
            shading: { fill: MID_BLUE, type: ShadingType.CLEAR },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, font: "Microsoft JhengHei", size: 20, bold: true, color: WHITE_COL })] })]
          })
        )
      }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((cell, ci) => new TableCell({
          borders, width: { size: cw[ci], type: WidthType.DXA }, margins: cellM,
          shading: { fill: ri % 2 === 0 ? PALE_BLUE : WHITE_COL, type: ShadingType.CLEAR },
          children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Microsoft JhengHei", size: 19, color: DARK_GRAY })] })]
        }))
      })),
      new TableRow({
        children: [totalLabel, totalAmount].map((cell, ci) => new TableCell({
          borders, width: { size: cw[ci], type: WidthType.DXA }, margins: cellM,
          shading: { fill: DEEP_BLUE, type: ShadingType.CLEAR },
          children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Microsoft JhengHei", size: 21, bold: true, color: GOLD })] })]
        }))
      })
    ]
  });
}

function makeHeader(title, date) {
  return new Header({
    children: [new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: DEEP_BLUE, space: 1 } },
      spacing: { after: 80 },
      children: [
        new TextRun({ text: title, font: "Microsoft JhengHei", size: 16, color: MID_BLUE }),
        new TextRun({ text: "\t" + date, font: "Microsoft JhengHei", size: 16, color: MID_GRAY }),
      ],
      tabStops: [{ type: "right", position: 8600 }],
    })]
  });
}

function makeFooter() {
  return new Footer({
    children: [new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: LIGHT_BLUE, space: 1 } },
      alignment: AlignmentType.CENTER,
      spacing: { before: 60 },
      children: [
        new TextRun({ text: "第 ", font: "Microsoft JhengHei", size: 16, color: MID_GRAY }),
        new TextRun({ children: [PageNumber.CURRENT], font: "Microsoft JhengHei", size: 16, color: MID_GRAY }),
        new TextRun({ text: " 頁", font: "Microsoft JhengHei", size: 16, color: MID_GRAY }),
      ]
    })]
  });
}

function buildDoc(cfg) {
  const c = [];

  // === 封面 ===
  c.push(spacer(80));
  c.push(coverTitle(cfg.title));
  c.push(coverSubtitle("七天親子旅行完全攻略"));
  c.push(spacer(20));
  c.push(coverRoute(cfg.route));
  c.push(spacer(40));
  c.push(coverMeta("成員：父親 ＋ 兩位女兒（17歲以上，共3人）"));
  c.push(coverMeta("出發日期：2026年7月1日（三）"));
  c.push(spacer(80));
  c.push(new Paragraph({ children: [new PageBreak()] }));

  // === 旅程概覽 ===
  c.push(sectionHeading("旅程概覽"));
  c.push(spacer(8));
  c.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: cfg.overview, font: "Microsoft JhengHei", size: 21, color: MID_BLUE, bold: true })]
  }));
  c.push(hr());

  c.push(normalText("七月初氣候參考", { bold: true, size: 22, color: DEEP_BLUE }));
  c.push(makeTable(
    ["地區", "氣溫", "天氣特點"],
    cfg.climate,
    [Math.round(CONTENT_W * 0.18), Math.round(CONTENT_W * 0.20), Math.round(CONTENT_W * 0.62)]
  ));
  c.push(spacer(12));
  c.push(new Paragraph({ children: [new PageBreak()] }));

  // === 逐日行程 ===
  c.push(sectionHeading("逐日行程"));
  c.push(spacer(8));

  for (const d of cfg.days) {
    c.push(dayHeading(d.title));
    for (const item of d.items) c.push(bullet(item.text, item.star));
    c.push(stayLine(d.stay));
  }

  c.push(new Paragraph({ children: [new PageBreak()] }));

  // === 交通規劃 ===
  c.push(sectionHeading("交通規劃"));
  c.push(spacer(8));
  c.push(makeTable(
    ["交通工具", "區間", "備註"],
    cfg.transport,
    [Math.round(CONTENT_W * 0.24), Math.round(CONTENT_W * 0.34), Math.round(CONTENT_W * 0.42)]
  ));
  c.push(spacer(8));
  for (const n of cfg.transportNotes) c.push(noteItem(n));
  c.push(spacer(16));

  // === 住宿一覽 ===
  c.push(sectionHeading("住宿一覽"));
  c.push(spacer(8));
  c.push(makeTable(
    ["地點", "推薦住宿", "特色", "估價/晚（3人）"],
    cfg.hotels,
    [Math.round(CONTENT_W * 0.20), Math.round(CONTENT_W * 0.28), Math.round(CONTENT_W * 0.26), Math.round(CONTENT_W * 0.26)]
  ));
  c.push(spacer(16));

  // === 費用概估 ===
  c.push(sectionHeading("費用概估（每人）"));
  c.push(spacer(8));
  c.push(makeCostTable(cfg.costs, "合計/人", cfg.totalCost));
  c.push(spacer(16));

  // === 注意事項 ===
  c.push(sectionHeading("重要注意事項"));
  c.push(spacer(8));
  for (const n of cfg.notes) c.push(bullet(n, false));
  c.push(spacer(12));

  // === 備選景點 ===
  c.push(sectionHeading("備選景點"));
  c.push(spacer(8));
  for (const e of cfg.extras) c.push(bullet(e, false));

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Microsoft JhengHei", size: 20 } } }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      headers: { default: makeHeader(cfg.headerTitle, "2026.07.01 出發") },
      footers: { default: makeFooter() },
      children: c
    }]
  });

  return doc;
}

// ================================================================
//  版本一：不含函館（弘前→白神山地→奧入瀨→青森市→山寺→銀山溫泉）
// ================================================================
const cfgNoHakodate = {
  title: "七月初東北・青森・山形",
  route: "弘前 → 白神山地 → 奧入瀨 → 青森市 → 山寺 → 銀山温泉 → 山形",
  overview: "台灣出發  →  弘前（2晚）→  青森市（2晚）→  山形＋銀山（2晚）→  東京轉機  →  台灣",
  headerTitle: "七月初東北七天（不含函館版）",
  climate: [
    ["函館（路過）", "—", "本版不前往函館"],
    ["青森", "18–25°C", "梅雨尾聲，奧入瀨溪流新綠盛況"],
    ["山形", "20–28°C", "仍有梅雨，山寺翠綠壯觀"],
    ["建議裝備", "—", "折疊傘、舒適運動鞋（必備）"],
  ],
  days: [
    {
      title: "Day 1（7/1）台灣出發 → 抵達弘前",
      items: [
        { text: "出發：台灣桃園機場飛東京羽田（HND，約3.5–4小時）", star: false },
        { text: "移動：✈ 國內線 東京(HND) → 青森(AOJ)（ANA/JAL外國人優惠票 ¥10,800，約1小時）", star: true },
        { text: "移動：機場巴士 青森機場 → 弘前（¥900，約55分鐘）", star: false },
        { text: "傍晚：弘前市內輕鬆散步，土手町商店街覓食", star: false },
      ],
      stay: "JR Inn 弘前（近弘前車站・新穎舒適）"
    },
    {
      title: "Day 2（7/2）弘前深度遊",
      items: [
        { text: "上午：弘前城（夏季護城河倒映翠綠，蓮花盛開）", star: true },
        { text: "上午：追手門廣場・石崁樹海散步道（歷史街區）", star: false },
        { text: "中午：弘前市蘋果公園（免費入園，7月青蘋果可摘，附設蘋果甜點）", star: false },
        { text: "下午：津輕藩ねぷた村（弘前ねぷた燈籠文化體驗，含手工體驗）", star: false },
        { text: "傍晚：弘前城夕照、城濠倒影", star: true },
      ],
      stay: "JR Inn 弘前"
    },
    {
      title: "Day 3（7/3）白神山地 ★ UNESCO 世界遺產",
      items: [
        { text: "早晨：租車出發（弘前 → 白神山地 十二湖，約1.5–2小時）", star: false },
        { text: "上午：青池（十二湖中最知名，鈷藍色神秘湖泊，非人工染色）", star: true },
        { text: "上午：十二湖湖岸步道散步（全步道約1–2小時，平緩好走）", star: false },
        { text: "下午：千畳敷海岸（日本海奇岩海蝕平台，廣大而壯觀）", star: false },
        { text: "傍晚：還車，移動至青森市入住", star: false },
      ],
      stay: "Richmond Hotel 青森（大型連鎖・位置優・CP值高）"
    },
    {
      title: "Day 4（7/4）奧入瀨溪流 & 十和田湖",
      items: [
        { text: "早晨：租車出發（青森 → 奧入瀨溪流 約1小時）", star: false },
        { text: "上午－下午：奧入瀨溪流（9km步道，雲井瀑布・玉簾瀑布・銚子大瀑布，新綠盛況）", star: true },
        { text: "下午：十和田湖（火山口湖，湖岸散步＋遊覽船可選）", star: false },
        { text: "下午：十和田湖畔「乙女像」（高村光太郎作品）", star: false },
        { text: "傍晚：返回青森市", star: false },
      ],
      stay: "Richmond Hotel 青森"
    },
    {
      title: "Day 5（7/5）青森市精華 → 山形・山寺",
      items: [
        { text: "上午：青森縣立美術館（奈良美智「青森犬」，門票¥510/人）", star: true },
        { text: "上午：青森魚菜中心のっけ丼（代幣自選，海鮮蓋飯體驗）", star: false },
        { text: "中午：高速巴士 青森 → 仙台（¥3,000–4,000，約3.5小時）★省錢替代新幹線", star: true },
        { text: "中午：仙山線 仙台 → 山寺（IC卡 ¥860，約1小時）", star: false },
        { text: "下午：山寺・立石寺（1,015段石階登頂，俯瞰山谷，芭蕉「閑さや」俳句聖地）", star: true },
        { text: "傍晚：仙山線移動至山形市入住", star: false },
      ],
      stay: "ドーミーイン山形（天然溫泉大浴場・近車站）"
    },
    {
      title: "Day 6（7/6）銀山温泉 ★ 本版最大亮點",
      items: [
        { text: "上午：JR奧羽本線或包車前往銀山温泉（山形→大石田站約30分，再轉巴士20分）", star: false },
        { text: "上午：銀山溫泉街散步（大正浪漫復古建築，石橋・銀山川瀑布・木造多層旅館）", star: true },
        { text: "下午：白銀瀑布步道（溫泉街最深處，清涼瀑布）", star: false },
        { text: "下午：入住溫泉旅館，泡溫泉、享用日式懷石料理晚餐", star: false },
        { text: "夜間：煤氣燈點亮的溫泉街夜景（如宮崎駿「神隱少女」場景，超適合拍照）", star: true },
      ],
      stay: "能登屋旅館（最具代表性大正建築・強烈建議提早2–3個月訂）"
    },
    {
      title: "Day 7（7/7）銀山温泉 → 回台灣",
      items: [
        { text: "清晨：溫泉街晨光散步（遊客未至，最靜謐最美）", star: true },
        { text: "早餐：旅館早餐（日式懷石早膳）", star: false },
        { text: "清晨：銀山溫泉晨光散步後退房，JR巴士＋奧羽本線返回山形站（¥1,130，約1小時）", star: false },
        { text: "上午：仙山線 山形 → 仙台（¥860，約1小時）→ 仙台機場 Access 線（¥660，約17分）", star: false },
        { text: "中午：✈ 國內線 仙台(SDJ) → 東京(HND)（ANA/JAL外國人優惠票，¥10,800，約1小時）", star: true },
        { text: "下午：東京成田/羽田轉機回台灣", star: false },
      ],
      stay: "—（返程）"
    },
  ],
  transport: [
    ["✈ 國內線（ANA/JAL外國人優惠票）", "東京(HND) → 青森(AOJ)", "¥10,800/人，約1小時（進程）"],
    ["機場巴士", "青森機場 → 弘前", "¥900/人，約55分鐘"],
    ["高速巴士", "青森 → 仙台", "¥3,000–4,000/人，約3.5小時"],
    ["仙山線", "仙台 → 山寺 → 山形", "IC卡 ¥1,720，各約1小時"],
    ["JR奧羽本線＋巴士", "山形 → 大石田 → 銀山温泉", "¥630 + ¥500，約1小時"],
    ["✈ 國內線（ANA/JAL外國人優惠票）", "仙台(SDJ) → 東京(HND)", "¥10,800/人，約1小時（返程）"],
    ["仙台機場 Access 線", "仙台站 → 仙台機場", "IC卡 ¥660，約17分鐘"],
    ["租車", "Day 3–4（白神山地・奧入瀨）", "需備台灣國際駕照＋駕照正本"],
  ],
  transportNotes: [
    "兩段國內線合計¥21,600/人，全程不需 JR Pass：進程飛青森省¥6,670，回程飛東京省¥7,540 vs 新幹線",
    "ANA Experience Japan Fare / JAL Explorer Pass：¥10,800固定/段，需在日本境外購票，憑外國護照搭乘",
    "仙台機場有直飛東京羽田（HND）航班，班次多，建議選早午班次以利當天轉機回台",
    "租車：弘前市租車行，日系車型，保險含碰撞。Day 3–4約¥8,000–12,000/天",
  ],
  hotels: [
    ["弘前（2晚）", "JR Inn 弘前", "近車站・新穎舒適", "¥15,000–20,000"],
    ["青森市（2晚）", "Richmond Hotel 青森", "大型連鎖・CP值高", "¥18,000–25,000"],
    ["山形市（1晚）", "ドーミーイン山形", "天然溫泉大浴場", "¥20,000–28,000"],
    ["銀山温泉（1晚）", "能登屋旅館", "大正浪漫・必體驗", "¥45,000–70,000"],
  ],
  costs: [
    ["機票（台灣往返，轉東京羽田）", "NT$ 18,000–28,000"],
    ["✈ 國內線 東京→青森（進程 ¥10,800）", "NT$ 2,500"],
    ["✈ 國內線 仙台→東京（返程 ¥10,800）", "NT$ 2,500"],
    ["高速巴士 青森→仙台 ¥3,500", "NT$ 810"],
    ["仙台機場 Access 線＋IC卡 ¥4,000", "NT$ 925"],
    ["住宿（6晚，均分3人）", "NT$ 14,000–22,000"],
    ["銀山温泉旅館（能登屋・單人份）", "NT$ 5,000–8,000"],
    ["餐飲（每日約NT$1,200）", "NT$ 8,400"],
    ["租車（2天，均分3人）", "NT$ 2,500"],
    ["景點門票・雜費", "NT$ 2,000"],
  ],
  totalCost: "約 NT$ 57,000–76,000（2段國內線，全程不需 JR Pass）",
  notes: [
    "1. 銀山温泉旅館（尤其能登屋）建議出發前2–3個月以上訂房，旺季極難搶",
    "2. 七月初青森、山形仍在梅雨季，請備折疊傘，奧入瀨當日如遇雨反有雨中美景",
    "3. 山寺（立石寺）1,015段石階，請穿舒適運動鞋，建議預留2–3小時",
    "4. 奧入瀨溪流步道平緩（落差少），適合親子，建議早上9時前出發避開人潮",
    "5. 白神山地道路部分為山路，租車前確認女兒中是否有人能輔助駕駛",
    "6. 台灣國際駕照申辦：出發前至各縣市監理站，費用約NT$250，當日可取",
    "7. 青森縣立美術館「青森犬」必拍，建議預留1.5–2小時",
  ],
  extras: [
    "蔵王御釜（山形，火山口湖，7–10月開放，自駕約1小時，霧景壯觀）",
    "弘前市植物園・植物園溫室（弘前公園內，夏季荷花盛開）",
    "山形文翔館（大正建築，免費入場，適合拍照）",
    "恐山（靈場，下北半島，需額外半天，不在主線但值得一提）",
    "青森棟方志功紀念館（版畫巨匠・國際知名，門票¥600）",
  ]
};

// ================================================================
//  版本二：含函館（函館→弘前→白神山地→奧入瀨→青森市→山寺→山形）
// ================================================================
const cfgWithHakodate = {
  title: "七月初東北・函館・青森・山形",
  route: "函館 → 弘前 → 白神山地 → 奧入瀨 → 青森市 → 山寺 → 山形 → 東京",
  overview: "台灣出發  →  函館（2晚）→  弘前（1晚）→  青森市（2晚）→  山形（1晚）→  東京轉機  →  台灣",
  headerTitle: "七月初東北七天（含函館版）",
  climate: [
    ["函館", "16–22°C", "北海道夏季涼爽，無梅雨，早晚需薄外套"],
    ["青森", "18–25°C", "梅雨尾聲，奧入瀨溪流新綠盛況"],
    ["山形", "20–28°C", "仍有梅雨，山寺翠綠壯觀"],
    ["建議裝備", "—", "折疊傘、薄外套（函館用）、舒適運動鞋"],
  ],
  days: [
    {
      title: "Day 1（7/1）台灣出發 → 抵達函館",
      items: [
        { text: "出發：台灣桃園機場飛東京羽田（HND，約3.5–4小時）", star: false },
        { text: "移動：✈ 國內線 東京(HND) → 函館(HKD)（ANA/JAL外國人優惠票 ¥10,800，約1.5小時，省¥11,890！）", star: true },
        { text: "移動：函館空港巴士 → 函館市區（¥450，約20分鐘）", star: false },
        { text: "傍晚：函館元町區域漫步，坂道風景、教會群初覽", star: false },
      ],
      stay: "La Vista Hakodate Bay（函館灣景、早餐著名、強推）"
    },
    {
      title: "Day 2（7/2）函館深度遊 ★ 夜景日本三大",
      items: [
        { text: "早晨：函館朝市（6–10時，海膽・毛蟹・烏賊，新鮮直送）", star: true },
        { text: "上午：五稜郭公園（明治時代西式星形城廓，塔上俯瞰最佳）", star: false },
        { text: "下午：元町洋館群（天主教教會・舊函館公會堂・坡道街景）", star: false },
        { text: "傍晚：湯之川温泉或立待岬眺望（任選）", star: false },
        { text: "夜間：函館山夜景（日本三大夜景之一，纜車或步行登頂）", star: true },
      ],
      stay: "La Vista Hakodate Bay"
    },
    {
      title: "Day 3（7/3）函館 → 弘前（新幹線移動日）",
      items: [
        { text: "早晨：函館朝市二訪或旅館知名早餐（La Vista早餐口碑極佳）", star: false },
        { text: "上午：東北・北海道新幹線 新函館北斗 → 新青森（はやぶさ，約1小時）", star: false },
        { text: "中午：奧羽本線 新青森 → 弘前（約40分）", star: false },
        { text: "下午：弘前城夕照・追手門廣場散策", star: false },
        { text: "下午：津輕藩ねぷた村（弘前ねぷた燈籠文化體驗）", star: false },
      ],
      stay: "JR Inn 弘前（近弘前車站・新穎舒適）"
    },
    {
      title: "Day 4（7/4）白神山地 ★ UNESCO 世界遺產",
      items: [
        { text: "早晨：租車出發（弘前 → 白神山地 十二湖，約1.5–2小時）", star: false },
        { text: "上午：青池（十二湖中最知名，鈷藍色神秘湖泊）", star: true },
        { text: "上午：十二湖湖岸步道散步（平緩好走，約1–2小時）", star: false },
        { text: "下午：千畳敷海岸（日本海奇岩海蝕平台）", star: false },
        { text: "傍晚：還車，移動至青森市入住", star: false },
      ],
      stay: "Richmond Hotel 青森（大型連鎖・位置優・CP值高）"
    },
    {
      title: "Day 5（7/5）奧入瀨溪流 & 十和田湖",
      items: [
        { text: "早晨：租車出發（青森 → 奧入瀨溪流 約1小時）", star: false },
        { text: "上午－下午：奧入瀨溪流（9km步道，雲井瀑布・玉簾瀑布・銚子大瀑布）", star: true },
        { text: "下午：十和田湖（火山口湖，湖岸散步＋遊覽船可選）", star: false },
        { text: "傍晚：WA-RASSE 睡魔祭博物館（全年開放，門票¥600/人）", star: false },
      ],
      stay: "Richmond Hotel 青森"
    },
    {
      title: "Day 6（7/6）青森市 → 山形・山寺",
      items: [
        { text: "上午：青森縣立美術館（奈良美智「青森犬」，門票¥510/人）", star: true },
        { text: "上午：青森魚菜中心のっけ丼（代幣自選，海鮮蓋飯體驗）", star: false },
        { text: "中午：東北新幹線 新青森 → 仙台（約1小時）", star: false },
        { text: "中午：仙山線 仙台 → 山寺（約1小時）", star: false },
        { text: "下午：山寺・立石寺（1,015段石階登頂，俯瞰山谷，芭蕉俳句聖地）", star: true },
        { text: "傍晚：仙山線移動至山形市入住", star: false },
      ],
      stay: "ドーミーイン山形（天然溫泉大浴場・近車站）"
    },
    {
      title: "Day 7（7/7）山形巡遊 → 回台灣",
      items: [
        { text: "早晨：霞城公園（山形城遺跡・荷花池・歷史氛圍）", star: false },
        { text: "上午：山形文翔館（大正7年建造英式石造建築，免費參觀）", star: false },
        { text: "上午：山形縣鄉土館（山形歷史展示）", star: false },
        { text: "中午：仙山線 山形 → 仙台（¥860，約1小時）→ 仙台機場 Access 線（¥660，約17分）", star: false },
        { text: "下午：✈ 國內線 仙台(SDJ) → 東京(HND)（ANA/JAL外國人優惠票，¥10,800，約1小時）", star: true },
        { text: "傍晚：東京成田/羽田轉機回台灣", star: false },
      ],
      stay: "—（返程）"
    },
  ],
  transport: [
    ["✈ 國內線（ANA/JAL外國人優惠票）", "東京(HND) → 函館(HKD)", "¥10,800/人，約1.5小時（進程）"],
    ["函館空港巴士", "函館機場 → 函館市區", "¥450/人，約20分鐘"],
    ["東北・北海道新幹線（はやぶさ）", "新函館北斗 → 新青森", "點對點票 ¥9,500，約1小時"],
    ["奧羽本線", "新青森 → 弘前", "IC卡 ¥680，約40分"],
    ["高速巴士", "青森 → 仙台", "¥3,000–4,000/人，約3.5小時"],
    ["仙山線", "仙台 → 山寺 → 山形", "IC卡 ¥1,720，各約1小時"],
    ["✈ 國內線（ANA/JAL外國人優惠票）", "仙台(SDJ) → 東京(HND)", "¥10,800/人，約1小時（返程）"],
    ["仙台機場 Access 線", "仙台站 → 仙台機場", "IC卡 ¥660，約17分鐘"],
    ["租車", "Day 4–5（白神・奧入瀨）", "需備台灣國際駕照＋駕照正本"],
  ],
  transportNotes: [
    "兩段國內線合計¥21,600/人，全程不需 JR Pass：進程飛函館省¥11,890，回程飛東京省¥7,540 vs 新幹線",
    "ANA Experience Japan Fare / JAL Explorer Pass：¥10,800固定/段，需在日本境外購票，憑外國護照搭乘",
    "新函館北斗→新青森 仍需新幹線，點對點購票 ¥9,500，不需買整個 JR Pass",
    "仙台機場有直飛東京羽田（HND）航班，班次多，建議選早午班次以利當天轉機回台",
    "租車：青森市租車行，Day 4–5約¥8,000–12,000/天",
  ],
  hotels: [
    ["函館（2晚）", "La Vista Hakodate Bay", "海灣景・早餐著名", "¥30,000–45,000"],
    ["弘前（1晚）", "JR Inn 弘前", "近車站・新穎舒適", "¥15,000–20,000"],
    ["青森市（2晚）", "Richmond Hotel 青森", "大型連鎖・CP值高", "¥18,000–25,000"],
    ["山形市（1晚）", "ドーミーイン山形", "天然溫泉大浴場", "¥20,000–28,000"],
  ],
  costs: [
    ["機票（台灣往返，轉東京羽田）", "NT$ 18,000–28,000"],
    ["✈ 國內線 東京→函館（進程 ¥10,800）", "NT$ 2,500"],
    ["✈ 國內線 仙台→東京（返程 ¥10,800）", "NT$ 2,500"],
    ["新幹線 新函館北斗→新青森（¥9,500）", "NT$ 2,200"],
    ["高速巴士 青森→仙台 ¥3,500", "NT$ 810"],
    ["仙台機場 Access 線＋IC卡 ¥4,000", "NT$ 925"],
    ["住宿（6晚，均分3人）", "NT$ 16,000–24,000"],
    ["餐飲（每日約NT$1,500，含函館海鮮）", "NT$ 10,500"],
    ["租車（2天，均分3人）", "NT$ 2,500"],
    ["景點門票・雜費", "NT$ 2,000"],
  ],
  totalCost: "約 NT$ 58,000–75,000（2段國內線，全程不需 JR Pass）",
  notes: [
    "1. La Vista Hakodate Bay 早餐超有名，建議提前訂含早餐房型",
    "2. 函館山夜景：旺季（7月）纜車等待時間長，建議17:30前上山佔位",
    "3. 函館朝市：最新鮮的時段是6–9時，過10時後部分攤位收攤",
    "4. 七月初青森、山形仍有梅雨，請備折疊傘（函館則無梅雨，涼爽舒適）",
    "5. 山寺（立石寺）1,015段石階，建議穿舒適運動鞋，預留2–3小時",
    "6. 台灣國際駕照申辦：出發前至各縣市監理站，費用約NT$250，當日可取",
    "7. 道南いさりび鐵道（新函館北斗→函館）約¥360/人，不需另購票",
  ],
  extras: [
    "銀山温泉（山形，大正浪漫溫泉街，可改Day 7白天體驗，住宿不易）",
    "蔵王御釜（山形，火山口湖，7–10月開放，自駕約1小時）",
    "青森棟方志功紀念館（版畫巨匠，門票¥600）",
    "立待岬（函館，海崖眺望津輕海峽，日落最美）",
    "恐山（靈場，下北半島，如多一天可加入）",
  ]
};

// ---- 生成兩份文件 ----
async function main() {
  const pairs = [
    { cfg: cfgNoHakodate,   out: "C:\\Users\\user\\allen\\travel_七月青森不含函館版.docx" },
    { cfg: cfgWithHakodate, out: "C:\\Users\\user\\allen\\travel_七月青森含函館版.docx"   },
  ];
  for (const { cfg, out } of pairs) {
    const doc = buildDoc(cfg);
    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(out, buf);
    console.log("✅ 已產生：" + out);
  }
}

main().catch(e => { console.error(e); process.exit(1); });
