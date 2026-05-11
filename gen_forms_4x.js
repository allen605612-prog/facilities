// 根據 設備組內控修訂.docx 產生
// 4.1 主機房工作紀錄表
// 4.2 檔案備份計畫
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, LevelFormat, TabStopType, TabStopPosition,
  PageBreak
} = require('C:/Users/user/AppData/Roaming/npm/node_modules/docx');
const fs = require('fs');

const FONT  = '標楷體';
const C_HDR = '1F3864';   // 深藍 header
const C_SUB = '2E5F8A';   // 中藍 section
const C_FIL = 'D6E4F0';   // 淡藍 label 底色
const C_BD  = '2E5F8A';   // 邊框色
const W     = 9746;        // A4 content width (11906 - 2×1080)

// ── 基礎元件 ─────────────────────────────────────────────────
const run  = (t, o={}) => new TextRun({ text:t, font:{name:FONT}, size:20, ...o });
const brun = (t, o={}) => run(t, { bold:true, ...o });
const hrun = (t, sz, col, bold=true) => new TextRun({ text:t, font:{name:FONT}, size:sz, color:col, bold });
const par  = (ch, o={}) => new Paragraph({ children:Array.isArray(ch)?ch:[ch], ...o });
const sp   = (b=80) => par([], { spacing:{before:b,after:0} });
const PB   = () => par([new PageBreak()]);

const bd  = (c=C_BD) => ({ style:BorderStyle.SINGLE, size:6, color:c });
const aB  = (c=C_BD) => ({ top:bd(c), bottom:bd(c), left:bd(c), right:bd(c) });
const nB  = ()       => ({ top:bd('FFFFFF'), bottom:bd('FFFFFF'), left:bd('FFFFFF'), right:bd('FFFFFF') });

// WeakMap 讓 tr 攜帶欄寬
const _rw = new WeakMap();
const tr = (cells, hOrW=400, h2=null) => {
  const isArr = Array.isArray(hOrW);
  const h = isArr ? (h2||400) : hOrW;
  const row = new TableRow({ height:{value:h,rule:'atleast'}, children:cells });
  if(isArr) _rw.set(row, hOrW);
  return row;
};
const tbl = (rows, widths) => {
  const w = (widths.length===1 && rows.length && _rw.has(rows[0]))
    ? _rw.get(rows[0]) : widths;
  return new Table({
    width:{size:w.reduce((a,b)=>a+b,0), type:WidthType.DXA},
    columnWidths:w, rows
  });
};

// label cell (藍底)
const lc = (t, w, extra={}) => new TableCell({
  borders:aB(), width:{size:w,type:WidthType.DXA},
  shading:{fill:C_FIL,type:ShadingType.CLEAR},
  verticalAlign:VerticalAlign.CENTER,
  margins:{top:60,bottom:60,left:100,right:100},
  ...extra,
  children:[par([brun(t,{size:19,color:C_HDR})],{alignment:AlignmentType.CENTER})]
});
// value cell (空白)
const vc = (t, w, extra={}) => new TableCell({
  borders:aB(), width:{size:w,type:WidthType.DXA},
  verticalAlign:VerticalAlign.CENTER,
  margins:{top:60,bottom:60,left:120,right:60},
  ...extra,
  children:[par([run(t)])]
});
// 深藍 header cell
const hc = (t, w, extra={}) => new TableCell({
  borders:aB(), width:{size:w,type:WidthType.DXA},
  shading:{fill:C_SUB,type:ShadingType.CLEAR},
  verticalAlign:VerticalAlign.CENTER,
  margins:{top:60,bottom:60,left:100,right:100},
  ...extra,
  children:[par([hrun(t,19,'FFFFFF')],{alignment:AlignmentType.CENTER})]
});

// ── 表單大標 ────────────────────────────────────────────────
function fHdr(num, sub, title) {
  return [
    tbl([
      tr([new TableCell({
        borders:aB(), width:{size:W,type:WidthType.DXA},
        shading:{fill:C_HDR,type:ShadingType.CLEAR},
        margins:{top:80,bottom:20,left:140,right:140},
        children:[
          par([hrun('○○高中　教務處設備組',18,'FFFFFF',false)],{alignment:AlignmentType.CENTER}),
          par([hrun(title,28,'FFFFFF')],{alignment:AlignmentType.CENTER})
        ]
      })],600),
      tr([new TableCell({
        borders:aB(), width:{size:W,type:WidthType.DXA},
        margins:{top:30,bottom:30,left:140,right:140},
        children:[par([
          run(`表單編號：${num}　　`,{size:17,color:'555555'}),
          run(`版次：第　　版　　填表日期：　　　年　　月　　日`,{size:17,color:'555555'})
        ])]
      })],300)
    ],[W]),
    sp(60)
  ];
}
// 簽核列
function sigTbl(sigs) {
  const w = Math.floor(W/sigs.length);
  const ws = sigs.map((_,i)=> i<sigs.length-1 ? w : W-w*(sigs.length-1));
  return tbl([
    tr(sigs.map((s,i)=>lc(s,ws[i])),360),
    tr(sigs.map((_,i)=>vc('',ws[i])),700)
  ],ws);
}
// 節標題
const secL = (t) => par([brun(t,{size:21,color:C_SUB})],{spacing:{before:100,after:30}});

// ═══════════════════════════════════════════════════════
const body = [];

// ─────────────────────────────────────────────────────────
// 4.1  主機房工作紀錄表
// ─────────────────────────────────────────────────────────
body.push(...fHdr('ICT-F01','4.1','主機房工作紀錄表'));

// 基本資料
body.push(
  tbl([
    tr([lc('巡查年月',1600),vc('',2000),lc('巡查日期',1600),vc('',2000),lc('巡查人員',1600),vc('',1946)],
       [1600,2000,1600,2000,1600,1946])
  ],[W]),
  sp(40)
);

// 一、機房環境狀況
body.push(
  secL('一、機房環境狀況'),
  tbl([
    tr([hc('檢查項目',2200),hc('規定標準',3000),hc('實測／現況',2000),hc('正常',600),hc('異常',600),hc('備註',1346)],
       [2200,3000,2000,600,600,1346]),
    tr([lc('室內溫度',2200),vc('18–25 °C',3000),vc('　　　°C',2000),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',1346)],480),
    tr([lc('室內濕度',2200),vc('40–60 %',3000),vc('　　　%',2000),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',1346)],480),
    tr([lc('不斷電系統（UPS）',2200),vc('正常供電、指示燈正常',3000),vc('',2000),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',1346)],480),
    tr([lc('穩壓器／電源設備',2200),vc('電壓穩定、無異常聲響',3000),vc('',2000),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',1346)],480),
    tr([lc('門禁管制設備',2200),vc('門鎖正常、無異常開啟紀錄',3000),vc('',2000),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',1346)],480),
    tr([lc('監視攝影設備',2200),vc('錄影正常、畫面清晰',3000),vc('',2000),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',1346)],480),
    tr([lc('二氧化碳滅火器',2200),vc('壓力正常、有效期內、外觀完整',3000),vc('',2000),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',1346)],480),
    tr([lc('機房無易燃物品',2200),vc('機房內無紙箱、雜物堆置',3000),vc('',2000),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',1346)],480),
    tr([lc('逃生路線',2200),vc('出口暢通、緊急照明正常',3000),vc('',2000),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',1346)],480),
  ],[W]),
  sp(40)
);

// 二、主要設備運作狀況
body.push(
  secL('二、主要設備運作狀況'),
  tbl([
    tr([hc('設備名稱',2400),hc('運作狀況',2400),hc('正常',600),hc('異常',600),hc('說明',3746)],
       [2400,2400,600,600,3746]),
    tr([lc('核心交換器（Core Switch）',2400),vc('',2400),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',3746)],480),
    tr([lc('防火牆／路由器',2400),vc('',2400),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',3746)],480),
    tr([lc('校網伺服器',2400),vc('',2400),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',3746)],480),
    tr([lc('NAS 備份設備',2400),vc('',2400),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),
        vc('□',600,{children:[par([run('□')],{alignment:AlignmentType.CENTER})]}),vc('',3746)],480),
  ],[W]),
  sp(40)
);

// 三、檔案備份執行狀況
body.push(
  secL('三、檔案備份執行狀況'),
  tbl([
    tr([hc('備份項目',2400),hc('排定時間',1600),hc('實際完成時間',1800),hc('備份結果',1200),hc('備份媒介確認',2746)],
       [2400,1600,1800,1200,2746]),
    tr([lc('校網資料（差異備份）',2400),vc('每日 23:00',1600),vc('',1800),
        vc('□完成 □失敗',1200),vc('□本機 □NAS',2746)],480),
    tr([lc('校網資料（完整備份）',2400),vc('每週五 22:00',1600),vc('',1800),
        vc('□完成 □失敗',1200),vc('□NAS □雲端',2746)],480),
    tr([lc('網路設備組態',2400),vc('每週五 21:00',1600),vc('',1800),
        vc('□完成 □失敗',1200),vc('□NAS',2746)],480),
  ],[W]),
  sp(40)
);

// 四、異常情形及處理紀錄
body.push(
  secL('四、異常情形及處理紀錄（無異常請填「無」）'),
  tbl([
    tr([hc('異常項目',2000),hc('發生原因',3000),hc('排除方法',2600),hc('處理人員',1200),hc('完成時間',946)],
       [2000,3000,2600,1200,946]),
    tr([vc('',2000),vc('',3000),vc('',2600),vc('',1200),vc('',946)],520),
    tr([vc('',2000),vc('',3000),vc('',2600),vc('',1200),vc('',946)],520),
    tr([vc('',2000),vc('',3000),vc('',2600),vc('',1200),vc('',946)],520),
  ],[W]),
  sp(80),
  sigTbl(['巡查人員','設備組長核閱','教務主任知悉'])
);

// ─────────────────────────────────────────────────────────
// 4.2  檔案備份計畫
// ─────────────────────────────────────────────────────────
body.push(PB(),...fHdr('ICT-F02','4.2','檔案備份計畫'));

// 計畫基本資料
body.push(
  tbl([
    tr([lc('計畫年度',2000),vc('',2400),lc('制定／更新日期',2400),vc('',2946)],
       [2000,2400,2400,2946]),
    tr([lc('制定人員',2000),vc('',2400),lc('版　　次',2400),vc('第　　版',2946)],
       [2000,2400,2400,2946]),
  ],[W]),
  sp(40)
);

// 一、備份政策說明
body.push(
  secL('一、備份政策說明'),
  tbl([
    tr([lc('備份原則',2000),vc('遵循 3-2-1 原則：3 份備份、2 種儲存媒介（本機磁碟＋NAS）、1 份存放異地或雲端。',7746)],
       [2000,7746]),
    tr([lc('備份對象範圍',2000),vc('本校自管系統資料，包含：校網（學校官網）資料、網路設備組態檔、設備組業務重要文件。\n國教署集中管理之校務系統（成績、學籍、人事等）由國教署負責備份，不在本計畫範疇。',7746)],
       [2000,7746]),
    tr([lc('備份責任人',2000),vc('設備組；每次備份執行後於主機房工作紀錄表記錄結果。',7746)],
       [2000,7746]),
  ],[W]),
  sp(40)
);

// 二、備份排程明細
body.push(
  secL('二、備份排程明細'),
  tbl([
    tr([hc('系統／資料名稱',2000),hc('備份類型',1200),hc('執行頻率',1200),hc('排定時間',1200),hc('儲存媒介',1400),hc('主要存放位置',2746)],
       [2000,1200,1200,1200,1400,2746]),
    tr([lc('校網（學校官網）資料',2000),vc('差異備份',1200),vc('每日',1200),vc('23:00',1200),vc('NAS',1400),vc('伺服器室 NAS',2746)],480),
    tr([lc('校網（學校官網）資料',2000),vc('完整備份',1200),vc('每週五',1200),vc('22:00',1200),vc('NAS＋雲端',1400),vc('NAS＋Google Drive',2746)],480),
    tr([lc('網路設備組態檔',2000),vc('完整備份',1200),vc('每週五',1200),vc('21:00',1200),vc('NAS',1400),vc('伺服器室 NAS',2746)],480),
    tr([lc('設備組業務文件',2000),vc('完整備份',1200),vc('每月1日',1200),vc('20:00',1200),vc('NAS＋雲端',1400),vc('NAS＋Google Drive',2746)],480),
    tr([vc('',2000),vc('',1200),vc('',1200),vc('',1200),vc('',1400),vc('',2746)],480),
    tr([vc('',2000),vc('',1200),vc('',1200),vc('',1200),vc('',1400),vc('',2746)],480),
  ],[W]),
  sp(40)
);

// 三、備份保留期限
body.push(
  secL('三、備份保留期限'),
  tbl([
    tr([hc('備份類型',2400),hc('保留期限',2400),hc('到期處理方式',4946)],
       [2400,2400,4946]),
    tr([lc('每日差異備份',2400),vc('保留最近 7 天',2400),vc('自動覆蓋或手動清除',4946)],480),
    tr([lc('每週完整備份',2400),vc('保留最近 4 週',2400),vc('第 5 週備份完成後刪除最舊一份',4946)],480),
    tr([lc('每月完整備份',2400),vc('保留最近 3 個月',2400),vc('超過 3 個月後安全清除',4946)],480),
    tr([lc('學年度封存備份',2400),vc('永久保存（歸檔）',2400),vc('存放於獨立媒介，標示清楚，不得覆蓋',4946)],480),
  ],[W]),
  sp(40)
);

// 四、異地及雲端備份設定
body.push(
  secL('四、異地／雲端備份設定'),
  tbl([
    tr([lc('雲端服務平台',2400),vc('',3200),lc('帳號管理人',1800),vc('',2346)],
       [2400,3200,1800,2346]),
    tr([lc('雲端存放路徑',2400),vc('',7346)],[2400,7346]),
    tr([lc('異地實體位置\n（如有）',2400),vc('',3200),lc('異地負責人',1800),vc('',2346)],
       [2400,3200,1800,2346]),
    tr([lc('異地媒介更換頻率',2400),vc('□ 每月　□ 每季　□ 每學期　□ 其他：',7346)],
       [2400,7346]),
  ],[W]),
  sp(40)
);

// 五、3-2-1 原則確認
body.push(
  secL('五、3-2-1 原則確認'),
  tbl([
    tr([lc('3 份備份',2000),vc('□ 已達成　　說明：',7746)],[2000,7746]),
    tr([lc('2 種儲存媒介',2000),vc('□ 已達成　　說明：',7746)],[2000,7746]),
    tr([lc('1 份異地／雲端',2000),vc('□ 已達成　　說明：',7746)],[2000,7746]),
  ],[W]),
  sp(40)
);

// 六、備份測試計畫
body.push(
  secL('六、備份還原測試計畫'),
  tbl([
    tr([hc('測試頻率',2000),hc('測試方式',4000),hc('負責人員',1800),hc('紀錄表單',1946)],
       [2000,4000,1800,1946]),
    tr([lc('每月一次（最低要求）',2000),vc('從 NAS 還原校網資料至測試環境，驗證資料完整性',4000),vc('設備組',1800),vc('備份還原測試紀錄表',1946)],520),
    tr([lc('每學期一次（完整演練）',2000),vc('模擬災害，執行完整 DRP 流程，確認 RTO ≤ 4 小時、RPO ≤ 24 小時',4000),vc('設備組',1800),vc('系統復原測試報告',1946)],560),
  ],[W]),
  sp(40)
);

// 七、備份異常處理
body.push(
  secL('七、備份異常處理程序'),
  tbl([
    tr([lc('發現備份異常',2400),vc('立即於主機房工作紀錄表記錄原因，通知設備組長。',7346)],[2400,7346]),
    tr([lc('連續 2 次以上失敗',2400),vc('設備組長通報教務主任，評估是否啟動緊急備份程序或採購替代媒介。',7346)],[2400,7346]),
    tr([lc('媒介損壞',2400),vc('立即更換媒介，確認備份資料可用性後於媒體銷毀紀錄表記錄損壞媒介銷毀情形。',7346)],[2400,7346]),
  ],[W]),
  sp(80),
  sigTbl(['制定人','設備組長審核','教務主任核准'])
);

// ── 建立文件 ─────────────────────────────────────────────────
const doc = new Document({
  numbering:{ config:[
    { reference:'b0', levels:[{ level:0, format:LevelFormat.BULLET, text:'●',
        alignment:AlignmentType.LEFT,
        style:{paragraph:{indent:{left:480,hanging:300}}} }] }
  ]},
  styles:{
    default:{ document:{ run:{ font:{name:FONT}, size:20, color:'1A1A1A' } } },
    paragraphStyles:[
      { id:'Heading1', name:'Heading 1', basedOn:'Normal', next:'Normal', quickFormat:true,
        run:{size:24,bold:true,font:{name:FONT},color:C_SUB},
        paragraph:{spacing:{before:240,after:80},outlineLevel:0} }
    ]
  },
  sections:[{
    properties:{
      page:{
        size:{ width:11906, height:16838 },
        margin:{ top:1080, right:1080, bottom:1080, left:1080 }
      }
    },
    headers:{
      default: new Header({
        children:[par([
          run('○○高中　教務處設備組　資訊內控表單',{size:17,color:'888888'}),
          new TextRun({children:['\t'],font:{name:FONT}}),
          new TextRun({children:['第 ',PageNumber.CURRENT,' 頁'],font:{name:FONT},size:17,color:'888888'}),
        ],{
          tabStops:[{type:TabStopType.RIGHT,position:TabStopPosition.MAX}],
          border:{bottom:{style:BorderStyle.SINGLE,size:4,color:'AEC9E0',space:4}}
        })]
      })
    },
    children:body
  }]
});

Packer.toBuffer(doc).then(buf=>{
  const out = 'D:\\D\\114設備組\\主機房工作紀錄表_檔案備份計畫.docx';
  fs.writeFileSync(out, buf);
  console.log('DONE', out);
}).catch(e=>{ console.error(e); process.exit(1); });
