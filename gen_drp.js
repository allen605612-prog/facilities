// 正心中學主機備援計畫（DRP）
// 依據 設備組內控修訂.docx ◎系統復原計畫及測試作業 2.1~2.3 產生
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, LevelFormat, TabStopType, TabStopPosition,
  PageBreak, HeadingLevel
} = require('C:/Users/user/AppData/Roaming/npm/node_modules/docx');
const fs = require('fs');

const FONT  = '標楷體';
const C_DEEP = '1F3864';
const C_MID  = '2E5F8A';
const C_LITE = '2E75B6';
const C_FILL = 'D6E4F0';
const C_SUB2 = 'EBF3FB';
const C_RED  = '9E1B1B';
const C_BORD = 'AEC9E0';
const W      = 9360;   // A4 1吋邊距 content width

// ── 基礎元件 ─────────────────────────────────────────────────
const r  = (t, o={}) => new TextRun({ text:t, font:{name:FONT}, size:22, ...o });
const rb = (t, o={}) => r(t, {bold:true, ...o});
const rh = (t, sz, col, bold=true) => new TextRun({text:t,font:{name:FONT},size:sz,color:col,bold});
const p  = (ch, o={}) => new Paragraph({children:Array.isArray(ch)?ch:[ch], ...o});
const sp = (b=80)  => p([], {spacing:{before:b,after:0}});
const PB = ()      => p([new PageBreak()]);

const H1 = (t) => p([rh(t,28,C_MID,true)], {
  heading: HeadingLevel.HEADING_1,
  spacing:{before:400,after:120},
  border:{bottom:{style:BorderStyle.SINGLE,size:6,color:C_MID,space:4}}
});
const H2 = (t) => p([rh(t,24,C_LITE,true)], {
  heading: HeadingLevel.HEADING_2,
  spacing:{before:280,after:80}
});
const H3 = (t) => p([rb(t,{size:22,color:'333333'})], {spacing:{before:160,after:60}});

const item = (num, text, bold=false) => p(
  [rb(`${num}　`,{size:22,color:C_LITE}), bold?rb(text):r(text)],
  {indent:{left:360}, spacing:{after:80}}
);
const bul = (text) => p([r(text)], {
  numbering:{reference:'b0',level:0}, spacing:{after:60}
});

// ── 表格工廠 ─────────────────────────────────────────────────
const bd  = (c=C_BORD) => ({style:BorderStyle.SINGLE, size:4, color:c});
const aB  = () => ({top:bd(),bottom:bd(),left:bd(),right:bd()});
const nB  = () => ({top:bd('FFFFFF'),bottom:bd('FFFFFF'),left:bd('FFFFFF'),right:bd('FFFFFF')});

// label cell
const lc = (t, w, span=1) => new TableCell({
  borders:aB(), width:{size:w,type:WidthType.DXA},
  columnSpan:span,
  shading:{fill:C_FILL,type:ShadingType.CLEAR},
  verticalAlign:VerticalAlign.CENTER,
  margins:{top:60,bottom:60,left:120,right:80},
  children:[p([rb(t,{size:20,color:C_DEEP})],{alignment:AlignmentType.CENTER})]
});
// value cell
const vc = (t, w, span=1, align=AlignmentType.LEFT) => new TableCell({
  borders:aB(), width:{size:w,type:WidthType.DXA},
  columnSpan:span,
  verticalAlign:VerticalAlign.CENTER,
  margins:{top:60,bottom:60,left:120,right:80},
  children:[p([r(t)],{alignment:align})]
});
// header cell
const hc = (t, w, span=1) => new TableCell({
  borders:aB(), width:{size:w,type:WidthType.DXA},
  columnSpan:span,
  shading:{fill:C_MID,type:ShadingType.CLEAR},
  verticalAlign:VerticalAlign.CENTER,
  margins:{top:80,bottom:80,left:120,right:80},
  children:[p([rh(t,20,'FFFFFF')],{alignment:AlignmentType.CENTER})]
});
// 優先級 badge
const pri = (t, col, w) => new TableCell({
  borders:aB(), width:{size:w,type:WidthType.DXA},
  shading:{fill:col,type:ShadingType.CLEAR},
  verticalAlign:VerticalAlign.CENTER,
  margins:{top:60,bottom:60,left:80,right:80},
  children:[p([rb(t,{size:20,color:'FFFFFF'})],{alignment:AlignmentType.CENTER})]
});

function mkTbl(headers, rows, widths) {
  const tot = widths.reduce((a,b)=>a+b,0);
  return new Table({
    width:{size:tot,type:WidthType.DXA}, columnWidths:widths,
    rows:[
      new TableRow({tableHeader:true, children:headers.map((h,i)=>hc(h,widths[i]))}),
      ...rows.map((cols,ri)=>new TableRow({
        children:cols.map((c,i)=>new TableCell({
          borders:aB(), width:{size:widths[i],type:WidthType.DXA},
          shading:ri%2===1?{fill:C_SUB2,type:ShadingType.CLEAR}:undefined,
          verticalAlign:VerticalAlign.CENTER,
          margins:{top:60,bottom:60,left:120,right:80},
          children:[p([r(c)],{alignment:i===0?AlignmentType.CENTER:AlignmentType.LEFT})]
        }))
      }))
    ]
  });
}
function infoTbl(rows) {   // 兩欄 label/value 資訊表
  const w1=2800, w2=W-w1;
  return new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[w1,w2],
    rows: rows.map(([l,v])=>new TableRow({
      children:[lc(l,w1), vc(v,w2)]
    }))
  });
}
function flowTbl(steps) {
  const w = Math.floor(W/steps.length);
  const ws = steps.map((_,i)=>i<steps.length-1?w:W-w*(steps.length-1));
  return new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:ws,
    rows:[new TableRow({children:steps.map((s,i)=>new TableCell({
      borders:aB(), width:{size:ws[i],type:WidthType.DXA},
      shading:{fill:i%2===0?C_FILL:C_SUB2,type:ShadingType.CLEAR},
      verticalAlign:VerticalAlign.CENTER,
      margins:{top:120,bottom:120,left:60,right:60},
      children:[p([r(s,{size:19})],{alignment:AlignmentType.CENTER})]
    }))})]
  });
}

// ════════════════════════════════════════════════════════════
const body = [];

// ── 封面 ────────────────────────────────────────────────────
body.push(
  sp(1200),
  p([rh('正心中學',56,C_DEEP)], {alignment:AlignmentType.CENTER}),
  p([rh('主機備援計畫',52,C_MID)], {alignment:AlignmentType.CENTER}),
  p([rh('Disaster Recovery Plan (DRP)',24,'888888',false)], {alignment:AlignmentType.CENTER}),
  sp(300),
  new Table({
    width:{size:6000,type:WidthType.DXA}, columnWidths:[2400,3600],
    rows:[
      new TableRow({children:[lc('主管單位',2400), vc('教務處　設備組',3600)]}),
      new TableRow({children:[lc('文件版次',2400), vc('第 1.0 版',3600)]}),
      new TableRow({children:[lc('訂定日期',2400), vc('中華民國　　　年　　月　　日',3600)]}),
      new TableRow({children:[lc('最近更新',2400), vc('中華民國　　　年　　月　　日',3600)]}),
      new TableRow({children:[lc('核准主管',2400), vc('教務主任：',3600)]}),
    ]
  }),
  sp(200),
  p([r('本計畫依「資訊處理事項內部控制制度」◎系統復原計畫及測試作業訂定，',{size:19,color:'666666'}),
     r('每學年至少更新一次，重大環境變更後應立即修訂。',{size:19,color:'666666'})],
    {alignment:AlignmentType.CENTER}),
  PB()
);

// ── 目錄提示 ────────────────────────────────────────────────
body.push(
  H1('文件索引'),
  mkTbl(
    ['章節','標題','頁次'],
    [
      ['一','目的與適用範圍','—'],
      ['二','本校資訊系統環境說明','—'],
      ['三','關鍵系統與復原目標（RTO／RPO）','—'],
      ['四','緊急聯絡名單','—'],
      ['五','復原優先順序與流程','—'],
      ['六','備援設備與資源清單','—'],
      ['七','一般故障處理程序','—'],
      ['八','重大故障及資安事件應變程序','—'],
      ['九','備份媒體管理','—'],
      ['十','備援計畫測試與維護','—'],
      ['十一','相關表單與依據法規','—'],
    ],
    [1200,6000,2160]
  ),
  sp(40),
  PB()
);

// ── 第一章 ───────────────────────────────────────────────────
body.push(
  H1('第一章　目的與適用範圍'),
  H2('1.1 目的'),
  p([r('本計畫旨在確保正心中學資訊系統於發生硬體故障、軟體損毀、資安事件或天然災害時，能依照預定程序在最短時間內恢復服務，降低對教學與行政作業之衝擊，並符合「資訊處理事項內部控制制度」之要求。')]),
  sp(60),
  H2('1.2 適用範圍'),
  p([r('本計畫適用範圍為本校設備組直接管理之資訊系統與設備：')]),
  bul('校園網路基礎設施（核心交換器、防火牆、路由器、無線 AP）'),
  bul('學校官網伺服器（校網）'),
  bul('NAS 備份儲存設備'),
  bul('電腦教室及行政電腦相關網路服務'),
  sp(40),
  p([rb('注意：',{color:C_RED}),
     r('國教署集中管理之校務行政系統（成績、學籍、人事、財務等）由國教署負責其備援計畫；本校設備組之責任為確保本校網路連線能於 RTO 規定時間內恢復，使各處室得以正常存取上述系統。')]),
  sp(40),
  PB()
);

// ── 第二章 ───────────────────────────────────────────────────
body.push(
  H1('第二章　本校資訊系統環境說明'),
  mkTbl(
    ['類別','系統／設備','管理主體','備援責任'],
    [
      ['本校自管','校網伺服器、網路核心設備、NAS','設備組','本計畫全部涵蓋'],
      ['本校自管','電腦教室磁碟還原系統、MDM','設備組','本計畫涵蓋'],
      ['國教署集中','成績、學籍、人事、財務等校務系統','國教署','國教署 DRP；本校僅確保網路連線'],
    ],
    [1400,3400,2000,2560]
  ),
  sp(40),
  PB()
);

// ── 第三章 ───────────────────────────────────────────────────
body.push(
  H1('第三章　關鍵系統與復原目標（RTO／RPO）'),
  p([r('各關鍵系統之復原目標如下表。'), rb(' RTO',{color:C_RED}), r('（Recovery Time Objective，最長可接受停機時間）；'), rb(' RPO',{color:C_RED}), r('（Recovery Point Objective，最大可接受資料遺失區間）。')]),
  sp(60),
  mkTbl(
    ['優先級','系統／服務','RTO 目標','RPO 目標','備援方式','主責人員'],
    [
      ['P1','網路核心設備（交換器、防火牆）','2 小時','即時（組態備份）','備援設備熱切換','設備組長'],
      ['P1','校園網際網路連線','2 小時','即時','聯絡 ISP 恢復','設備組長'],
      ['P2','校網伺服器（學校官網）','4 小時','24 小時','備份還原至備援主機','設備組'],
      ['P2','NAS 備份系統','4 小時','24 小時','備援 NAS 或雲端還原','設備組'],
      ['P3','國教署校務系統（使用端）','依國教署 SLA','依國教署 SLA','確保網路連線恢復','設備組（網路端）'],
      ['P3','電腦教室服務','8 小時','當日備份','磁碟還原重置','設備組'],
    ],
    [800,2600,1400,1400,2000,1160]
  ),
  sp(40),
  PB()
);

// ── 第四章 ───────────────────────────────────────────────────
body.push(
  H1('第四章　緊急聯絡名單'),
  p([r('發生系統故障或重大事件時，依下列順序通報。緊急聯絡名單每學年開學前更新，並於重大人事異動後立即修訂。')]),
  sp(60),
  mkTbl(
    ['優先序','角色','姓名','單位分機','手機','電子郵件'],
    [
      ['1','設備組長（第一通報）','','','',''],
      ['2','教務主任','','','',''],
      ['3','設備組人員（支援）','','','',''],
      ['4','校長','','','',''],
      ['5','網路廠商（ISP）','','','',''],
      ['6','網路設備維護廠商','','','',''],
      ['7','校網伺服器廠商','','','',''],
      ['8','縣市政府教育局資訊人員','','','',''],
    ],
    [600,2200,1400,1400,1760,1960]
  ),
  sp(40),
  H2('4.1 重大資安事件額外通報'),
  mkTbl(
    ['事件類型','通報對象','通報時限','通報管道'],
    [
      ['任何資安事件','設備組長 → 教務主任','立即','電話'],
      ['重大資安事件（影響多台主機、資料外洩）','教育局 / 教育部資安通報平台','1 小時內初報','教育部資安通報系統'],
      ['勒索軟體或系統入侵','上述＋校長','1 小時內','電話＋書面'],
    ],
    [2000,2600,1360,3400]
  ),
  sp(40),
  PB()
);

// ── 第五章 ───────────────────────────────────────────────────
body.push(
  H1('第五章　復原優先順序與流程'),
  H2('5.1 復原優先順序'),
  p([r('依系統對學校核心業務影響程度，復原工作依下列優先順序執行：')]),
  sp(40),
  new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[800,2600,5960],
    rows:[
      new TableRow({tableHeader:true, children:[hc('優先級',800),hc('系統',2600),hc('說明',5960)]}),
      new TableRow({children:[
        pri('P1','C00000',800),
        vc('網路核心設備\n（必須最先恢復）',2600),
        vc('無網路則全校所有系統（含國教署校務系統）均無法使用，優先恢復。',5960)
      ]}),
      new TableRow({children:[
        pri('P1','C00000',800),
        vc('網際網路連線',2600),
        vc('確保與國教署資料中心連線，各處室方能繼續使用校務系統。',5960)
      ]}),
      new TableRow({children:[
        pri('P2','E07020',800),
        vc('校網伺服器',2600),
        vc('學校官網、對外服務。',5960)
      ]}),
      new TableRow({children:[
        pri('P2','E07020',800),
        vc('NAS 備份系統',2600),
        vc('確保備份資料可用，支援其他系統後續還原作業。',5960)
      ]}),
      new TableRow({children:[
        pri('P3','2E75B6',800),
        vc('電腦教室服務',2600),
        vc('依課表排定優先恢復有課班級。',5960)
      ]}),
    ]
  }),
  sp(60),

  H2('5.2 故障應變總流程'),
  flowTbl(['發現故障\n／通報','設備組\n30 分鐘確認','可自修？\n→ 是：修復','→ 否：\n洽廠商或\n啟動備援','修復後\n資料驗證','紀錄歸檔\n通知使用者']),
  sp(40),
  PB()
);

// ── 第六章 ───────────────────────────────────────────────────
body.push(
  H1('第六章　備援設備與資源清單'),
  p([r('下列備援設備應每學期確認可用性，並於確認後於本表記錄日期。')]),
  sp(40),
  mkTbl(
    ['設備名稱','數量','存放位置','負責人','最近確認日期','備註'],
    [
      ['備援交換器（Managed）','1 台','伺服器室備用架','設備組','',''],
      ['備援防火牆','1 台','伺服器室備用架','設備組','',''],
      ['備援 NAS（或可攜式 HDD）','1 套','伺服器室保險箱','設備組','',''],
      ['UPS（備援電池模組）','依機房配置','伺服器室','設備組','','每年更換電池'],
      ['緊急備用筆電（設備組用）','1 台','設備組辦公室','設備組長','',''],
      ['原廠維修合約文件','1 份','設備組辦公室文件夾','設備組長','','含廠商聯絡資訊'],
    ],
    [2000,800,1800,1200,1560,1960]
  ),
  sp(40),
  H2('6.1 重要軟體與文件備份清單'),
  p([r('下列重要軟體授權與文件應定期抄錄備份至安全場所（詳見 4.2 檔案備份計畫）：')]),
  bul('系統軟體安裝介質（ISO 映像）及授權金鑰'),
  bul('網路設備組態備份檔（每週自動備份至 NAS）'),
  bul('校網伺服器完整快照（每週完整備份）'),
  bul('設備資產清冊（每學期更新版本）'),
  bul('各廠商維護合約及 SLA 文件掃描檔'),
  sp(40),
  PB()
);

// ── 第七章 ───────────────────────────────────────────────────
body.push(
  H1('第七章　一般故障處理程序'),
  H2('7.1 使用者報修流程'),
  p([r('一般硬體、軟體或網路故障，依下列步驟處理：')]),
  sp(40),
  item('步驟 1', '使用者於維修管理系統（或紙本）填具「維修申請單」，說明故障設備、位置及現象。'),
  item('步驟 2', '設備組人員於收到通報後 30 分鐘內確認問題，回覆使用者預計處理時間。'),
  item('步驟 3', '判斷故障類型：'),
  p([r('　　(a) 可自行修復：由設備組人員排除，填寫完成時間與處理說明。')],{indent:{left:720},spacing:{after:60}}),
  p([r('　　(b) 需送外維修：記錄送修日期、廠商及預計返回日期；送修前確認敏感資料已清除。')],{indent:{left:720},spacing:{after:60}}),
  p([r('　　(c) 需啟動備援：依第五章優先順序，動用第六章備援設備。')],{indent:{left:720},spacing:{after:60}}),
  item('步驟 4', '於 4 小時內完成修復或啟動備援，使系統恢復至可用狀態。'),
  item('步驟 5', '修復完成後，執行資料完整性驗證，確認無資料遺失。'),
  item('步驟 6', '將處理結果回覆至維修管理系統，通知使用者，並將維修紀錄歸檔。'),
  sp(40),

  H2('7.2 設備送外維修規定'),
  item('7.2.1', '設備送外維修時，應由設備組人員全程陪同，或取得書面委託紀錄。'),
  item('7.2.2', '含有個人資料或機密資料之儲存媒介，送修前必須完成安全清除（低階格式化或加密抹除），並記錄於維修申請單。'),
  item('7.2.3', '記錄送修日期、廠商名稱、預計完成日期及設備序號，回收時核對。'),
  sp(40),
  PB()
);

// ── 第八章 ───────────────────────────────────────────────────
body.push(
  H1('第八章　重大故障及資安事件應變程序'),
  H2('8.1 重大硬體故障（無法自行修復）'),
  item('8.1.1', '立即通報設備組長，由設備組長判斷是否啟動緊急應變小組。'),
  item('8.1.2', '緊急應變小組組成：設備組長（召集）＋設備組人員＋教務主任（通報）＋相關處室主任（視需要）。'),
  item('8.1.3', '聯繫原採購廠商，說明故障現象，要求提供緊急備援設備（借用或租用），確保核心業務不中斷。'),
  item('8.1.4', '若設備損壞無法修復，立即依預算採購相容設備；由設備組人員執行資料回存。'),
  item('8.1.5', '重大事故應簽訂系統復原合約，合約內容應包含：修護完成交期、保固期間、違約損失賠償罰則及應變方式。'),
  sp(40),

  H2('8.2 資安事件（惡意程式、入侵、資料外洩）'),
  item('步驟 1', '立即隔離受感染或受入侵之系統，切斷其網路連線，防止橫向擴散。', true),
  item('步驟 2', '通報設備組長及教務主任（即時）。'),
  item('步驟 3', '保全現場數位證據，不得關機或修改任何系統設定（除非隔離需要）。'),
  item('步驟 4', '判斷事件等級：'),
  p([r('　　・一般事件（單台電腦受感染）：設備組自行處理，填寫資安事件通報紀錄表。')],{indent:{left:720},spacing:{after:60}}),
  p([r('　　・重大事件（多台主機、資料外洩、勒索軟體）：1 小時內向教育部資安通報平台初報，24 小時內完成詳報。')],{indent:{left:720},spacing:{after:60}}),
  item('步驟 5', '執行清除與復原作業（重灌、從乾淨備份還原），完成後重新連線。'),
  item('步驟 6', '事後追查根本原因，訂定改善措施並更新本計畫及相關安全設定。'),
  sp(40),

  H2('8.3 自然災害或重大意外（停電、淹水、火災）'),
  item('8.3.1', '依學校緊急應變計畫執行疏散及安全措施；優先確保人員安全。'),
  item('8.3.2', '災後由設備組評估設備損壞狀況，依本計畫第五章優先順序啟動復原。'),
  item('8.3.3', '若機房設備全損，向縣市教育局申請緊急資源支援，並利用雲端備份資料重建服務。'),
  sp(40),
  PB()
);

// ── 第九章 ───────────────────────────────────────────────────
body.push(
  H1('第九章　備份媒體管理'),
  H2('9.1 備份媒體分類與標示'),
  mkTbl(
    ['媒體類型','使用用途','標示規定','存放位置'],
    [
      ['NAS 硬碟','每日差異備份、每週完整備份','標示：系統名稱、備份日期範圍','伺服器室 NAS 設備'],
      ['外接硬碟（加密）','每月完整備份、異地備份','標示：年月、系統名稱、加密確認','設備組辦公室保險箱'],
      ['雲端（Google Drive for Education）','每週完整備份、學年歸檔','資料夾命名規則：YYYYMM_系統名','雲端（帳號由設備組長管理）'],
    ],
    [2000,2600,2800,1960]
  ),
  sp(40),

  H2('9.2 媒體保留期限'),
  mkTbl(
    ['備份類型','保留期限','到期處理'],
    [
      ['每日差異備份','7 天','自動覆蓋'],
      ['每週完整備份','4 週','第 5 週後刪除最舊份'],
      ['每月完整備份','3 個月','超過後安全清除'],
      ['學年度歸檔','永久','存放獨立媒介，不得覆蓋'],
    ],
    [2800,2400,4160]
  ),
  sp(40),

  H2('9.3 廢棄媒體銷毀'),
  item('9.3.1', '廢棄含個資或機密資料之儲存媒介（硬碟、隨身碟、光碟），應實體銷毀（碎磁機、物理穿孔）或執行低階格式化（覆寫 7 次）。'),
  item('9.3.2', '銷毀過程須留有「媒體銷毀紀錄表」，記錄：銷毀日期、媒體類型、數量、原存放系統、銷毀方式及執行人員。'),
  item('9.3.3', '銷毀作業須有見證人，紀錄表送設備組長核閱後歸檔。'),
  sp(40),
  PB()
);

// ── 第十章 ───────────────────────────────────────────────────
body.push(
  H1('第十章　備援計畫測試與維護'),
  H2('10.1 測試計畫'),
  mkTbl(
    ['測試類型','頻率','測試內容','紀錄表單'],
    [
      ['備份還原測試','每月一次（最低）','從 NAS 還原校網資料至測試環境，驗證完整性','備份還原測試紀錄表（ICT-F03）'],
      ['網路切換演練','每學期一次','切換至備援交換器，確認服務不中斷','系統復原測試報告（ICT-F12）'],
      ['完整 DRP 演練','每學期一次','模擬重大故障，執行完整流程，確認 RTO ≤ 4 hr','系統復原測試報告（ICT-F12）'],
      ['緊急聯絡測試','每學年一次（開學前）','確認聯絡名單所有人員可聯繫','更新第四章名單'],
    ],
    [2000,1600,3800,2160]
  ),
  sp(40),

  H2('10.2 測試結果處理'),
  item('10.2.1', '每次測試完成後，設備組撰寫「系統復原測試報告」，包含：測試情境、測試結果（是否達到 RTO/RPO）、發現問題及改善措施。'),
  item('10.2.2', '測試報告送教務主任核閱後建檔；發現問題應於下次測試前完成改善。'),
  item('10.2.3', '暫存於其他系統之測試資料，於確認完整回存後須安全銷毀。'),
  sp(40),

  H2('10.3 計畫維護'),
  item('10.3.1', '本計畫每學年至少全面審視更新一次（建議：每年 8 月開學前）。'),
  item('10.3.2', '發生下列情形時，應立即修訂本計畫：'),
  p([r('　　・更換核心網路設備或伺服器　・新增或下線重要系統　・人員異動（聯絡名單）　・發生實際災害復原事件後')],{indent:{left:720},spacing:{after:80}}),
  item('10.3.3', '計畫修訂後，應通知相關人員並重新確認緊急聯絡名單。'),
  sp(40),
  PB()
);

// ── 第十一章 ─────────────────────────────────────────────────
body.push(
  H1('第十一章　相關表單與依據法規'),
  H2('11.1 相關表單'),
  mkTbl(
    ['表單名稱','表單編號','使用時機'],
    [
      ['維修申請單','ICT-F10','使用者通報故障時填具'],
      ['系統復原計畫（DRP）本文','ICT-F11','每學年更新，重大變更後修訂'],
      ['系統復原測試報告','ICT-F12','每學期演練後填寫'],
      ['備份還原測試紀錄表','ICT-F03','每月備份還原測試後填寫'],
      ['媒體銷毀紀錄表','ICT-F04','廢棄媒體銷毀時填具'],
      ['資安事件通報紀錄表','ICT-F14','發生資安事件時填具'],
    ],
    [3000,1800,4560]
  ),
  sp(40),

  H2('11.2 依據法規與相關文件'),
  bul('資通安全管理法（民國 107 年）'),
  bul('個人資料保護法'),
  bul('教育部校園資安事件通報規定'),
  bul('教育部資安事件分級及通報處理辦法'),
  bul('本校資訊處理事項內部控制制度'),
  bul('本校緊急應變計畫'),
  bul('本校電腦機房管理辦法'),
  sp(80),

  H2('11.3 文件核准欄'),
  new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[2340,2340,2340,2340],
    rows:[
      new TableRow({children:[hc('制定人',2340),hc('設備組長審核',2340),hc('教務主任核准',2340),hc('校長知悉',2340)]}),
      new TableRow({children:[vc('',2340),vc('',2340),vc('',2340),vc('',2340)]}),
      new TableRow({children:[
        new TableCell({borders:aB(),width:{size:2340,type:WidthType.DXA},margins:{top:40,bottom:40,left:80,right:80},children:[p([r('日期：　　　／　　／　　',{size:18})])]}),
        new TableCell({borders:aB(),width:{size:2340,type:WidthType.DXA},margins:{top:40,bottom:40,left:80,right:80},children:[p([r('日期：　　　／　　／　　',{size:18})])]}),
        new TableCell({borders:aB(),width:{size:2340,type:WidthType.DXA},margins:{top:40,bottom:40,left:80,right:80},children:[p([r('日期：　　　／　　／　　',{size:18})])]}),
        new TableCell({borders:aB(),width:{size:2340,type:WidthType.DXA},margins:{top:40,bottom:40,left:80,right:80},children:[p([r('日期：　　　／　　／　　',{size:18})])]}),
      ]}),
    ]
  })
);

// ── 建立文件 ─────────────────────────────────────────────────
const doc = new Document({
  numbering:{config:[
    {reference:'b0', levels:[{level:0, format:LevelFormat.BULLET, text:'●',
      alignment:AlignmentType.LEFT,
      style:{paragraph:{indent:{left:480,hanging:300}}}}]}
  ]},
  styles:{
    default:{document:{run:{font:{name:FONT},size:22,color:'1A1A1A'}}},
    paragraphStyles:[
      {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,
        run:{size:28,bold:true,font:{name:FONT},color:C_MID},
        paragraph:{spacing:{before:400,after:120},outlineLevel:0}},
      {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,
        run:{size:24,bold:true,font:{name:FONT},color:C_LITE},
        paragraph:{spacing:{before:240,after:80},outlineLevel:1}}
    ]
  },
  sections:[{
    properties:{
      page:{
        size:{width:11906,height:16838},
        margin:{top:1440,right:1080,bottom:1440,left:1440}
      }
    },
    headers:{
      default: new Header({children:[p(
        [r('正心中學　主機備援計畫（DRP）',{size:18,color:'888888'}),
         new TextRun({children:['\t'],font:{name:FONT}}),
         new TextRun({children:['第 ',PageNumber.CURRENT,' 頁'],font:{name:FONT},size:18,color:'888888'})],
        {tabStops:[{type:TabStopType.RIGHT,position:TabStopPosition.MAX}],
         border:{bottom:{style:BorderStyle.SINGLE,size:4,color:C_BORD,space:4}}}
      )]})
    },
    footers:{
      default: new Footer({children:[p(
        [r('本文件為正心中學內部管理文件，未經授權不得對外揭露。',{size:16,color:'AAAAAA'})],
        {alignment:AlignmentType.CENTER}
      )]})
    },
    children:body
  }]
});

Packer.toBuffer(doc).then(buf=>{
  const out = 'D:\\D\\114設備組\\正心中學主機備援計畫.docx';
  fs.writeFileSync(out,buf);
  console.log('DONE', out);
}).catch(e=>{console.error(e);process.exit(1);});
