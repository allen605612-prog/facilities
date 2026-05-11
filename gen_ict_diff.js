const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat,
  TabStopType, TabStopPosition, PageBreak
} = require('C:/Users/user/AppData/Roaming/npm/node_modules/docx');
const fs = require('fs');

const FONT   = '標楷體';
const C_DEEP = '1F3864';
const C_MID  = '2E5F8A';
const C_LITE = '2E75B6';
const C_BORD = 'AEC9E0';
const C_GRAY = '808080';
const C_RED  = 'C00000';

const run  = (text, opts = {}) => new TextRun({ text, font: { name: FONT }, size: 22, ...opts });
const hrun = (text, size, color, bold = true) => new TextRun({ text, font: { name: FONT }, size, color, bold });

// diff helpers
const rDel = (text) => run(text, { strike: true, color: C_GRAY });
const rNew = (text) => run(text, { color: C_RED, bold: true });

const par = (children, opts = {}) =>
  new Paragraph({ children: Array.isArray(children) ? children : [children], ...opts });
const sp = (before = 80, after = 0) => par([], { spacing: { before, after } });

const bpar = (children) => par(
  Array.isArray(children) ? children : [children],
  { numbering: { reference: 'b0', level: 0 }, spacing: { after: 60 } }
);

// numbered item with state: 'same' | 'del' | 'new'
const numItem = (num, text, state = 'same') => {
  const nc = state === 'del' ? C_GRAY : state === 'new' ? C_RED : C_LITE;
  const tr = state === 'del' ? rDel(text) : state === 'new' ? rNew(text) : run(text);
  return par(
    [new TextRun({ text: `${num} `, font: { name: FONT }, size: 22, bold: true, color: nc, strike: state === 'del' }), tr],
    { indent: { left: 360 }, spacing: { after: 80 } }
  );
};

const border  = (c = C_BORD) => ({ style: BorderStyle.SINGLE, size: 4, color: c });
const allB    = { top: border(), bottom: border(), left: border(), right: border() };

function flowTbl(steps, state = 'same') {
  const col = state === 'del' ? C_GRAY : state === 'new' ? C_RED : C_LITE;
  const f0  = state === 'del' ? 'EEEEEE' : state === 'new' ? 'FFE7E7' : 'D6E4F0';
  const f1  = state === 'del' ? 'F5F5F5' : state === 'new' ? 'FFD0D0' : 'EBF3FB';
  const w   = Math.floor(8800 / steps.length);
  const ws  = steps.map((_, i) => i < steps.length - 1 ? w : 8800 - w * (steps.length - 1));
  return new Table({
    width: { size: 8800, type: WidthType.DXA },
    columnWidths: ws,
    rows: [new TableRow({
      children: steps.map((s, i) => new TableCell({
        borders: allB,
        width: { size: ws[i], type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? f0 : f1, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 60, right: 60 },
        verticalAlign: VerticalAlign.CENTER,
        children: [par([run(s, { color: col, strike: state === 'del' })], { alignment: AlignmentType.CENTER })]
      }))
    })]
  });
}

// rowStates: array of 'same'|'del'|'new' per row
function mkTbl(headers, rows, widths, rowStates = []) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => new TableCell({
          borders: allB,
          width: { size: widths[i], type: WidthType.DXA },
          shading: { fill: '2E5F8A', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          verticalAlign: VerticalAlign.CENTER,
          children: [par([hrun(h, 20, 'FFFFFF', true)], { alignment: AlignmentType.CENTER })]
        }))
      }),
      ...rows.map((cols, ri) => {
        const st = rowStates[ri] || 'same';
        const bg = st === 'del' ? 'EEEEEE' : st === 'new' ? 'FFE7E7' : (ri % 2 === 1 ? 'EBF3FB' : undefined);
        return new TableRow({
          children: cols.map((c, i) => new TableCell({
            borders: allB,
            width: { size: widths[i], type: WidthType.DXA },
            shading: bg ? { fill: bg, type: ShadingType.CLEAR } : undefined,
            margins: { top: 60, bottom: 60, left: 120, right: 120 },
            children: [par(
              [st === 'del' ? rDel(c) : st === 'new' ? rNew(c) : run(c)],
              { alignment: i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT }
            )]
          }))
        });
      })
    ]
  });
}

const H1 = (text, state = 'same') => {
  const c = state === 'del' ? C_GRAY : state === 'new' ? C_RED : C_MID;
  return par([hrun(text, 30, c, true)], {
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: c, space: 4 } }
  });
};
const H2 = (text, state = 'same') => {
  const c = state === 'del' ? C_GRAY : state === 'new' ? C_RED : C_LITE;
  return par([hrun(text, 24, c, true)], { heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 80 } });
};
const H3 = (text, state = 'same') => {
  const c = state === 'del' ? C_GRAY : state === 'new' ? C_RED : '333333';
  return par([hrun(text, 22, c, true)], { spacing: { before: 160, after: 60 } });
};
const PAGE_BREAK = () => par([new PageBreak()]);

// ═══════════════════════════════════════════════════════════
const body = [];

// ── 封面 ────────────────────────────────────────────────────
body.push(
  sp(1440),
  par([hrun('○○高中', 52, C_DEEP, true)], { alignment: AlignmentType.CENTER }),
  par([hrun('資訊處理事項內部控制制度', 52, C_DEEP, true)], { alignment: AlignmentType.CENTER }),
  sp(200),
  par([hrun('原始版本 vs. 修訂版本　對照文件', 28, C_RED, true)], { alignment: AlignmentType.CENTER }),
  sp(300),
  par([run('主管單位：教務處設備組', { size: 24 })], { alignment: AlignmentType.CENTER }),
  PAGE_BREAK()
);

// ── 圖例 ────────────────────────────────────────────────────
body.push(
  par([hrun('【圖例說明】', 24, C_MID, true)], { spacing: { before: 120, after: 80 } }),
  par([rDel('灰色刪除線'), run('　→　原始版本有此內容，修訂版本已刪除')], { spacing: { after: 60 } }),
  par([rNew('紅色粗體'), run('　→　修訂版本新增或修改之內容')], { spacing: { after: 120 } }),
  sp(60)
);

// ── 總說明 ───────────────────────────────────────────────────
body.push(
  H1('（五）資訊處理事項'),
  par([run('本項目依據設備組業務範疇，分為下列'), rDel('七'), run('  '), rNew('四'), run('大作業：')]),
  par([rDel('● 系統開發及程式修改作業')], { spacing: { after: 40 } }),
  par([rDel('● 程式及資料之存取作業')], { spacing: { after: 40 } }),
  par([rDel('● 資料輸出入及處理作業')], { spacing: { after: 40 } }),
  par([run('● 檔案及設備之安全作業')], { spacing: { after: 40 } }),
  par([run('● 硬體及系統軟體之使用及維護作業')], { spacing: { after: 40 } }),
  par([run('● 系統復原計畫及測試作業')], { spacing: { after: 40 } }),
  par([run('● 資訊安全之檢查作業')], { spacing: { after: 80 } }),
  sp(40),

  H2('【新增】本校資訊系統環境說明', 'new'),
  par([rNew('原版未說明系統管轄範疇，修訂版新增下列對照表，明確界定設備組職責邊界：')], { spacing: { after: 60 } }),
  mkTbl(
    ['類別', '說明', '設備組角色'],
    [
      ['校園網路及校網系統', '學校官網（校網）及本校自建之網路基礎設施，由設備組負責建置、維護與備份。', '主要負責單位'],
      ['各類校務行政系統', '成績、學籍、人事、財務等校務系統，均已向上集中至國教署，由國教署統一建置、維護及備份；本校各處室為系統使用端。', '使用端，非管理端'],
    ],
    [2200, 5000, 2200],
    ['new', 'new']
  ),
  sp(80)
);

// ════════════════════════════════════════════════════════════
//  【已刪除】一、系統開發及程式修改作業
// ════════════════════════════════════════════════════════════
body.push(
  PAGE_BREAK(),
  H1('◎ 系統開發及程式修改作業', 'del'),
  par([rDel('【已刪除】本校校務系統由教務處採購委外（天方科技），非設備組開發維護；帳號管理由註冊組負責，故從設備組內控中刪除。')],
    { spacing: { after: 100 } }),

  H2('1. 流程圖（原版）', 'del'),
  flowTbl(['需求提出\n申請單', '→ 設備組\n可行性評估', '→ 系統分析\n設計', '→ 程式開發\n測試', '→ 使用單位\n驗收', '→ 正式啟用'], 'del'),
  sp(80),

  H2('2. 作業程序（原版）', 'del'),
  numItem('2.1', '應用系統開發及管理制度：訂定開發制度，配合學校整體發展策略。', 'del'),
  numItem('2.2', '系統計劃、開發及管理：使用單位填具「電腦作業需求申請單」，經主管核准後送交設備組進行自行開發或委外可行性評估；含系統分析、設計等步驟。', 'del'),
  numItem('2.3', '系統發展及測試管理：設備組專責人員撰寫程式碼、進行整體性測試；測試結果經教務主任及使用單位認可後結案。', 'del'),
  numItem('2.4', '系統運作使用管理：設備人員負責教育訓練；程式驗收後於申請單登載驗收日期。', 'del'),
  numItem('2.5', '系統評核管理：啟用後持續回饋異常至設備組；設備組執行修正並評估績效。', 'del'),
  numItem('2.6', '系統維護與管理：修改需求由需求單位填申請單，經各主管核准後設備組辦理；修改後通知需求單位測試，測試通過後歸檔。', 'del'),
  numItem('2.7', '外包業務管理：外包應經既定程序核准，責任範圍明訂於契約；保固期滿後視需要訂維護合約。', 'del'),

  H2('4. 使用表單（原版）', 'del'),
  par([rDel('● 電腦作業需求申請單')], { spacing: { after: 40 } }),
  sp(40)
);

// ════════════════════════════════════════════════════════════
//  【已刪除】二、程式及資料之存取作業
// ════════════════════════════════════════════════════════════
body.push(
  PAGE_BREAK(),
  H1('◎ 程式及資料之存取作業', 'del'),
  par([rDel('【已刪除】本校無獨立程式開發；校務系統程式及資料存取屬天方科技系統本身機制，不在設備組業務範疇。')],
    { spacing: { after: 100 } }),

  H2('2. 作業程序（原版）', 'del'),
  numItem('2.1', '資訊安全規劃：設備組評估資訊安全問題，訂定偵防措施及管理辦法，訂定安全目標及政策（含法令遵循、業務永續）。', 'del'),
  numItem('2.2', '使用者權限管理：新進/異動人員提出「系統使用授權申請表」，主管核可後設定權限；離職時應同步刪除帳號。', 'del'),
  numItem('2.3', '資料變更管理：資料由各使用單位擁有；他單位需調取應經協調並設定權限；原始資料與電腦輸出可互相勾稽。', 'del'),
  numItem('2.4', '程式及檔案存取：申請→主管簽核→設備組辦理；存取均留紀錄；原始程式碼與執行檔分開存放。', 'del'),
  numItem('2.5', '程式管理：重要程式非經授權不得使用；修改均應核准並留紀錄；修改後更新說明文件。', 'del'),

  H2('4. 使用表單（原版）', 'del'),
  par([rDel('● 系統使用授權申請表')], { spacing: { after: 40 } }),
  sp(40)
);

// ════════════════════════════════════════════════════════════
//  【已刪除】三、資料輸出入及處理作業
// ════════════════════════════════════════════════════════════
body.push(
  PAGE_BREAK(),
  H1('◎ 資料輸出入及處理作業', 'del'),
  par([rDel('【已刪除】資料輸出入由各處室依業務職掌負責（教務/學務/總務），人事室以人工作業辦理，非設備組業務。')],
    { spacing: { after: 100 } }),

  H2('2. 作業程序（原版）', 'del'),
  numItem('2.1', '輸入：依原始單據執行輸入，輸入前審核是否經主管簽核；應用程式設定自動檢核；錯誤更正由授權專人負責。', 'del'),
  numItem('2.2', '輸出：機密性資料輸出有適當管制；使用後無保存需要應予銷毀；重要資料查詢須經核准。', 'del'),
  sp(40)
);

// ════════════════════════════════════════════════════════════
//  【保留修訂】一、檔案及設備之安全作業
// ════════════════════════════════════════════════════════════
body.push(
  PAGE_BREAK(),
  H1('◎ 檔案及設備之安全作業'),

  H2('1. 流程圖'),
  par([run('原版：無流程圖　→　'), rNew('新版新增備份流程（加入 3-2-1 原則及每月還原測試）')], { spacing: { after: 60 } }),
  flowTbl(['排程自動\n每日備份', '→ 備份完成\n記錄結果', '→ 異常？\n→ 立即通報', '→ 每週完整備份\n存放異地/雲端', '→ 每月\n還原測試', '→ 測試報告\n送主管核閱']),
  sp(120),

  H2('2. 作業程序'),
  H3('2.1 機房管理'),
  numItem('2.1.1', '機房環境標準：溫度 18–25°C、濕度 40–60%；配置不斷電系統（UPS）、穩壓設備及自動滅火系統；每月填寫「主機房工作記錄表」。'),
  numItem('2.1.2', '機房進出採門禁管制，訪客進入須登記姓名、事由及陪同人員；機房設置監視攝影設備。'),
  numItem('2.1.3', '機房內禁止存放易燃物品；二氧化碳滅火器每年定期檢驗並記錄；逃生路線保持暢通並設有緊急照明。'),
  numItem('2.1.4', '設備組每月定期巡查機房，填寫巡查紀錄，異常情形立即通報教務主任。'),

  H3('2.2 檔案備份'),
  par([rDel('2.2.1 依「檔案備份計畫」執行定期備份：重要業務系統每日差異備份，每週完整備份一次。')], { indent: { left: 360 }, spacing: { after: 40 } }),
  numItem('2.2.1', '【修改】依「檔案備份計畫」執行定期備份，備份對象為本校自管系統資料（校網、網路設備組態、設備組業務文件）；每日差異備份，每週完整備份一次。國教署集中管理之校務系統（成績、學籍等）由國教署負責備份，不在本節範疇。', 'new'),
  numItem('2.2.2', '【新增】備份遵循 3-2-1 原則：3 份備份、2 種儲存媒介（本機磁碟 + NAS）、1 份存放異地或雲端。', 'new'),
  numItem('2.2.3', '【新增】每月至少執行一次備份還原測試；測試結果記錄於「備份還原測試紀錄表」並送教務主任核閱。', 'new'),
  numItem('2.2.4', '備份媒體定期盤點；損壞媒體安全銷毀並留有紀錄；備份資料存放位置應有明顯標示。'),

  H3('2.3 安全管理'),
  numItem('2.3.1', '電腦使用區域之家具及地板宜採不易燃材質；設置二氧化碳滅火器。'),
  numItem('2.3.2', '對進出電腦使用區域之敏感地區，應有足夠管制措施，非相關人員不得進入。'),
  numItem('2.3.3', '設備送外維修時，應由設備組人員全程陪同，並確認敏感資料已清除。'),

  H3('2.4 可攜式媒體管理', 'new'),
  par([rNew('【新增子節】原版無此節')], { spacing: { after: 60 } }),
  numItem('2.4.1', '【新增】可攜式媒體應登記列管，標示擁有者及用途，不使用時存放於安全位置。', 'new'),
  numItem('2.4.2', '【新增】廢棄含個資媒體（含硬碟）應實體銷毀，並留有銷毀紀錄。', 'new'),
  numItem('2.4.3', '【新增】禁止使用未經設備組核可之可攜式媒體連接校內系統，防範惡意程式感染。', 'new'),

  H3('2.5 異地備援', 'new'),
  par([rDel('原版備註：「因經費與管理的緣故尚未施行」'), run('　→　'), rNew('【修改】已納入正式作業')], { spacing: { after: 60 } }),
  numItem('2.5.1', '【修改】每週完整備份資料應存放至異地或雲端，確保災害發生時可完整回復。', 'new'),
  numItem('2.5.2', '【新增】每學期進行一次異地備份還原演練，確保回復程序能在規定時間內完成。', 'new'),
  sp(80),

  H2('3. 控制重點'),
  mkTbl(
    ['項次', '控制重點', '查核方式'],
    [
      ['1', '機房進出是否有門禁管制且留有紀錄，非相關人員是否禁止進入。', '查閱門禁紀錄或訪客登記本'],
      ['2', '備份是否依「檔案備份計畫」確實執行，備份紀錄是否完整。', '查閱備份作業日誌'],
      ['3', '【新增】備份是否符合 3-2-1 原則，含異地或雲端備份。', '確認備份儲存位置'],
      ['4', '【新增】每月是否執行備份還原測試，測試報告是否送主管核閱。', '查閱還原測試紀錄'],
      ['5', '廢棄含個資媒體是否確實安全銷毀並留有銷毀紀錄。', '抽查銷毀紀錄'],
      ['6', '【新增】機房滅火器是否每年定期檢驗並記錄。', '查閱設備檢驗紀錄'],
    ],
    [400, 4800, 3200],
    ['same', 'same', 'new', 'new', 'same', 'new']
  ),
  sp(80),

  H2('4. 使用表單'),
  bpar(run('主機房工作記錄表')),
  bpar(run('檔案備份計畫')),
  bpar(rNew('【新增】備份還原測試紀錄表')),
  bpar(rNew('【新增】媒體銷毀紀錄表')),

  H2('5. 依據及相關文件'),
  bpar(run('本校機房管理辦法')),
  bpar(rNew('【新增】個人資料保護法')),
  sp(40)
);

// ════════════════════════════════════════════════════════════
//  【保留修訂】二、硬體及系統軟體之使用及維護作業
// ════════════════════════════════════════════════════════════
body.push(
  PAGE_BREAK(),
  H1('◎ 硬體及系統軟體之使用及維護作業'),

  H2('1. 流程圖'),
  par([run('原版：無流程圖　→　'), rNew('新版新增採購流程與維護流程')], { spacing: { after: 60 } }),
  flowTbl(['使用單位\n提出需求', '→ 設備組\n訂定規格', '→ 教務主任\n核准', '→ 依採購法\n辦理採購', '→ 驗收測試', '→ 納入資產\n清冊管理']),
  sp(80),
  flowTbl(['設備故障\n通報', '→ 設備組\n初步診斷', '→ 可自修？\n自行修復', '→ 否→洽廠商\n維修', '→ 修復驗收\n更新紀錄']),
  sp(120),

  H2('2. 作業程序'),
  H3('2.1 硬體設施管理'),
  numItem('2.1.1', '設備組建立各設備維護紀錄，記錄故障原因、排除方法及修復時間；主管定期核閱。'),
  numItem('2.1.2', '各設備依廠商建議排定預防性維護（PM）計畫，每學期至少執行一次，並作成紀錄。'),
  numItem('2.1.3', '設備送外維修時，記錄送修日期、廠商及預計返回日期，並確認敏感資料已清除。'),
  numItem('2.1.4', '各項電腦設備資產標籤應完整，資產清冊每學期更新一次。'),
  numItem('2.1.5', '【新增】設備故障達一定規模（影響 10 台以上或核心系統）時，應立即通報教務主任並啟動應急措施。', 'new'),

  H3('2.2 可攜式媒體管理'),
  numItem('2.2.1', '磁帶、光碟、隨身碟、記憶卡等媒體均應登記列管，標籤註明清楚，放置於安全位置。'),
  numItem('2.2.2', '廢棄數位媒體應實體銷毀或安全清除，不得再行使用，並留有銷毀紀錄。'),

  H3('2.3 智慧財產權管理'),
  numItem('2.3.1', '本校所有電腦應使用具合法授權之軟體；禁止安裝盜版或未授權軟體。'),
  numItem('2.3.2', '設備組每學期對行政電腦及電腦教室進行軟體稽查，確認授權狀態，軟體授權憑證妥善保存。'),
  numItem('2.3.3', '【修改】教師及學生不得自行安裝未經設備組核可之軟體；需安裝特定教學軟體，應填寫申請單由設備組統一處理。', 'new'),

  H3('2.4 軟硬體採購管理'),
  numItem('2.4.1', '採購依年度預算編列，並依採購法辦理（未達查核金額採比價，達查核金額採公開招標）。'),
  numItem('2.4.2', '採購前由設備組提出需求規格，與使用單位共同評估效益，送教務主任核准後辦理。'),
  numItem('2.4.3', '軟體採購前應確認授權數量足夠，並測試相容性；安裝前確認已知問題已解決。'),
  numItem('2.4.4', '採購完成後進行驗收測試，確認符合規格後納入資產管理清冊，並保存授權憑證。'),
  numItem('2.4.5', '未經授權，設備、資料或軟體不得攜離學校。'),

  H3('2.5 學生行動裝置管理', 'new'),
  par([rNew('【新增子節】原版無此節')], { spacing: { after: 60 } }),
  numItem('2.5.1', '【新增】學生平板電腦應部署行動裝置管理（MDM）系統，限制可安裝之應用程式範圍，並設定使用政策。', 'new'),
  numItem('2.5.2', '【新增】設備組每學期更新 MDM 政策；設備損壞由使用者填具損壞申報單，依學校賠償規定辦理。', 'new'),
  sp(80),

  H2('3. 控制重點'),
  mkTbl(
    ['項次', '控制重點', '查核方式'],
    [
      ['1', '設備故障是否完整記錄，主管是否定期核閱維護紀錄。', '查閱維護紀錄表'],
      ['2', '是否每學期執行預防性維護並留有紀錄。', '查閱 PM 計畫及紀錄'],
      ['3', '是否每學期稽查全校電腦軟體授權，禁止盜版軟體。', '查看軟體稽查報告'],
      ['4', '採購是否依採購法辦理？驗收紀錄是否完整？授權憑證是否保存？', '審閱採購文件及驗收紀錄'],
      ['5', '資產清冊是否每學期更新，設備資產標籤是否完整。', '現場盤點抽查'],
      ['6', '【新增】學生平板是否部署 MDM，政策是否定期更新。', '查看 MDM 管理主控台'],
    ],
    [400, 4800, 3200],
    ['same', 'same', 'same', 'same', 'same', 'new']
  ),
  sp(80),

  H2('4. 使用表單'),
  bpar(run('硬體設備維護紀錄表')),
  bpar(run('軟體授權稽查表（每學期）')),
  bpar(run('設備資產清冊')),
  bpar(run('軟硬體採購需求申請單')),
  bpar(rNew('【新增】設備損壞申報單（學生）')),

  H2('5. 依據及相關文件'),
  bpar(run('政府採購法及相關子法')),
  bpar(run('著作權法（軟體授權）')),
  bpar(run('本校設備採購及管理規定')),
  sp(40)
);

// ════════════════════════════════════════════════════════════
//  【保留修訂】三、系統復原計畫及測試作業
// ════════════════════════════════════════════════════════════
body.push(
  PAGE_BREAK(),
  H1('◎ 系統復原計畫及測試作業'),

  H2('1. 流程圖'),
  par([run('原版：無流程圖　→　'), rNew('新版新增故障應變及復原測試流程')], { spacing: { after: 60 } }),
  flowTbl(['系統發生故障', '→ 使用者\n填維修申請單', '→ 設備組\n30分鐘內回應', '→ 可自修\n→ 修復', '→ 否→洽廠商\n或啟動備援', '→ 修復後\n資料驗證', '→ 紀錄歸檔']),
  sp(120),

  H2('2. 作業程序'),
  H3('2.1 備援計畫訂定'),
  par([rDel('2.1.1 原版：應有備援計劃以處理關鍵性工作（未明訂 RTO/RPO 目標，未區分自管與委外系統）。')], { indent: { left: 360 }, spacing: { after: 40 } }),
  numItem('2.1.1', '【修改】設備組應訂定「系統復原計畫（DRP）」，識別本校自管關鍵系統（校網、網路核心設備、NAS 備份系統）並設定 RTO ≤ 4 小時、RPO ≤ 24 小時。國教署所管校務系統之 DRP 由國教署負責；本校負責確保使用端網路連線在 RTO 內恢復。', 'new'),
  numItem('2.1.2', '備援計畫應包含：緊急聯絡名單（設備組、廠商、教務主任）、復原優先順序、備援設備清單。'),
  numItem('2.1.3', '【修改】備援計畫每學年至少更新一次；重大環境變更後應立即修訂。', 'new'),
  par([rDel('原版：需定期測試備援計劃（未明訂每學年更新頻率）。')], { indent: { left: 360 }, spacing: { after: 60 } }),

  H3('2.2 故障復原'),
  numItem('2.2.1', '系統故障時，使用者填具「維修申請單」通報設備組；設備組應於 30 分鐘內確認問題，並於 4 小時內完成修復或啟動備援。'),
  numItem('2.2.2', '故障處理：由不同單位人員組成緊急應變小組；復原工作依備援計畫規定之優先順序執行。'),
  numItem('2.2.3', '硬體無法自行修復時，洽維護廠商；備份媒體由設備組執行還原；若設備損壞無法修復，立即採購相容設備。'),
  numItem('2.2.4', '【新增】發生資安事件（勒索軟體、入侵）時，立即隔離受感染系統，通報教務主任；重大資安事件依規定通報教育部（1 小時內初報）。', 'new'),
  numItem('2.2.5', '故障復原後，追查故障原因、研討解決方案，並更新備援計畫防止再次發生。'),

  H3('2.3 復原結果測試'),
  numItem('2.3.1', '【修改】每學期至少執行一次系統復原演練，測試備援計畫之可行性及備份資料完整性。', 'new'),
  par([rDel('原版：重置後確認資料回存完整性（無明訂每學期演練頻率）。')], { indent: { left: 360 }, spacing: { after: 60 } }),
  numItem('2.3.2', '測試完成後，暫存資料確認完整回存後，安全銷毀暫存資料。'),
  numItem('2.3.3', '設備組撰寫「系統復原測試報告」，包含測試範圍、發現問題及改善措施，送教務主任核閱後建檔。'),
  sp(80),

  H2('3. 控制重點'),
  mkTbl(
    ['項次', '控制重點', '查核方式'],
    [
      ['1', '【修改】是否制定書面系統復原計畫（DRP），並設定 RTO/RPO 目標。', '審閱 DRP 文件'],
      ['2', '【修改】備援計畫是否每學年更新，緊急聯絡名單是否正確。', '核對計畫更新日期'],
      ['3', '故障時設備組是否於 30 分鐘內回應、4 小時內修復或啟動備援。', '查閱維修申請單及處理紀錄'],
      ['4', '是否每學期執行復原演練，並留有測試報告送主管核閱。', '查閱復原測試報告'],
      ['5', '【新增】重大資安事件是否依規定時限通報教育部。', '查閱事件通報紀錄'],
      ['6', '復原後是否追查故障原因並更新備援計畫。', '審閱事後改善措施'],
    ],
    [400, 4800, 3200],
    ['new', 'new', 'same', 'same', 'new', 'same']
  ),
  sp(80),

  H2('4. 使用表單'),
  bpar(run('維修申請單')),
  bpar(rNew('【修改】系統復原計畫（DRP）（新增 RTO/RPO 目標）')),
  bpar(run('系統復原測試報告（每學期）')),

  H2('5. 依據及相關文件'),
  bpar(run('資通安全管理法')),
  bpar(rNew('【新增】教育部資安事件通報規定')),
  bpar(run('本校緊急應變計畫')),
  sp(40)
);

// ════════════════════════════════════════════════════════════
//  【保留修訂】四、資訊安全之檢查作業
// ════════════════════════════════════════════════════════════
body.push(
  PAGE_BREAK(),
  H1('◎ 資訊安全之檢查作業'),
  par([rNew('【架構修改】原版為 7 條平列項目，新版重新架構為 4 個子節（網路安全、端點安全、資安教育訓練【新增】、資安事件通報【新增】）')],
    { spacing: { after: 80 } }),

  H2('1. 流程圖'),
  par([run('原版：無流程圖　→　'), rNew('新版新增定期資安檢查流程（含弱點掃描分級修補）')], { spacing: { after: 60 } }),
  flowTbl(['擬定資安\n檢查計劃', '→ 執行定期\n弱點掃描', '→ 發現弱點\n→ 分級評估', '→ 高風險72hr\n內修補', '→ 填寫檢查\n紀錄', '→ 送主管\n核閱建檔'], 'new'),
  sp(120),

  H2('2. 作業程序'),
  H3('2.1 網路安全管理'),
  par([rDel('原版 2.1：設備組負責網路安全規範擬訂，確保系統與資料的安全性。')], { spacing: { after: 40 } }),
  numItem('2.1.1', '設備組負責網路安全規範擬訂，部署防火牆及入侵偵測/防禦系統（IDS/IPS），並定期審視防火牆規則（每學期至少一次）。'),
  numItem('2.1.2', '【新增】校園網路應劃分為三個獨立區域：行政網段、教學網段及訪客 Wi-Fi；各網段間存取應受管控。', 'new'),
  numItem('2.1.3', '【新增】本校對外開放之學校官網（校網）應每半年進行弱點掃描，高風險漏洞應於 72 小時內修補，中低風險應於 30 天內修補。國教署所管校務系統弱點掃描由國教署辦理，本校負責配合通報。', 'new'),
  numItem('2.1.4', '員工非經主管授權，禁止將本校相關資料對外傳送；郵件伺服器應部署防垃圾郵件及防毒功能。'),

  H3('2.2 端點安全管理'),
  par([rDel('原版 2.2：應具備病毒掃瞄軟體，定期掃瞄並更新病毒碼。')], { spacing: { after: 40 } }),
  numItem('2.2.1', '全校電腦應安裝防毒軟體，病毒碼每日自動更新；作業系統資安修補應於公告後 30 天內套用。'),
  numItem('2.2.2', '禁止師生使用 P2P 軟體或瀏覽非法網站；學生上網行為應透過網路閘道管控與記錄。'),
  numItem('2.2.3', '【新增】電腦教室設定使用者不得自行安裝軟體，學生操作環境採用磁碟還原系統（下課後自動還原）。', 'new'),
  numItem('2.2.4', '重要軟體及檔案應加密處理，並定期更新密碼；機密資料傳輸使用加密通道（HTTPS/TLS）。'),
  numItem('2.2.5', '定期備份重要檔案及資料，防止資料遺失。'),

  H3('2.3 資訊安全教育訓練', 'new'),
  par([rNew('【新增子節】原版無此節')], { spacing: { after: 60 } }),
  numItem('2.3.1', '【新增】每學期對教職員工辦理一次資安宣導（釣魚郵件辨識、密碼安全、個資保護、社交工程防範）。', 'new'),
  numItem('2.3.2', '【新增】新進教職員工應於到職一個月內完成資安基礎教育訓練，並簽署「資訊安全使用規範確認書」。', 'new'),
  numItem('2.3.3', '【新增】每學年開學時對學生進行資訊倫理與網路安全宣導，納入資訊課程，並留有宣導紀錄。', 'new'),

  H3('2.4 資安事件通報與處理', 'new'),
  par([rNew('【新增子節】原版無正式通報程序')], { spacing: { after: 60 } }),
  numItem('2.4.1', '【新增】發現資安事件應立即通報設備組長及教務主任。', 'new'),
  numItem('2.4.2', '【新增】重大資安事件依《資通安全管理法》及教育部規定，於 1 小時內向教育局/教育部通報。', 'new'),
  numItem('2.4.3', '【新增】資安事件應留有完整紀錄，包含：事件描述、影響範圍、處置措施及後續改善計畫。', 'new'),
  sp(80),

  H2('3. 控制重點'),
  mkTbl(
    ['項次', '控制重點', '查核方式'],
    [
      ['1', '是否建立電腦網路系統安全控管機制（防火牆、IDS/IPS）。', '查看防火牆設定及 IDS 日誌'],
      ['2', '【新增】網路是否依功能劃分行政、教學及訪客三個獨立網段。', '查看網路架構圖'],
      ['3', '【新增】對外服務是否每半年執行弱點掃描，高風險漏洞是否於 72 小時內修補。', '查閱弱點掃描報告'],
      ['4', '全校電腦防毒軟體是否安裝且病毒碼每日更新。', '查看防毒管理主控台'],
      ['5', '設備組是否每學期對郵件收發異常進行監控，是否陳報主管處理。', '查閱郵件管理紀錄'],
      ['6', '【新增】教職員是否每學期完成資安宣導，新進人員是否簽署使用規範確認書。', '查閱宣導紀錄及確認書'],
      ['7', '機密檔案是否加密儲存，傳輸是否使用加密通道。', '抽查加密設定'],
    ],
    [400, 4800, 3200],
    ['same', 'new', 'new', 'same', 'same', 'new', 'same']
  ),
  sp(80),

  H2('4. 使用表單'),
  par([rDel('原版：無使用表單')], { spacing: { after: 60 } }),
  bpar(rNew('【新增】資安弱點掃描報告（每半年）')),
  bpar(rNew('【新增】資安事件通報紀錄表')),
  bpar(rNew('【新增】資訊安全教育訓練簽到表')),
  bpar(rNew('【新增】資訊安全使用規範確認書（新進人員）')),

  H2('5. 依據及相關文件'),
  par([rDel('原版：無相關文件')], { spacing: { after: 60 } }),
  bpar(rNew('【新增】資通安全管理法')),
  bpar(rNew('【新增】個人資料保護法')),
  bpar(rNew('【新增】教育部校園資安事件通報規定')),
  bpar(rNew('【新增】本校資訊安全管理辦法')),
  sp(40)
);

// ── 建立文件 ─────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: 'b0', levels: [{ level: 0, format: LevelFormat.BULLET, text: '●', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 300 } } } }] },
      { reference: 'b1', levels: [{ level: 0, format: LevelFormat.BULLET, text: '◆', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }
    ]
  },
  styles: {
    default: { document: { run: { font: { name: FONT }, size: 22, color: '1A1A1A' } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 30, bold: true, font: { name: FONT }, color: C_MID }, paragraph: { spacing: { before: 400, after: 120 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 24, bold: true, font: { name: FONT }, color: C_LITE }, paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [par(
          [
            run('○○高中　資訊處理事項內部控制制度　對照版', { size: 18, color: '888888' }),
            new TextRun({ children: ['\t'], font: { name: FONT } }),
            new TextRun({ children: ['第 ', PageNumber.CURRENT, ' 頁'], font: { name: FONT }, size: 18, color: '888888' }),
          ],
          {
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C_BORD, space: 4 } }
          }
        )]
      })
    },
    children: body
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('D:\\D\\114設備組\\高中資訊內控制度_對照版v2.docx', buf);
  console.log('DONE');
}).catch(e => { console.error(e); process.exit(1); });
