const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, TabStopType,
  TabStopPosition
} = require('C:/Users/user/AppData/Roaming/npm/node_modules/docx');
const fs = require('fs');

// ─── 顏色與樣式設定 ────────────────────────────────────────
const COLOR_TITLE   = '1F3864'; // 深藍
const COLOR_H1      = '2E5F8A'; // 中藍
const COLOR_H2      = '2E75B6'; // 淺藍
const COLOR_ACCENT  = 'C00000'; // 深紅（控制重點標題）
const COLOR_BORDER  = '9DC3E6';
const FONT          = '標楷體';
const FONT_EN       = 'Times New Roman';

function p(children, opts = {}) {
  return new Paragraph({ children, ...opts });
}
function t(text, opts = {}) {
  return new TextRun({ text, font: { name: FONT }, ...opts });
}
function bold(text, color) {
  return new TextRun({ text, bold: true, font: { name: FONT }, color: color || '000000' });
}

const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER };
const allBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

function headerRow(texts, widths) {
  return new TableRow({
    tableHeader: true,
    children: texts.map((txt, i) =>
      new TableCell({
        borders: allBorders,
        width: { size: widths[i], type: WidthType.DXA },
        shading: { fill: '2E5F8A', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [p([new TextRun({ text: txt, bold: true, color: 'FFFFFF', font: { name: FONT } })],
          { alignment: AlignmentType.CENTER })]
      })
    )
  });
}
function dataRow(cells, widths, shadeEven = false) {
  return new TableRow({
    children: cells.map((txt, i) =>
      new TableCell({
        borders: allBorders,
        width: { size: widths[i], type: WidthType.DXA },
        shading: shadeEven ? { fill: 'EBF3FB', type: ShadingType.CLEAR } : undefined,
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [p([t(txt)], { alignment: i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT })]
      })
    )
  });
}

// ─── 章節工廠函式 ──────────────────────────────────────────
function sectionTitle(num, title) {
  return p([
    new TextRun({ text: `${num}、${title}`, bold: true, size: 32, color: COLOR_H1, font: { name: FONT } })
  ], {
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLOR_H1, space: 4 } }
  });
}
function subTitle(num, title, colorOverride) {
  return p([
    new TextRun({ text: `${num} ${title}`, bold: true, size: 26, color: colorOverride || COLOR_H2, font: { name: FONT } })
  ], { heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 80 } });
}
function item(text) {
  return p([t(text)], {
    numbering: { reference: 'bullets', level: 0 },
    spacing: { after: 60 }
  });
}
function item2(text) {
  return p([t(text)], {
    numbering: { reference: 'bullets2', level: 0 },
    spacing: { after: 60 }
  });
}
function space(before = 120) {
  return p([], { spacing: { before } });
}

// ─── 正文內容 ───────────────────────────────────────────────
const content = [];

// 封面標題
content.push(
  space(720),
  p([new TextRun({ text: '○○高中', bold: true, size: 52, color: COLOR_TITLE, font: { name: FONT } })],
    { alignment: AlignmentType.CENTER, spacing: { after: 160 } }),
  p([new TextRun({ text: '資訊內部控制作業制度', bold: true, size: 52, color: COLOR_TITLE, font: { name: FONT } })],
    { alignment: AlignmentType.CENTER, spacing: { after: 160 } }),
  p([new TextRun({ text: '（Information Internal Control System）', size: 28, color: '555555', font: { name: FONT_EN } })],
    { alignment: AlignmentType.CENTER, spacing: { after: 400 } }),
  p([t(`訂定日期：中華民國　　年　　月　　日`)], { alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
  p([t('版　　本：第 1.0 版')], { alignment: AlignmentType.CENTER }),
  new Paragraph({ children: [], pageBreakBefore: true })
);

// ━━━ 壹、總則 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
content.push(
  sectionTitle('壹', '總則'),
  subTitle('一、', '目的'),
  item('為強化本校資訊系統之管理與控制，確保資訊資產之安全性、完整性及可用性，並遵循相關法規與政策，特訂定本制度。'),

  subTitle('二、', '適用範圍'),
  item('本制度適用於本校所有行政單位、教師及學生，凡使用本校資訊系統、設備及網路資源者，均應遵守本制度。'),

  subTitle('三、', '主管單位與權責'),
  new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [2500, 3000, 3500],
    rows: [
      headerRow(['權責單位', '職稱', '主要職責'], [2500, 3000, 3500]),
      dataRow(['主管', '教務主任', '核准重大資訊政策與採購'], [2500, 3000, 3500]),
      dataRow(['執行', '設備組長', '統籌資訊系統與設備管理'], [2500, 3000, 3500], true),
      dataRow(['協辦', '設備組組員', '執行日常維運、帳號與備份管理'], [2500, 3000, 3500]),
      dataRow(['配合', '各處室主任', '確認本單位使用需求及帳號申請'], [2500, 3000, 3500], true),
    ]
  }),
  space(),

  subTitle('四、', '名詞定義'),
  item('資訊資產：包含硬體設備、系統軟體、應用程式、資料庫及紀錄文件。'),
  item('使用者：凡取得帳號並使用本校資訊系統之教職員工生。'),
  item('設備組：負責本校資訊相關業務之單位。'),
  space(),
  new Paragraph({ children: [], pageBreakBefore: true })
);

// ━━━ 貳、系統開發及程式修改作業 ━━━━━━━━━━━━━━━━━━━━━━━━
content.push(
  sectionTitle('貳', '系統開發及程式修改作業'),
  subTitle('一、', '目的'),
  p([t('確保本校自行開發或委外開發之資訊系統，依循既定程序進行需求分析、設計、測試及上線，並管控程式修改以避免未授權變更。')],
    { spacing: { after: 120 } }),

  subTitle('二、', '作業程序'),

  p([bold('（一）需求提出', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('由使用單位填寫「電腦作業需求申請單」，載明功能需求、預期效益及優先順序，經該單位主任簽核後送設備組。'),
  item('設備組評估自行開發或委外開發可行性，並於 5 個工作天內回覆評估結果。'),
  item('委外開發應依政府採購法規辦理，並於合約載明功能規格、保固期限及原始碼授權條款。'),

  p([bold('（二）系統分析與設計', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('設備組指派專責人員進行系統分析，依需求訂定輸入介面、資料結構、輸出報表及流程設計。'),
  item('系統分析人員應與使用單位充分討論，並製作系統設計規格書，送教務主任核准後始進行開發。'),

  p([bold('（三）程式開發與測試', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('程式開發應於獨立測試環境進行，不得直接於正式環境修改。'),
  item('完成開發後依序執行單元測試、整合測試，並由使用單位進行使用者驗收測試（UAT）。'),
  item('測試過程及發現之錯誤應完整記錄，修正後重新測試，所有記錄應存檔備查。'),
  item('正式上線前須取得教務主任書面核准，並由使用單位於「電腦作業需求申請單」簽名確認。'),

  p([bold('（四）系統維護與修改', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('程式修改需求應由使用單位提出修改申請，經設備組評估影響範圍及必要性後，報教務主任核准始得進行。'),
  item('修改完成後須通知使用單位測試，並將測試結果回饋設備組；確認無誤後更新操作文件。'),
  item('修改過程、前後版本差異及測試記錄應完整保存，原始程式碼應採版本控制管理（如 Git）。'),

  p([bold('（五）外包管理', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('委外廠商人員進入本校時，應登記訪客紀錄，並由設備組人員全程陪同。'),
  item('委外合約應明訂保固期限、資料保密義務（NDA）及系統原始碼之歸屬。'),
  item('保固期滿後，視需要另訂維護合約，並由教務主任核准。'),

  subTitle('三、', '控制重點', COLOR_ACCENT),
  new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [600, 4200, 4200],
    rows: [
      headerRow(['項次', '控制項目', '稽核查核重點'], [600, 4200, 4200]),
      dataRow(['1', '需求管理', '申請單是否完整填寫並經主管簽核？'], [600, 4200, 4200]),
      dataRow(['2', '測試環境', '開發是否與正式環境分離？'], [600, 4200, 4200], true),
      dataRow(['3', '使用者驗收', 'UAT 是否由使用單位執行並留下紀錄？'], [600, 4200, 4200]),
      dataRow(['4', '版本控制', '程式修改是否有版本紀錄可供追溯？'], [600, 4200, 4200], true),
      dataRow(['5', '外包管理', '委外合約是否載明保密義務及原始碼歸屬？'], [600, 4200, 4200]),
    ]
  }),
  space(),

  subTitle('四、', '使用表單'),
  item('電腦作業需求申請單。'),
  space(),
  new Paragraph({ children: [], pageBreakBefore: true })
);

// ━━━ 參、程式及資料存取管理 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
content.push(
  sectionTitle('參', '程式及資料存取管理'),
  subTitle('一、', '目的'),
  p([t('建立使用者身分識別與存取控制機制，防止未經授權之人員存取本校資訊系統及資料，保護師生個人資料之安全。')],
    { spacing: { after: 120 } }),

  subTitle('二、', '作業程序'),

  p([bold('（一）帳號申請與授權', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('新進教職員工帳號由人事室通知設備組，依「系統使用授權申請表」設定初始帳號及最低必要權限。'),
  item('學生帳號依學籍資料批次建立；轉入學生由教務處通知設備組新增，轉出或退學者應於三個工作天內停用帳號。'),
  item('各系統功能模組之存取權限，依使用者工作職責設定，遵循「最小授權原則」。'),
  item('帳號設定完成後，由設備組通知使用者，並要求首次登入時更改預設密碼。'),

  p([bold('（二）密碼管理', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('密碼長度至少 8 碼，須包含大小寫英文字母、數字或特殊符號之組合。'),
  item('教職員工帳號密碼應每學期更新一次；連續 5 次輸入錯誤應自動鎖定帳號，需由設備組解鎖。'),
  item('嚴禁共用帳號或將密碼告知他人；各系統不得儲存明文密碼。'),

  p([bold('（三）離職/異動處理', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('人員離職或調職時，人事室應於最後上班日前通知設備組，停用或調整其帳號權限。'),
  item('離職人員帳號應於離職當日鎖定，相關資料依保存政策轉交業務接辦人員。'),

  p([bold('（四）資料變更管理', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('各系統資料由各主辦單位負責，其他單位需調取資料，應經雙方主任同意並填具申請單。'),
  item('系統資料更動應留有完整稽核軌跡（Audit Log），包含操作者、時間、修改前後內容。'),
  item('成績資料、學籍資料等敏感資訊之修改，須經教務主任書面授權始得執行。'),

  p([bold('（五）個人資料保護', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('本校蒐集、處理及利用學生個人資料，應依《個人資料保護法》及本校個資管理辦法辦理。'),
  item('嚴禁將師生個資傳送至個人信箱或非經授權之雲端服務。'),

  subTitle('三、', '控制重點', COLOR_ACCENT),
  new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [600, 4200, 4200],
    rows: [
      headerRow(['項次', '控制項目', '稽核查核重點'], [600, 4200, 4200]),
      dataRow(['1', '帳號管理', '離職人員帳號是否於離職當日停用？'], [600, 4200, 4200]),
      dataRow(['2', '密碼政策', '是否定期強制更換密碼？是否有鎖定機制？'], [600, 4200, 4200], true),
      dataRow(['3', '最小授權', '是否依工作職責設定最低必要權限？'], [600, 4200, 4200]),
      dataRow(['4', '稽核軌跡', '敏感資料異動是否有完整的操作日誌？'], [600, 4200, 4200], true),
      dataRow(['5', '個資保護', '師生個資是否依《個資法》蒐集、處理及利用？'], [600, 4200, 4200]),
    ]
  }),
  space(),

  subTitle('四、', '使用表單'),
  item('系統使用授權申請表。'),
  space(),
  new Paragraph({ children: [], pageBreakBefore: true })
);

// ━━━ 肆、資料輸出入及處理作業 ━━━━━━━━━━━━━━━━━━━━━━━━━━
content.push(
  sectionTitle('肆', '資料輸出入及處理作業'),
  subTitle('一、', '目的'),
  p([t('確保本校各資訊系統之資料輸入正確、輸出受控，防止資料遺失、竄改或未授權外洩。')],
    { spacing: { after: 120 } }),

  subTitle('二、', '作業程序'),

  p([bold('（一）資料輸入', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('各作業單位依原始單據（如課表、成績單、報名表）執行資料輸入，輸入前應先由主管簽核。'),
  item('系統應設置自動驗證功能，包含資料型態檢核、必填欄位、數值範圍檢查及重複資料比對。'),
  item('資料輸入完成後應留有系統處理紀錄，供事後查核。'),
  item('錯誤資料之更正，應由業務主辦單位授權專人執行；修正紀錄應記載更正前後內容及更正人員。'),
  item('批次資料匯入（如學籍、成績匯入）應於非上課時段進行，並於執行前完整備份相關資料。'),

  p([bold('（二）資料輸出', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('機密或敏感資料（師生個資、成績、財務資料）之列印或匯出，須經主管核准並留有紀錄。'),
  item('含個人資料之輸出文件，若無保存必要，應確實銷毀（紙本碎紙機、電子檔案安全刪除）。'),
  item('批次匯出資料應記錄匯出份數、使用者及時間；電子媒體匯出資料應定期確認其可讀取性。'),

  subTitle('三、', '控制重點', COLOR_ACCENT),
  new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [600, 4200, 4200],
    rows: [
      headerRow(['項次', '控制項目', '稽核查核重點'], [600, 4200, 4200]),
      dataRow(['1', '輸入驗證', '系統是否設有自動格式及範圍檢核？'], [600, 4200, 4200]),
      dataRow(['2', '輸入紀錄', '資料輸入及修正是否均有操作日誌？'], [600, 4200, 4200], true),
      dataRow(['3', '機密輸出', '敏感資料輸出是否經主管核准並留存紀錄？'], [600, 4200, 4200]),
      dataRow(['4', '資料銷毀', '含個資文件廢棄是否確實執行安全銷毀？'], [600, 4200, 4200], true),
    ]
  }),
  space(),

  subTitle('四、', '使用表單'),
  item('系統使用授權申請表。'),
  space(),
  new Paragraph({ children: [], pageBreakBefore: true })
);

// ━━━ 伍、檔案及設備安全管理 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
content.push(
  sectionTitle('伍', '檔案及設備安全管理'),
  subTitle('一、', '目的'),
  p([t('維護本校電腦機房及設備之安全，確保資料檔案的完整性與可用性，防範天災、人禍或設備故障所造成之資料損失。')],
    { spacing: { after: 120 } }),

  subTitle('二、', '作業程序'),

  p([bold('（一）機房管理', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('機房應明確訂定環境標準（溫度 18–25°C、濕度 40–60%），並配置不斷電系統（UPS）、自動穩壓設備及獨立消防系統。'),
  item('進出機房應以門禁卡或鑰匙管制，訪客進入須登記姓名、事由及陪同人員；機房應設置監視攝影設備。'),
  item('機房不得放置易燃物品；滅火器（二氧化碳）應每年定期檢驗，並記錄於「主機房工作記錄表」。'),
  item('資訊管理人員應每月定期巡查機房設備，填寫巡查紀錄表。'),

  p([bold('（二）檔案備份', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('依「檔案備份計畫」執行定期備份：重要業務系統每日備份，完整備份每週執行一次。'),
  item('備份資料採「3-2-1 原則」：3 份備份、2 種儲存媒介、1 份存放異地（或雲端）。'),
  item('每月至少執行一次備份還原測試，確認備份資料可完整回復，並記錄測試結果。'),
  item('備份媒體應定期盤點、更新，損壞媒體應依規定程序銷毀，避免資料外洩。'),

  p([bold('（三）可攜式媒體管理', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('隨身碟、光碟等可攜式媒體應登記列管，標示擁有者及用途；使用後應放置於安全位置。'),
  item('廢棄媒體（含含個資之隨身碟、硬碟）應進行實體銷毀或安全清除（低階格式化/覆寫），並留有銷毀紀錄。'),
  item('禁止使用未經設備組核可的可攜式媒體連接校內系統，防範惡意程式感染。'),

  p([bold('（四）設備安全', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('電腦教室及行政電腦設備不得攜出校外維修，如有必要，須由設備組人員全程陪同，並清除機密資料後再送修。'),
  item('各班/各辦公室電腦設備之資產標籤應完整，設備清冊每學期更新一次。'),

  subTitle('三、', '控制重點', COLOR_ACCENT),
  new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [600, 4200, 4200],
    rows: [
      headerRow(['項次', '控制項目', '稽核查核重點'], [600, 4200, 4200]),
      dataRow(['1', '機房門禁', '機房進出是否有門禁管制且留有紀錄？'], [600, 4200, 4200]),
      dataRow(['2', '備份完整性', '備份是否每月測試還原並留有紀錄？'], [600, 4200, 4200], true),
      dataRow(['3', '3-2-1 原則', '備份是否符合 3-2-1 原則（含異地備份）？'], [600, 4200, 4200]),
      dataRow(['4', '媒體銷毀', '廢棄含個資媒體是否確實安全銷毀？'], [600, 4200, 4200], true),
      dataRow(['5', '設備清冊', '設備清冊是否每學期更新，資產標籤是否完整？'], [600, 4200, 4200]),
    ]
  }),
  space(),

  subTitle('四、', '使用表單'),
  item('主機房工作記錄表。'),
  item('檔案備份計畫。'),
  item('設備資產清冊。'),
  space(),
  new Paragraph({ children: [], pageBreakBefore: true })
);

// ━━━ 陸、硬體及系統軟體使用及維護 ━━━━━━━━━━━━━━━━━━━━━━
content.push(
  sectionTitle('陸', '硬體及系統軟體使用及維護管理'),
  subTitle('一、', '目的'),
  p([t('確保本校電腦硬體與系統軟體之合法使用、正常維護，並管控採購流程以合理運用教育資源。')],
    { spacing: { after: 120 } }),

  subTitle('二、', '作業程序'),

  p([bold('（一）設備維護', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('設備組應建立各設備之維護紀錄（硬體維護管理系統或紙本），記錄故障原因、排除方法及修復時間。'),
  item('各項設備應依廠商建議排定預防性維護（PM）計畫，並至少每學期執行一次。'),
  item('設備送外維修時，應記錄送修日期、廠商及預計返回日期，並確認敏感資料已清除。'),

  p([bold('（二）軟體授權管理', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('本校所有電腦應使用具合法授權之軟體；禁止安裝未經授權或盜版軟體。'),
  item('設備組應定期（每學期）對行政電腦及電腦教室進行軟體稽查，確認授權狀態及版本合規性。'),
  item('教師及學生不得自行安裝未經設備組核可之軟體；需安裝特定教學軟體，應填寫申請單由設備組統一處理。'),

  p([bold('（三）可攜式設備管理', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('教師使用筆記型電腦等可攜式設備存取學校資源，應確保設備安裝防毒軟體及作業系統更新。'),
  item('行動裝置管理（MDM）系統應設定學生平板電腦之應用程式安裝限制及螢幕使用時間。'),

  p([bold('（四）軟硬體採購', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('軟硬體採購應依學校年度預算編列，並依採購法辦理（未達查核金額採比價，達查核金額採公開招標）。'),
  item('採購前應由設備組提出需求規格，與使用單位共同評估效益，送教務主任核准後辦理採購。'),
  item('採購完成後應進行驗收測試，確認符合規格後才正式納入資產管理清冊。'),
  item('軟體授權應確認授權數量足夠，並保存授權憑證（含序號、授權合約）備查。'),

  subTitle('三、', '控制重點', COLOR_ACCENT),
  new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [600, 4200, 4200],
    rows: [
      headerRow(['項次', '控制項目', '稽核查核重點'], [600, 4200, 4200]),
      dataRow(['1', '維護紀錄', '設備故障與維修是否完整記錄？'], [600, 4200, 4200]),
      dataRow(['2', '軟體授權', '是否每學期稽查盜版或未授權軟體？'], [600, 4200, 4200], true),
      dataRow(['3', '學生裝置', '學生平板是否部署 MDM 管理？'], [600, 4200, 4200]),
      dataRow(['4', '採購程序', '採購是否依採購法辦理？驗收是否留有紀錄？'], [600, 4200, 4200], true),
      dataRow(['5', '授權憑證', '軟體授權憑證是否妥善保存？'], [600, 4200, 4200]),
    ]
  }),
  space(),
  new Paragraph({ children: [], pageBreakBefore: true })
);

// ━━━ 柒、系統復原計畫及測試 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
content.push(
  sectionTitle('柒', '系統復原計畫及測試作業'),
  subTitle('一、', '目的'),
  p([t('確保本校關鍵資訊系統在發生故障、天災或資安事件時，能依既定程序迅速恢復運作，降低對教學行政之衝擊。')],
    { spacing: { after: 120 } }),

  subTitle('二、', '作業程序'),

  p([bold('（一）備援計畫訂定', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('設備組應訂定「系統復原計畫（DRP）」，識別本校關鍵系統（成績系統、學籍系統、選課系統等）並設定復原時間目標（RTO ≤ 4 小時）。'),
  item('備援計畫應包含：緊急聯絡名單、復原優先順序、備援設備清單及廠商支援聯絡方式。'),
  item('備援計畫每學年至少更新一次，重大變更後應立即修訂。'),

  p([bold('（二）緊急應變措施', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('發生系統故障時，使用者應立即填寫「維修申請單」並通報設備組；設備組應於 30 分鐘內回應。'),
  item('硬體故障時，設備組應洽維護廠商進行維修；如無法即時修復，應啟動備援設備或雲端系統（如 Google Workspace for Education）。'),
  item('設備損壞無法修復時，設備組應立即採購相容性設備，並從最近一次備份還原資料。'),
  item('發生資安事件（勒索軟體、入侵）時，應立即隔離受感染系統，通報教務主任，必要時通報教育部資安通報系統。'),

  p([bold('（三）復原測試', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('每學期至少執行一次系統復原演練，測試備援計畫之可行性及資料完整性。'),
  item('復原測試完成後，設備組應撰寫測試報告，包含測試範圍、發現問題及改善措施，送教務主任核閱。'),
  item('重置後之系統，應確認資料完整性後，方可恢復正式作業。'),

  subTitle('三、', '控制重點', COLOR_ACCENT),
  new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [600, 4200, 4200],
    rows: [
      headerRow(['項次', '控制項目', '稽核查核重點'], [600, 4200, 4200]),
      dataRow(['1', 'DRP 文件', '復原計畫是否每學年更新並完整記錄？'], [600, 4200, 4200]),
      dataRow(['2', 'RTO 目標', '關鍵系統復原目標是否設定且可達成？'], [600, 4200, 4200], true),
      dataRow(['3', '復原演練', '是否每學期執行復原演練並留有報告？'], [600, 4200, 4200]),
      dataRow(['4', '資安通報', '發生資安事件是否依規定通報教育部？'], [600, 4200, 4200], true),
    ]
  }),
  space(),

  subTitle('四、', '使用表單'),
  item('系統維修申請單。'),
  item('系統復原計畫（DRP）。'),
  item('復原測試報告。'),
  space(),
  new Paragraph({ children: [], pageBreakBefore: true })
);

// ━━━ 捌、資訊安全檢查作業 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
content.push(
  sectionTitle('捌', '資訊安全檢查作業'),
  subTitle('一、', '目的'),
  p([t('透過定期及不定期之資安檢查，偵測本校網路環境及資訊系統之安全弱點，防範駭客入侵、資料外洩及惡意程式感染。')],
    { spacing: { after: 120 } }),

  subTitle('二、', '作業程序'),

  p([bold('（一）網路安全管理', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('設備組應部署防火牆（Firewall）及入侵偵測/防禦系統（IDS/IPS），並定期審視防火牆規則。'),
  item('校園網路應劃分為行政網段、教學網段及訪客 Wi-Fi 三個獨立區域，各網段之間存取應受管控。'),
  item('郵件伺服器應部署防垃圾郵件及防毒功能；不明附件或連結應提醒師生勿輕易點擊。'),
  item('對外開放之服務（網站、教務系統）應定期（每半年）進行弱點掃描，發現高風險漏洞應於 72 小時內修補。'),

  p([bold('（二）端點防護', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('全校電腦應安裝防毒軟體，病毒碼每日自動更新；作業系統資安修補應於公告後 30 天內套用。'),
  item('禁止師生使用 P2P 軟體或瀏覽非法網站；學生上網行為應透過網路閘道進行管控與記錄。'),
  item('電腦教室應設定使用者不得自行安裝軟體；學生操作環境可採用磁碟還原系統（如 Deep Freeze）。'),

  p([bold('（三）資安教育訓練', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('每學期至少對教職員工辦理一次資安宣導，內容包含釣魚郵件辨識、密碼安全、個資保護。'),
  item('新進教職員工應於到職一個月內完成資安基礎教育訓練。'),
  item('對學生應於每學年開學時進行資訊倫理與網路安全宣導，並納入資訊課程。'),

  p([bold('（四）資安事件通報', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('發現資安事件（資料外洩、惡意程式、系統入侵）應立即通報設備組長及教務主任。'),
  item('重大資安事件應依《資通安全管理法》及教育部資安規定，於 1 小時內向教育局/教育部通報。'),
  item('資安事件應留有完整紀錄，包含事件描述、影響範圍、處置措施及後續改善計畫。'),

  p([bold('（五）機密資料保護', COLOR_H2)], { spacing: { before: 120, after: 80 } }),
  item('含師生個資或機密資料之檔案，應加密儲存（AES-256 或同等強度）；傳輸應使用 HTTPS 或加密通道。'),
  item('禁止教職員工將含個資之資料儲存於個人雲端（Google Drive 個人帳號、Dropbox 等），應使用本校核可之雲端服務。'),

  subTitle('三、', '控制重點', COLOR_ACCENT),
  new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [600, 4200, 4200],
    rows: [
      headerRow(['項次', '控制項目', '稽核查核重點'], [600, 4200, 4200]),
      dataRow(['1', '防火牆', '防火牆規則是否定期審視？IDS/IPS 是否啟用？'], [600, 4200, 4200]),
      dataRow(['2', '弱點掃描', '對外服務是否每半年執行弱點掃描？'], [600, 4200, 4200], true),
      dataRow(['3', '防毒更新', '全校電腦防毒碼是否每日更新？'], [600, 4200, 4200]),
      dataRow(['4', '資安訓練', '教職員每學期是否完成資安宣導？'], [600, 4200, 4200], true),
      dataRow(['5', '事件通報', '資安事件是否依規定時限通報教育部？'], [600, 4200, 4200]),
      dataRow(['6', '資料加密', '機密檔案是否加密儲存？傳輸是否使用加密通道？'], [600, 4200, 4200], true),
    ]
  }),
  space(),
  new Paragraph({ children: [], pageBreakBefore: true })
);

// ━━━ 玖、附則 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
content.push(
  sectionTitle('玖', '附則'),

  subTitle('一、', '修訂程序'),
  item('本制度每學年至少檢討修訂一次；法規修訂或資訊環境重大變更時，應即時調整。'),
  item('修訂程序：設備組提案 → 教務主任核准 → 行政會議通過 → 公告施行。'),

  subTitle('二、', '違規處理'),
  item('教職員工違反本制度，視情節輕重依本校教職員工考核辦法或相關規定辦理。'),
  item('學生違反本制度，依學生獎懲辦法辦理；情節嚴重者移請法律途徑處理。'),

  subTitle('三、', '施行'),
  item('本制度經校長核准後公告施行，修訂時亦同。'),

  space(360),
  p([
    t('校　長：＿＿＿＿＿＿', { size: 24 }),
    new TextRun({ text: '        ', size: 24 }),
    t('教務主任：＿＿＿＿＿＿', { size: 24 }),
  ], { spacing: { before: 240, after: 120 }, alignment: AlignmentType.LEFT }),
  p([
    t('設備組長：＿＿＿＿＿＿', { size: 24 }),
    new TextRun({ text: '        ', size: 24 }),
    t('訂定日期：＿＿年＿＿月＿＿日', { size: 24 }),
  ], { spacing: { after: 120 } }),
);

// ─── 建立文件 ──────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '●',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 320 } } }
        }]
      },
      {
        reference: 'bullets2',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '◆',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: {
      document: { run: { font: { name: FONT }, size: 24, color: '222222' } }
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: { name: FONT }, color: COLOR_H1 },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: { name: FONT }, color: COLOR_H2 },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: '○○高中　資訊內部控制作業制度', font: { name: FONT }, size: 18, color: '888888' }),
              new TextRun({ children: ['\t'], font: { name: FONT } }),
              new TextRun({ children: ['第 ', PageNumber.CURRENT, ' 頁'], font: { name: FONT }, size: 18, color: '888888' }),
            ],
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER, space: 4 } }
          })
        ]
      })
    },
    children: content
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('D:\\D\\114設備組\\高中資訊內控制度.docx', buf);
  console.log('DONE');
}).catch(e => { console.error(e); process.exit(1); });
