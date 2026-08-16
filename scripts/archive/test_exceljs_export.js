const ExcelJS = require('exceljs');
const fs = require('fs');

const entries = JSON.parse(fs.readFileSync('evaluation_data.json', 'utf8'));

const COLOR_PINK_BLUSH = "FFF4CCCC";   // Top chunk header (Soft coral blush)
const COLOR_PEACH_CREAM = "FFFCE5CD";  // Sub-table headers (Soft apricot cream)
const COLOR_WHITE = "FFFFFFFF";
const COLOR_WARN_BG = "FFFFF2D6";      // Pending self-eval

const COLOR_L5_BG = "FFE4ECD3";        // Soft Sage (Amazing)
const COLOR_L4_BG = "FFE2F3F0";        // Soft Cyan (Good)
const COLOR_L3_BG = "FFFFF4CD";        // Soft Butter (Keep)
const COLOR_L2_BG = "FFFCE5CD";        // Soft Apricot (Grow)
const COLOR_L1_BG = "FFF4CCCC";        // Soft Coral (Start)

const thinBorder = {
  top: { style: 'thin', color: { argb: 'FFD7CCC8' } },
  left: { style: 'thin', color: { argb: 'FFD7CCC8' } },
  bottom: { style: 'thin', color: { argb: 'FFD7CCC8' } },
  right: { style: 'thin', color: { argb: 'FFD7CCC8' } }
};

const fontChunkTitle = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FF3E2723' } };
const fontSubHeader = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FF4E342E' } };
const fontBody = { name: '微軟正黑體', size: 9.5, color: { argb: 'FF2D2323' } };
const fontBodyBold = { name: '微軟正黑體', size: 9.5, bold: true, color: { argb: 'FF2D2323' } };
const fontWarn = { name: '微軟正黑體', size: 9.5, bold: true, color: { argb: 'FFB45309' } };

const alignCenter = { horizontal: 'center', vertical: 'middle', wrapText: true };
const alignLeft = { horizontal: 'left', vertical: 'top', wrapText: true };
const alignHeader = { horizontal: 'center', vertical: 'middle', wrapText: true };

const SUPERVISOR_TEAMS = {
  "張希慈": ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔"],
  "何維安": ["林文琇"],
  "姚品瑄": ["薛筑瑄", "戴佑珍"],
  "張希慈_執行長": ["張希慈"]
};

const RATING_GUIDE_ROWS = [
  ["L5", "Amazing!", "遠超職位期待，表現為團隊之標竿與典範。", "讚賞其突出貢獻，探討經驗複製機制，在對話中轉為「帶其他人一起做」方向。", COLOR_L5_BG],
  ["L4", "Good", "優於職位期待，持續展現高標準成果。", "肯定並具體指出哪些行為超出標準，設定具挑戰性的下一步目標。", COLOR_L4_BG],
  ["L3", "Keep", "符合職位門檻，展現穩定的工作交付。", "確認穩定度，指出下一階可以往前之處，維持節奏並選一至兩項深化。", COLOR_L3_BG],
  ["L2", "Grow", "部分符合，部分能力/行為仍在建立階段。", "說明落差所在，聚焦一項具體可練習的行為，納入下一期 IDP 設定指標。", COLOR_L2_BG],
  ["L1", "Start", "目前的具體事證還看不到這項職能，或事證與職能要求落差明顯。", "明確對齊職位基本門檻與要求，提供即時支援與回饋引導。", COLOR_L1_BG],
];

function styleRange(ws, startRow, startCol, endRow, endCol, font, fill, alignment) {
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const cell = ws.getCell(r, c);
      if (font) cell.font = font;
      if (fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
      cell.border = thinBorder;
      if (alignment) cell.alignment = alignment;
    }
  }
}

async function generateSupervisorSheet(wb, supervisorName, memberNames, allEntries) {
  const ws = wb.addWorksheet(supervisorName === "張希慈_執行長" ? "自評_執行長" : `自評_${supervisorName}`, {
    views: [{ showGridLines: true }]
  });

  ws.columns = [
    { width: 16 },
    { width: 36 },
    { width: 68 },
    { width: 16 },
    { width: 34 }
  ];

  // 1. Title Banner (Merged A1:E1)
  ws.mergeCells(1, 1, 1, 5);
  const tCell = ws.getCell(1, 1);
  tCell.value = `好好星球文化基金會 360 年中成長評估 - 【${supervisorName}】部屬自評與主管評核表`;
  styleRange(ws, 1, 1, 1, 5, { name: '微軟正黑體', size: 13, bold: true, color: { argb: 'FF3E2723' } }, COLOR_PINK_BLUSH, { horizontal: 'left', vertical: 'middle', indent: 1 });
  ws.getRow(1).height = 32;

  // 2. Rating Guide Section Header (Merged A2:E2)
  ws.mergeCells(2, 1, 2, 5);
  const gTitle = ws.getCell(2, 1);
  gTitle.value = "職能評分標準與面向分數落點說明（不對員工公布分數與總分，只回饋落點）";
  styleRange(ws, 2, 1, 2, 5, { name: '微軟正黑體', size: 10.5, bold: true, color: { argb: 'FF3E2723' } }, COLOR_PEACH_CREAM, { horizontal: 'left', vertical: 'middle', indent: 1 });
  ws.getRow(2).height = 24;

  // 3. Rating Guide Table Headers
  ws.getCell(3, 1).value = "等級 (Level)";
  ws.getCell(3, 2).value = "落點名稱";
  ws.mergeCells(3, 3, 3, 4);
  ws.getCell(3, 3).value = "定義說明";
  ws.getCell(3, 5).value = "回饋語氣與後續動作";
  styleRange(ws, 3, 1, 3, 5, fontSubHeader, COLOR_PEACH_CREAM, alignHeader);
  ws.getRow(3).height = 22;

  // 4. Rating Guide Rows
  RATING_GUIDE_ROWS.forEach((row, i) => {
    const r = 4 + i;
    const [lvl, name, def, act, bg] = row;
    ws.getCell(r, 1).value = lvl;
    ws.getCell(r, 2).value = name;
    ws.mergeCells(r, 3, r, 4);
    ws.getCell(r, 3).value = def;
    ws.getCell(r, 5).value = act;

    styleRange(ws, r, 1, r, 5, fontBody, bg, alignLeft);
    ws.getCell(r, 1).alignment = alignCenter;
    ws.getCell(r, 1).font = fontBodyBold;
    ws.getCell(r, 2).alignment = alignCenter;
    ws.getCell(r, 2).font = fontBodyBold;
    ws.getRow(r).height = 26;
  });

  let rowIdx = 10;
  for (const member of memberNames) {
    const memEntry = allEntries.find(e => e.target === member && e.relation === "自評");
    const roleStr = memEntry && memEntry.self_eval ? memEntry.self_eval.job_role : "（尚未填寫自評）";
    const timeStr = memEntry ? ` ｜ 填答時間：${memEntry.timestamp}` : "";

    // Chunk Top Header Banner
    ws.mergeCells(rowIdx, 1, rowIdx, 5);
    ws.getCell(rowIdx, 1).value = `部屬姓名：${member}    ｜    職位：${roleStr}${timeStr}`;
    styleRange(ws, rowIdx, 1, rowIdx, 5, fontChunkTitle, COLOR_PINK_BLUSH, { horizontal: 'left', vertical: 'middle', indent: 1 });
    ws.getRow(rowIdx).height = 28;
    rowIdx++;

    if (!memEntry || !memEntry.self_eval) {
      ws.getCell(rowIdx, 1).value = "自評狀態";
      styleRange(ws, rowIdx, 1, rowIdx, 1, fontBodyBold, COLOR_PEACH_CREAM, alignCenter);

      ws.mergeCells(rowIdx, 2, rowIdx, 5);
      ws.getCell(rowIdx, 2).value = `目前表單中尚未收到 ${member} 的自我評估回覆紀錄。收到新回覆後可重新載入。`;
      styleRange(ws, rowIdx, 2, rowIdx, 5, fontWarn, COLOR_WARN_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(rowIdx).height = 28;
      rowIdx += 3;
      continue;
    }

    const se = memEntry.self_eval;

    // 工作特質盤點
    ws.mergeCells(rowIdx, 1, rowIdx + 1, 1);
    ws.getCell(rowIdx, 1).value = "工作特質盤點";
    styleRange(ws, rowIdx, 1, rowIdx + 1, 1, fontBodyBold, COLOR_PEACH_CREAM, alignCenter);

    ws.getCell(rowIdx, 2).value = "最穩定、最具代表性 Top 3";
    styleRange(ws, rowIdx, 2, rowIdx, 2, fontBodyBold, COLOR_PEACH_CREAM, alignCenter);
    ws.mergeCells(rowIdx, 3, rowIdx, 5);
    ws.getCell(rowIdx, 3).value = (se.top3_stable || []).join("、") || "（無）";
    styleRange(ws, rowIdx, 3, rowIdx, 5, fontBody, COLOR_WHITE, { horizontal: 'left', vertical: 'middle', indent: 1 });
    ws.getRow(rowIdx).height = 24;
    rowIdx++;

    ws.getCell(rowIdx, 2).value = "目前在練習 / 期望發展 3 項";
    styleRange(ws, rowIdx, 2, rowIdx, 2, fontBodyBold, COLOR_PEACH_CREAM, alignCenter);
    ws.mergeCells(rowIdx, 3, rowIdx, 5);
    ws.getCell(rowIdx, 3).value = (se.top3_practice || []).join("、") || "（無）";
    styleRange(ws, rowIdx, 3, rowIdx, 5, fontBody, COLOR_WHITE, { horizontal: 'left', vertical: 'middle', indent: 1 });
    ws.getRow(rowIdx).height = 24;
    rowIdx++;

    // 四大文化實踐
    ws.mergeCells(rowIdx, 1, rowIdx, 5);
    ws.getCell(rowIdx, 1).value = "四大文化實踐實例（STAR 敘述）";
    styleRange(ws, rowIdx, 1, rowIdx, 5, fontSubHeader, COLOR_PEACH_CREAM, { horizontal: 'left', vertical: 'middle', indent: 1 });
    ws.getRow(rowIdx).height = 24;
    rowIdx++;

    ws.getCell(rowIdx, 1).value = "評估面向";
    ws.getCell(rowIdx, 2).value = "四大文化定義與說明";
    ws.getCell(rowIdx, 3).value = "部屬自評實例 (STAR)";
    ws.mergeCells(rowIdx, 4, rowIdx, 5);
    ws.getCell(rowIdx, 4).value = "組織文化整體主管回饋 (Feedback)";
    styleRange(ws, rowIdx, 1, rowIdx, 5, fontSubHeader, COLOR_PEACH_CREAM, alignHeader);
    ws.getRow(rowIdx).height = 24;
    rowIdx++;

    const cultStart = rowIdx;
    const cultRows = [
      ["【信任】獨立行動與決策、主動協作、雙向溝通，在高彈性下給予彼此信任", se.values?.['信任'] || "（無）"],
      ["【多元】尊重差異、多元工作方法、主動表達不同觀點與想法", se.values?.['多元'] || "（無）"],
      ["【實驗】透過開放心態嘗試修正與反思，勇於檢討及給予回饋", se.values?.['實驗'] || "（無）"],
      ["【可持續】內在韌性、自我照顧、彈性的人際與工作邊界", se.values?.['可持續'] || "（無）"],
    ];

    cultRows.forEach(([title, ans]) => {
      ws.getCell(rowIdx, 2).value = title;
      ws.getCell(rowIdx, 3).value = ans;
      styleRange(ws, rowIdx, 2, rowIdx, 3, fontBody, COLOR_WHITE, alignLeft);
      const textLen = (ans || "").length;
      ws.getRow(rowIdx).height = Math.max(28, Math.min(140, Math.floor(textLen / 45 * 18) + 20));
      rowIdx++;
    });
    const cultEnd = rowIdx - 1;

    ws.mergeCells(cultStart, 1, cultEnd, 1);
    ws.getCell(cultStart, 1).value = "組織文化";
    styleRange(ws, cultStart, 1, cultEnd, 1, fontBodyBold, COLOR_WHITE, alignCenter);

    ws.mergeCells(cultStart, 4, cultEnd, 5);
    ws.getCell(cultStart, 4).value = "";
    styleRange(ws, cultStart, 4, cultEnd, 5, fontBody, COLOR_WHITE, alignLeft);

    // 職能展現
    if (se.competencies && se.competencies.length > 0) {
      ws.mergeCells(rowIdx, 1, rowIdx, 5);
      ws.getCell(rowIdx, 1).value = `職位專屬職能展現實例【${se.job_role}】`;
      styleRange(ws, rowIdx, 1, rowIdx, 5, fontSubHeader, COLOR_PEACH_CREAM, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(rowIdx).height = 24;
      rowIdx++;

      ws.getCell(rowIdx, 1).value = "評估面向";
      ws.getCell(rowIdx, 2).value = "職能項目與題目定義";
      ws.getCell(rowIdx, 3).value = "部屬自評實例";
      ws.getCell(rowIdx, 4).value = "評分 (Lv.1-5選單)";
      ws.getCell(rowIdx, 5).value = "主管回饋 (Feedback)";
      styleRange(ws, rowIdx, 1, rowIdx, 5, fontSubHeader, COLOR_PEACH_CREAM, alignHeader);
      ws.getRow(rowIdx).height = 24;
      rowIdx++;

      const compStart = rowIdx;
      se.competencies.forEach(comp => {
        ws.getCell(rowIdx, 2).value = comp.title;
        ws.getCell(rowIdx, 3).value = comp.answer || "（無填寫）";
        ws.getCell(rowIdx, 4).value = "";
        ws.getCell(rowIdx, 5).value = "";

        // Add DataValidation for Column D in ExcelJS
        ws.getCell(rowIdx, 4).dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: ['"L1,L2,L3,L4,L5"'],
          showErrorMessage: true,
          errorTitle: '評分無效',
          error: '請從下拉選單中選取 L1 到 L5',
          showInputMessage: true,
          promptTitle: '等級選單',
          prompt: '請選取評核等級：L1(Start), L2(Grow), L3(Keep), L4(Good), L5(Amazing!)'
        };

        styleRange(ws, rowIdx, 2, rowIdx, 5, fontBody, COLOR_WHITE, alignLeft);
        ws.getCell(rowIdx, 4).alignment = alignCenter;
        ws.getCell(rowIdx, 4).font = fontBodyBold;

        const textLen = (comp.answer || "").length;
        ws.getRow(rowIdx).height = Math.max(28, Math.min(140, Math.floor(textLen / 45 * 18) + 20));
        rowIdx++;
      });
      const compEnd = rowIdx - 1;

      ws.mergeCells(compStart, 1, compEnd, 1);
      ws.getCell(compStart, 1).value = "專業職能";
      styleRange(ws, compStart, 1, compEnd, 1, fontBodyBold, COLOR_WHITE, alignCenter);
    }

    rowIdx += 2;
  }
}

async function run() {
  const wb = new ExcelJS.Workbook();
  await generateSupervisorSheet(wb, "何維安", ["林文琇"], entries);
  await wb.xlsx.writeFile("test_exceljs_output.xlsx");
  console.log("Exported test_exceljs_output.xlsx successfully with colors!");
}

run();
