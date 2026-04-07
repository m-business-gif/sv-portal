// =============================================
// SV加盟店管理ポータル - GAS API v4
// =============================================

const SS_ID        = "1K-4ub8YvFh__JrseNKGiCkGigDYykraIwocOhLQLevY";
const SHEET_GOAL   = "加盟店目標";
const SHEET_REAL   = "実数値";
const SHEET_TASK   = "タスクボード";
const SHEET_MIKOMI = "見込み数値";
const SHEET_STAFF  = "スタッフランク";
const SHEET_SALES  = "売上明細（9~3月）";
const SHEET_CONFIG = "【眉毛】加盟店管理集計";

// タスクボード列定義
// A:店舗名 B:担当SV C:カテゴリ D:タスク名 E:ステータス F:優先度 G:メモ H:完了日

// 売上明細（9~3月）列定義 ※略称列(A)挿入後
const SC_ABBR  = 0;   // A: 略称
const SC_NAME  = 1;   // B: 正式名称
const SC_DATE  = 2;   // C: 日付(YYYYMMDD)
const SC_KUBUN = 6;   // G: 区分
const SC_CAT   = 8;   // I: カテゴリ
const SC_CAT2  = 9;   // J: カテゴリ2
const SC_COUNT = 12;  // M: 件数
const SC_AMT   = 13;  // N: 金額
const SC_STAFF = 14;  // O: スタッフ名

function doGet(e) {
  try {
    // アジェンダ生成リクエスト（JSONデータ）
    if (e && e.parameter && e.parameter.action === "createAgenda") {
      const result = createAgenda(
        e.parameter.store || "",
        e.parameter.format || "doc",
        decodeURIComponent(e.parameter.memo || "")
      );
      return json(result);
    }
    // 売上明細に略称列を挿入: ?action=addAbbrColumn
    if (e && e.parameter && e.parameter.action === "addAbbrColumn") {
      return json({ result: addAbbrColumnToSales() });
    }
    // 売上データ同期実行: ?action=syncSalesData
    if (e && e.parameter && e.parameter.action === "syncSalesData") {
      syncSalesData();
      return json({ ok: true, message: "同期完了" });
    }
    // ベストプラクティス読み込み（デバッグ用）
    if (e && e.parameter && e.parameter.action === "getBestPractices") {
      return json(getBestPractices());
    }
    // 売上明細のユニーク店舗名一覧: ?action=listSalesStores
    if (e && e.parameter && e.parameter.action === "listSalesStores") {
      const ss2 = SpreadsheetApp.openById(SS_ID);
      const ws2 = ss2.getSheetByName(SHEET_SALES);
      if (!ws2) return json({ error: "売上明細シートなし" });
      const rows2 = ws2.getDataRange().getValues();
      const nameCount = {};
      for (let i = 1; i < rows2.length; i++) {
        const nm = String(rows2[i][SC_NAME] || "").trim();
        if (nm) nameCount[nm] = (nameCount[nm] || 0) + 1;
      }
      return json({ total: Object.keys(nameCount).length, stores: nameCount });
    }
    // メニュー比率デバッグ: ?action=debugMenu&store=店舗名&ym=202603
    if (e && e.parameter && e.parameter.action === "debugMenu") {
      const storeName = e.parameter.store || "";
      const ym = parseInt(e.parameter.ym) || 202603;
      const nameMap = _getSalesNameMap();
      const officialName = nameMap.abbrToOfficial[storeName] || storeName;
      const ss2 = SpreadsheetApp.openById(SS_ID);
      const ws2 = ss2.getSheetByName(SHEET_SALES);
      const rows2 = ws2 ? ws2.getDataRange().getValues() : [];
      const sample = [];
      let matchCount = 0, kubunCount = 0;
      const ymDist = {};
      for (let i = 1; i < rows2.length; i++) {
        const store = String(rows2[i][SC_NAME] || "").trim();
        if (store !== officialName) continue;
        matchCount++;
        const kubun = String(rows2[i][SC_KUBUN] || "").trim();
        const rawDate = rows2[i][SC_DATE];
        const dn = parseFloat(rawDate) || 0;
        const rowYM = Math.floor(dn / 100);
        ymDist[rowYM] = (ymDist[rowYM] || 0) + 1;
        const mn = (String(rows2[i][SC_CAT] || "") + " " + String(rows2[i][SC_CAT2] || "")).trim();
        if (sample.length < 5) sample.push({ dn, rowYM, kubun, mn });
        if (kubun !== "施術") continue;
        kubunCount++;
      }
      return json({ storeName, officialName, matchCount, kubunCount, ymDist, sample });
    }
    // 指定月の店舗データ取得
    if (e && e.parameter && e.parameter.action === "getStores") {
      const ym = parseInt(e.parameter.ym) || null;
      return json({ stores: getStores(ym) });
    }
    // 店舗レポート: ?action=getStoreReport&store=店舗名[&ym=202603]
    if (e && e.parameter && e.parameter.action === "getStoreReport") {
      const staffYM = e.parameter.ym ? parseInt(e.parameter.ym) : null;
      return json(getStoreReport(e.parameter.store || "", staffYM));
    }
    // アジェンダ外部ファイル生成（Google Docs/Slides）
    if (e && e.parameter && e.parameter.action === "createAgendaExternal") {
      const url = createAgendaExternal(
        e.parameter.store || "",
        e.parameter.format || "doc",
        decodeURIComponent(e.parameter.memo || "")
      );
      return json({ ok: true, url });
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetNames = ss.getSheets().map(s => s.getName());
    const required = [SHEET_GOAL, SHEET_REAL, SHEET_MIKOMI];
    const missing = required.filter(n => !sheetNames.includes(n));
    if (missing.length > 0) {
      return json({ error: "シートが見つかりません: " + missing.join(", ") });
    }
    // タスクボードが存在しないか旧形式なら自動セットアップ
    if (!sheetNames.includes(SHEET_TASK)) {
      setupTaskBoard();
    } else {
      const ws = ss.getSheetByName(SHEET_TASK);
      const h = ws.getRange(1,1,1,1).getValue();
      if (String(h||"").trim() !== "店舗名") setupTaskBoard();
    }
    // 設定シートが存在しなければ自動作成
    if (!sheetNames.includes(SHEET_CONFIG)) setupConfig();
    const safe = fn => { try { return fn(); } catch(err2) { Logger.log("doGet safe error: " + err2); return null; } };
    return json({
      stores:          safe(getStores) || [],
      availableMonths: safe(getAvailableMonths) || [],
      tasks:           safe(getTasks) || [],
      staffRanks:      safe(getStaffRanks) || [],
      staffSales:      safe(getStaffSales) || [],
      config:          safe(getConfig) || {}
    });
  } catch(err) {
    return json({ error: err.toString() });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === "upsertTask") {
      const row = upsertTask(data.row, data.task);
      return json({ ok: true, row });
    }
    if (data.action === "deleteTask") {
      deleteTaskRow(data.row);
      return json({ ok: true });
    }
    if (data.action === "updateStaffRank") {
      updateStaffRank(data.store, data.staff, data.rank, data.score);
      return json({ ok: true });
    }
    if (data.action === "deleteStaff") {
      deleteStaff(data.store, data.staff);
      return json({ ok: true });
    }
    if (data.action === "addStore") {
      addStoreFn(data.sv, data.name, data.ym, data.goals || {});
      return json({ ok: true });
    }
    if (data.action === "updateGoal") {
      updateGoalFn(data.name, data.ym, data.goals || {});
      return json({ ok: true });
    }
    if (data.action === "createAgenda") {
      const result = createAgenda(data.store, data.format || "doc", data.memo || "");
      return json({ ok: true, html: result.html, title: result.title });
    }
    if (data.action === "repairGoalSheet") {
      const count = repairGoalSheetFn();
      return json({ ok: true, repairedRows: count });
    }
    if (data.action === "syncGoalData") {
      const count = syncGoalDataFn();
      return json({ ok: true, syncedRows: count });
    }
    if (data.action === "restoreImportRange") {
      restoreImportRangeFn();
      return json({ ok: true });
    }
    return json({ error: "unknown action" });
  } catch(err) {
    return json({ error: err.toString() });
  }
}

// ─ 設定シート ─
function setupConfig() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let ws = ss.getSheetByName(SHEET_CONFIG);
  if (!ws) ws = ss.insertSheet(SHEET_CONFIG);
  // 既にデータがある場合は上書きしない
  if (ws.getLastRow() > 0) return;
  ws.getRange(1,1,1,3).setValues([["種別","値1","値2"]]);
  ws.getRange(1,1,1,3).setFontWeight("bold").setBackground("#f1f5f9");
  const defaults = [
    ["SV","山田","#2563eb"],
    ["SV","髙橋","#7c3aed"],
    ["SV","向井","#0891b2"],
    ["SV","子龍","#059669"],
    ["SV","宮脇","#c2410c"],
    ["カテゴリ","集客",""],
    ["カテゴリ","教育",""],
    ["カテゴリ","店舗運営",""],
    ["カテゴリ","数値管理",""],
    ["カテゴリ","MTG",""],
    ["カテゴリ","HP/SNS",""],
    ["カテゴリ","その他",""],
  ];
  ws.getRange(2,1,defaults.length,3).setValues(defaults);
  ws.setColumnWidths(1,3,[100,120,100]);
}

function getConfig() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_CONFIG);
  if (!ws) return { svList: [], taskCats: [] };
  const rows = ws.getDataRange().getValues();
  const svList = [];
  const taskCats = [];
  for (let i = 1; i < rows.length; i++) {
    const type = String(rows[i][0]||"").trim();
    const v1   = String(rows[i][1]||"").trim();
    const v2   = String(rows[i][2]||"").trim();
    if (!v1) continue;
    if (type === "SV") svList.push({ name: v1, color: v2||"#6b7280" });
    else if (type === "カテゴリ") taskCats.push(v1);
  }
  return { svList, taskCats };
}

// ─ タスクボード初期セットアップ ─
function setupTaskBoard() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let ws = ss.getSheetByName(SHEET_TASK);
  if (!ws) ws = ss.insertSheet(SHEET_TASK);
  ws.clearContents();
  ws.getRange(1,1,1,8).setValues([["店舗名","担当SV","カテゴリ","タスク名","ステータス","優先度","メモ","完了日"]]);
  ws.getRange(1,1,1,8).setFontWeight("bold").setBackground("#f1f5f9");
  ws.setFrozenRows(1);
  ws.setColumnWidths(1,8,[120,80,100,220,80,60,200,90]);
}

// ─ 旧形式→新形式 移行（GASエディタで1回実行） ─
function migrateTaskBoard() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_TASK);
  if (!ws) { Logger.log("タスクボードシートが見つかりません"); return; }

  const rows = ws.getDataRange().getValues();
  const b = v => v === true || String(v).toUpperCase() === "TRUE";

  // 旧形式かどうか判定（A1がSVなら旧形式）
  if (String(rows[0][0]).trim() !== "SV") {
    Logger.log("既に新形式です。移行をスキップします。");
    return;
  }

  const newRows = [["店舗名","担当SV","カテゴリ","タスク名","ステータス","優先度","メモ","完了日"]];

  // 旧形式列: SV(0), 店舗名(1), 施策(2), 次回予約特典(3),
  //           ミニモ導入(4), スタッフMTG(5), オーナーMTG(6),
  //           HPBミーティング(7), 店舗集計分析(8), LINE配信(9),
  //           回数券POP共有(10), [空](11), 未実行タスク(12)
  const checkItems = [
    {col:4, cat:"集客",    name:"ミニモ導入"},
    {col:5, cat:"MTG",     name:"スタッフMTG"},
    {col:6, cat:"MTG",     name:"オーナーMTG"},
    {col:7, cat:"MTG",     name:"HPBミーティング"},
    {col:8, cat:"数値管理", name:"店舗集計分析"},
    {col:9, cat:"HP/SNS",  name:"LINE配信"},
    {col:10,cat:"集客",    name:"回数券POP共有"},
  ];

  const seen = {};
  for (let i = 1; i < rows.length; i++) {
    const sv    = String(rows[i][0] || "").trim();
    const store = String(rows[i][1] || "").trim();
    if (!store) continue;

    const key = sv + "|" + store;
    if (seen[key]) continue; // 重複行をスキップ
    seen[key] = true;

    const shisku = String(rows[i][2]  || "").trim();
    const yoyaku = String(rows[i][3]  || "").trim();
    const misei  = String(rows[i][12] || "").trim();

    // 施策テキスト
    if (shisku) newRows.push([store, sv, "集客", shisku, "進行中", "中", "", ""]);
    // 次回予約特典テキスト
    if (yoyaku) newRows.push([store, sv, "集客", "次回予約特典: " + yoyaku, "完了", "中", yoyaku, ""]);
    // チェックボックス項目（TRUE=完了、FALSE=未着手）
    checkItems.forEach(({col, cat, name}) => {
      const status = b(rows[i][col]) ? "完了" : "未着手";
      newRows.push([store, sv, cat, name, status, "中", "", ""]);
    });
    // 未実行タスク
    if (misei) newRows.push([store, sv, "その他", misei, "未着手", "高", "", ""]);
  }

  // シートを新形式で上書き
  ws.clearContents();
  if (newRows.length > 0) {
    ws.getRange(1, 1, newRows.length, 8).setValues(newRows);
    ws.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#f1f5f9");
    ws.setFrozenRows(1);
    ws.setColumnWidths(1, 8, [120, 80, 100, 220, 80, 60, 200, 90]);
    // ステータス列に色付け
    for (let r = 2; r <= newRows.length; r++) {
      const sts = String(newRows[r-1][4] || "");
      const bg = sts === "完了" ? "#dcfce7" : sts === "進行中" ? "#fef9c3" : "#f8fafc";
      ws.getRange(r, 5).setBackground(bg);
    }
  }
  Logger.log("移行完了: " + (newRows.length - 1) + "件のタスクを作成しました。");
}

// ─ タスク取得 ─
function getTasks() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_TASK);
  if (!ws) return [];
  const rows = ws.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const tasks = [];
  for (let i = 1; i < rows.length; i++) {
    const store    = String(rows[i][0] || "").trim();
    const sv       = String(rows[i][1] || "").trim();
    const category = String(rows[i][2] || "").trim();
    const taskName = String(rows[i][3] || "").trim();
    const status   = String(rows[i][4] || "未着手").trim() || "未着手";
    const priority = String(rows[i][5] || "中").trim() || "中";
    const memo     = String(rows[i][6] || "").trim();
    let doneDate = "";
    if (rows[i][7]) {
      try { doneDate = Utilities.formatDate(new Date(rows[i][7]), "Asia/Tokyo", "yyyy-MM-dd"); } catch(e) { doneDate = String(rows[i][7]); }
    }
    if (!store && !taskName) continue;
    tasks.push({ store, sv, category, taskName, status, priority, memo, doneDate, row: i + 1 });
  }
  return tasks;
}

// ─ タスク追加・更新 ─
function upsertTask(rowNum, taskData) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_TASK);
  if (!ws) return -1;

  let doneDate = taskData.doneDate || "";
  if (taskData.status === "完了" && !doneDate) {
    doneDate = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
  }
  if (taskData.status !== "完了") doneDate = "";

  if (!rowNum || rowNum < 2) {
    ws.appendRow([
      taskData.store    || "",
      taskData.sv       || "",
      taskData.category || "",
      taskData.taskName || "",
      taskData.status   || "未着手",
      taskData.priority || "中",
      taskData.memo     || "",
      doneDate,
    ]);
    return ws.getLastRow();
  } else {
    const r = ws.getRange(rowNum, 1, 1, 8);
    const cur = r.getValues()[0];
    r.setValues([[
      taskData.store    !== undefined ? (taskData.store    || "") : cur[0],
      taskData.sv       !== undefined ? (taskData.sv       || "") : cur[1],
      taskData.category !== undefined ? (taskData.category || "") : cur[2],
      taskData.taskName !== undefined ? (taskData.taskName || "") : cur[3],
      taskData.status   !== undefined ? (taskData.status   || "未着手") : cur[4],
      taskData.priority !== undefined ? (taskData.priority || "中")     : cur[5],
      taskData.memo     !== undefined ? (taskData.memo     || "") : cur[6],
      doneDate !== "" || taskData.status === "完了" ? doneDate : cur[7],
    ]]);
    return rowNum;
  }
}

// ─ タスク削除 ─
function deleteTaskRow(rowNum) {
  if (!rowNum || rowNum < 2) return;
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_TASK);
  if (ws) ws.deleteRow(rowNum);
}

// ─ 利用可能な月一覧 ─
function getAvailableMonths() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const ws = ss.getSheetByName(SHEET_GOAL);
    if (!ws) return [];
    const rows = ws.getDataRange().getValues();
    const yms = new Set();
    const today = new Date();
    const curYM = today.getFullYear() * 100 + (today.getMonth() + 1);
    yms.add(curYM);
    for (let i = 1; i < rows.length; i++) {
      const dv = Math.round(parseFloat(rows[i][3]) || 0);
      if (dv >= 200001) yms.add(dv);
    }
    return Array.from(yms).sort((a,b) => b - a).slice(0, 12);
  } catch(e) {
    Logger.log("getAvailableMonths error: " + e);
    return [];
  }
}

// ─ 店舗データ ─
function getStores(targetYM) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const rowsG = ss.getSheetByName(SHEET_GOAL).getDataRange().getValues();
  const rowsR = ss.getSheetByName(SHEET_REAL).getDataRange().getValues();
  const rowsM = ss.getSheetByName(SHEET_MIKOMI).getDataRange().getValues();

  const today = new Date();
  const curYM = targetYM || (today.getFullYear() * 100 + (today.getMonth() + 1));
  const ymYear = Math.floor(curYM / 100);
  const ymMonth = curYM % 100;
  const daysInMonth = new Date(ymYear, ymMonth, 0).getDate();
  // 過去月は月全体（満日）、当月は経過日数
  const isCurrentMonth = curYM === (today.getFullYear() * 100 + (today.getMonth() + 1));
  const elapsed = isCurrentMonth ? today.getDate() - 1 : daysInMonth;

  // 加盟店目標の店舗名セット（正規化キー → 正式名称）
  const goalNameMap = {};
  for (let i = 1; i < rowsG.length; i++) {
    const nm = String(rowsG[i][2] || "").trim();
    if (nm) goalNameMap[_normalizeForMatch(nm)] = nm;
  }
  // 実数値・見込み数値の店舗名を加盟店目標に正規化してマッピング
  const _resolveNm = raw => {
    if (!raw) return raw;
    if (goalNameMap[_normalizeForMatch(raw)]) return goalNameMap[_normalizeForMatch(raw)];
    const normRaw = _normalizeForMatch(raw);
    for (const [normGoal, goalNm] of Object.entries(goalNameMap)) {
      if (normGoal.includes(normRaw) || normRaw.includes(normGoal)) return goalNm;
    }
    return raw;
  };

  const actualMap = {};
  for (let i = 1; i < rowsR.length; i++) {
    const type = String(rowsR[i][0] || "").trim();
    if (type && type !== "加盟") continue;
    const nm = _resolveNm(String(rowsR[i][2] || "").trim());
    if (!nm) continue;
    const dv = Math.round(parseFloat(rowsR[i][3]) || 0);
    if (dv >= 200001 && dv !== curYM) continue;
    actualMap[nm] = rowsR[i];
  }

  const mikomiMap = {};
  for (let i = 1; i < rowsM.length; i++) {
    const nm = _resolveNm(String(rowsM[i][2] || "").trim());
    if (!nm) continue;
    const dv = Math.round(parseFloat(rowsM[i][3]) || 0);
    if (dv >= 200001 && dv !== curYM) continue;
    mikomiMap[nm] = rowsM[i];
  }

  const stores = [];
  for (let i = 1; i < rowsG.length; i++) {
    const g    = rowsG[i];
    const type = String(g[0] || "").trim();
    if (type && type !== "加盟") continue;
    const dv = Math.round(parseFloat(g[3]) || 0);
    if (dv >= 200001 && dv !== curYM) continue;
    const sv = String(g[1] || "").trim();
    const nm = String(g[2] || "").trim();
    if (!nm || !sv) continue;

    const a = actualMap[nm] || [];
    const m = mikomiMap[nm] || [];

    const tgt = pf(g[5]);
    const act = pf(a[5]);
    const pct = tgt > 0 ? round1(act / tgt * 100) : null;

    const mkSales  = elapsed > 0 ? Math.round(act / elapsed * daysInMonth) : null;
    const mkPct    = tgt > 0 && mkSales ? round1(mkSales / tgt * 100) : null;
    const mkTotal  = elapsed > 0 ? Math.round(pf(a[8]) / elapsed * daysInMonth) : 0;
    const mkShinki = elapsed > 0 ? Math.round(pf(a[6]) / elapsed * daysInMonth) : 0;
    const mkRairai = elapsed > 0 ? Math.round(pf(a[7]) / elapsed * daysInMonth) : 0;

    stores.push({
      sv, name: nm,
      売上目標:          pf(g[5]),
      新規目標:          pf(g[6]),
      再来目標:          pf(g[7]),
      総客数目標:        pf(g[8]),
      客単価目標:        pf(g[9]),
      回数券売上目標:    pf(g[14]),
      次回予約率目標:    pf(g[13]),
      ロイヤリティ目標:  pf(g[32]),
      SV売上目標:        pf(g[34]),
      実績売上:          act,
      達成率:            pct,
      新規実績:          pf(a[6]),
      再来実績:          pf(a[7]),
      総客数実績:        pf(a[8]),
      客単価実績:        pf(a[9]),
      施術単価実績:      pf(a[10]),
      新規売上実績:      pf(a[11]),
      再来売上実績:      pf(a[13]),
      回数券売上実績:    pf(a[15]),
      物販売上実績:      pf(a[20]),
      次回予約率実績:    pf(a[30]),
      ロイヤリティ実績:  pf(a[31]),
      SV売上実績:        pf(a[33]),
      見込み新規:        mkShinki,
      見込み再来:        mkRairai,
      見込み総客数:      mkTotal,
      見込み総客数率:    pf(m[9]),
      見込み客単価:      pf(m[10]) || pf(a[9]),
      見込み回数券:      pf(m[12]) || mk(a[15], elapsed, daysInMonth),
      見込み物販:        pf(m[14]) || mk(a[20], elapsed, daysInMonth),
      見込みロイヤリティ: mk(a[31], elapsed, daysInMonth),
      見込みSV売上:       mk(a[33], elapsed, daysInMonth),
      見込み売上:        mkSales,
      見込み達成率:      mkPct,
    });
  }
  return stores;
}

function pf(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function mk(v, elapsed, dim) { const n = pf(v); return (!n || elapsed <= 0) ? 0 : Math.round(n / elapsed * dim); }
function round1(v) { return Math.round(v * 10) / 10; }

function getStaffRanks() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let ws = ss.getSheetByName(SHEET_STAFF);
  if (!ws) {
    ws = ss.insertSheet(SHEET_STAFF);
    ws.getRange(1,1,1,7).setValues([["店舗名","スタッフ名","ランク","点数","指名数","オプション売上","技術面点数"]]);
  }
  const rows = ws.getDataRange().getValues();
  const staffs = [];
  for (let i = 1; i < rows.length; i++) {
    const store   = String(rows[i][0] || "").trim();
    const name    = String(rows[i][1] || "").trim();
    const rank    = String(rows[i][2] || "").trim();
    const score   = pf(rows[i][3]);
    const shimei  = pf(rows[i][4]);
    const option  = pf(rows[i][5]);
    const gijutsu = pf(rows[i][6]);
    if (store && name) staffs.push({ store, name, rank, score, shimei, option, gijutsu, row: i + 1 });
  }
  return staffs;
}

function getStaffSales() {
  const rows = _getSalesRows();
  if (!rows.length) return [];
  // 正式名称 → 略称に変換してSVポータルと名称を統一
  const nameMap = _getSalesNameMap();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const kubun    = String(rows[i][SC_KUBUN] || "").trim();
    const category = String(rows[i][SC_CAT]   || "").trim();
    const menuName = String(rows[i][SC_CAT2]  || "").trim();
    const staffRaw = String(rows[i][SC_STAFF] || "").trim();
    const amt      = parseFloat(rows[i][SC_AMT])   || 0;
    const cnt      = parseFloat(rows[i][SC_COUNT]) || 1;
    const dateNum  = parseFloat(rows[i][SC_DATE])  || 0;
    const storeOfficial = String(rows[i][SC_NAME] || "").trim();
    // 正式名称を略称にマッピング（見つからなければそのまま）
    const store = nameMap.officialToAbbr[storeOfficial] || storeOfficial;
    if (!store || !staffRaw || staffRaw === "フリー 指名なし") continue;
    const month = Math.floor(dateNum / 100) % 100;
    const key = store + "|" + staffRaw;
    if (!map[key]) map[key] = { store, name: staffRaw, sales: {}, bussan: {}, kaisu: {}, option: {}, shimei: {} };
    if (amt > 0) {
      if (kubun === "施術") map[key].sales[month] = (map[key].sales[month] || 0) + amt;
      else if (kubun === "店販" && !category.includes("指名料")) map[key].sales[month] = (map[key].sales[month] || 0) + amt;
      else if (kubun === "その他" && category === "オプション") map[key].sales[month] = (map[key].sales[month] || 0) + amt;
    }
    if (kubun === "店販" && !category.includes("指名料") && !category.includes("回数券") && amt > 0)
      map[key].bussan[month] = (map[key].bussan[month] || 0) + amt;
    if ((kubun === "施術" || kubun === "その他") && (category.includes("回数券") || menuName.includes("回数券")))
      map[key].kaisu[month] = (map[key].kaisu[month] || 0) + cnt;
    if (kubun === "施術" && amt > 0 && (category.includes("オプション") || category.includes("OP") || menuName.includes("オプション") || menuName.includes("OP")))
      map[key].option[month] = (map[key].option[month] || 0) + amt;
    if (kubun === "店販" && category.includes("指名料"))
      map[key].shimei[month] = (map[key].shimei[month] || 0) + 1;
  }
  const result = [];
  Object.values(map).forEach(s => {
    const salesVals = Object.values(s.sales);
    if (!salesVals.length) return;
    const months    = salesVals.length;
    const avgSales  = Math.round(salesVals.reduce((a,b)=>a+b,0) / months);
    const bVals     = Object.values(s.bussan);
    const avgBussan = bVals.length ? Math.round(bVals.reduce((a,b)=>a+b,0) / bVals.length) : 0;
    const kVals     = Object.values(s.kaisu);
    const avgKaisu  = kVals.length ? Math.round(kVals.reduce((a,b)=>a+b,0) / kVals.length) : 0;
    const oVals     = Object.values(s.option);
    const avgOption = oVals.length ? Math.round(oVals.reduce((a,b)=>a+b,0) / oVals.length) : 0;
    const sVals     = Object.values(s.shimei);
    const avgShimei = sVals.length ? Math.round(sVals.reduce((a,b)=>a+b,0) / sVals.length) : 0;
    const pSales   = avgSales   >= 850000 ? 8  : avgSales   >= 730000 ? 5 : 2;
    const pShimei  = avgShimei  >= 30     ? 12 : avgShimei  >= 10     ? 9 : 6;
    const pBussan  = avgBussan  >= 90000  ? 10 : avgBussan  >= 30000  ? 7 : 4;
    const pKaisu   = avgKaisu   >= 11     ? 12 : avgKaisu   >= 5      ? 9 : 6;
    const pOption  = avgOption  >= 45000  ? 8  : avgOption  >= 30000  ? 5 : 2;
    const pGijutsu = 7;
    const total    = pSales + pShimei + pBussan + pKaisu + pOption + pGijutsu;
    const rank     = total >= 48 ? "松" : total >= 36 ? "竹" : "梅";
    result.push({ store: s.store, name: s.name, avg: avgSales, avgBussan, avgKaisu, avgOption, avgShimei, pSales, pShimei, pBussan, pKaisu, pOption, pGijutsu, total, rank, hasManual: true, monthly: s.sales, months });
  });
  return result.sort((a,b) => b.avg - a.avg);
}

function updateStaffRank(storeName, staffName, rank, score) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let ws = ss.getSheetByName(SHEET_STAFF);
  if (!ws) {
    ws = ss.insertSheet(SHEET_STAFF);
    ws.getRange(1,1,1,7).setValues([["店舗名","スタッフ名","ランク","点数","指名数","オプション売上","技術面点数"]]);
  }
  const rows = ws.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]||"").trim()===storeName && String(rows[i][1]||"").trim()===staffName) { targetRow=i+1; break; }
  }
  if (targetRow < 0) {
    ws.appendRow([storeName, staffName, rank||"", score||"", "", "", ""]);
  } else {
    if (rank  !== undefined) ws.getRange(targetRow,3).setValue(rank||"");
    if (score !== undefined) ws.getRange(targetRow,4).setValue(score||"");
  }
}

function deleteStaff(storeName, staffName) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_STAFF);
  if (!ws) return;
  const rows = ws.getDataRange().getValues();
  for (let i = rows.length-1; i >= 1; i--) {
    if (String(rows[i][0]||"").trim()===storeName && String(rows[i][1]||"").trim()===staffName) { ws.deleteRow(i+1); break; }
  }
}

// ─ 店舗追加 ─
// goals: { sales, newGuest, repeat, total, unitPrice, ticket, royalty, wholesale, svSales }
function addStoreFn(sv, name, ym, goals) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_GOAL);
  const row = new Array(35).fill('');
  row[0] = '加盟'; row[1] = sv; row[2] = name; row[3] = String(ym); row[4] = 'TRUE';
  if (goals.sales      !== undefined) row[5]  = Number(goals.sales)      || '';
  if (goals.newGuest   !== undefined) row[6]  = Number(goals.newGuest)   || '';
  if (goals.repeat     !== undefined) row[7]  = Number(goals.repeat)     || '';
  if (goals.total      !== undefined) row[8]  = Number(goals.total)      || '';
  if (goals.unitPrice  !== undefined) row[9]  = Number(goals.unitPrice)  || '';
  if (goals.ticket     !== undefined) row[14] = Number(goals.ticket)     || '';
  if (goals.royalty    !== undefined) row[32] = Number(goals.royalty)    || '';
  if (goals.wholesale  !== undefined) row[33] = Number(goals.wholesale)  || '';
  if (goals.svSales    !== undefined) row[34] = Number(goals.svSales)    || '';
  ws.appendRow(row);
}

// ─ 目標更新（行が無ければ追加） ─
// goals: 同上
function updateGoalFn(name, ym, goals) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_GOAL);
  const rows = ws.getDataRange().getValues();
  const ymStr = String(ym);
  // col番号は1-indexed
  const colMap = { sales:6, newGuest:7, repeat:8, total:9, unitPrice:10,
                   ticket:15, royalty:33, wholesale:34, svSales:35 };
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]||'').trim() === name && String(rows[i][3]||'').trim() === ymStr) {
      for (const [key, col] of Object.entries(colMap)) {
        if (goals[key] !== undefined) ws.getRange(i+1, col).setValue(Number(goals[key]) || '');
      }
      return;
    }
  }
  // 行が存在しない場合は新規追加
  const svRow = rows.find(r => String(r[2]||'').trim() === name);
  const sv = svRow ? String(svRow[1]||'') : '';
  addStoreFn(sv, name, ym, goals);
}

// ─ 加盟店目標シートのIMPORTRANGEを復元 ─
function restoreImportRangeFn() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_GOAL);
  // 現在の値データをクリア（行2以降）
  const lastRow = ws.getLastRow();
  if (lastRow >= 2) {
    ws.getRange(2, 1, lastRow - 1, ws.getLastColumn()).clearContent();
  }
  // A2にIMPORTRANGE数式をセット
  ws.getRange("A2").setFormula(
    '=IMPORTRANGE("1B2eQ8K4oN7DgvTU3-mWF8ZShfDDVPXM8aU6GuxlWwMI","目標!A2:AI1057")'
  );
}

// ─ 分析SSから加盟店目標シートにデータを直接同期 ─
// IMPORTRANGEが壊れた場合の代替。GASがSpreadsheetApp経由で両SSにアクセスし値をコピー
function syncGoalDataFn() {
  const SOURCE_SS_ID = "1B2eQ8K4oN7DgvTU3-mWF8ZShfDDVPXM8aU6GuxlWwMI";
  const SOURCE_SHEET = "目標";

  // 参照元から取得
  const srcSS = SpreadsheetApp.openById(SOURCE_SS_ID);
  const srcWs = srcSS.getSheetByName(SOURCE_SHEET);
  const srcData = srcWs.getDataRange().getValues();
  // ヘッダー行を除いたデータ（行2〜）
  const dataRows = srcData.slice(1); // A2〜の値

  // 参照先の加盟店目標シートのA2セルの数式を削除して値で上書き
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_GOAL);

  // A2のIMPORTRANGE数式を削除
  ws.getRange("A2").clearContent();

  // 既存データ（行2以降）をクリア
  const lastRow = ws.getLastRow();
  if (lastRow >= 2) {
    ws.getRange(2, 1, lastRow - 1, ws.getLastColumn()).clearContent();
  }

  // 分析SSのデータを値として貼り付け
  if (dataRows.length > 0) {
    const cols = dataRows[0].length;
    ws.getRange(2, 1, dataRows.length, cols).setValues(dataRows);
  }

  return dataRows.length;
}

// ─ 加盟店目標シート 歴史データ修復 ─
// 202603行（完全データあり）をテンプレートに、それ以前の年月で
// A/B/C列（直営/加盟・SV・店舗名）が空の行を一括補完する
function repairGoalSheetFn() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_GOAL);
  const rows = ws.getDataRange().getValues();

  // 202603のデータからテンプレート（順序付き）を構築
  const template = [];
  for (let i = 1; i < rows.length; i++) {
    const ym = Math.round(parseFloat(rows[i][3]) || 0);
    if (ym !== 202603) continue;
    const cat = String(rows[i][0] || "").trim();
    const sv  = String(rows[i][1] || "").trim();
    const nm  = String(rows[i][2] || "").trim();
    if (!nm || !sv) continue;
    template.push([cat, sv, nm]);
  }

  // 修復対象の年月ブロックを収集（202603以外で A/B/C が空の行）
  const ymBlocks = {}; // ym -> 行インデックス配列（0-based in rows）
  for (let i = 1; i < rows.length; i++) {
    const ym = Math.round(parseFloat(rows[i][3]) || 0);
    if (ym < 200001 || ym === 202603) continue;
    const sv = String(rows[i][1] || "").trim();
    const nm = String(rows[i][2] || "").trim();
    if (nm && sv) continue; // 既にデータあり
    if (!ymBlocks[ym]) ymBlocks[ym] = [];
    ymBlocks[ym].push(i);
  }

  let repairedRows = 0;
  for (const ym of Object.keys(ymBlocks)) {
    const indices = ymBlocks[ym];
    for (let j = 0; j < indices.length && j < template.length; j++) {
      const sheetRow = indices[j] + 1; // 1-based
      ws.getRange(sheetRow, 1, 1, 3).setValues([template[j]]);
      repairedRows++;
    }
  }
  return repairedRows;
}

// ─ ベストプラクティス読み込み ─
const BP_SS_ID = "1SW1b7hTSD0y2VbLYLN6ndFpXyhWgQrmZ3DKDNd2mexw";

function getBestPractices() {
  const ss = SpreadsheetApp.openById(BP_SS_ID);
  const sheets = ss.getSheets();
  const result = {};
  sheets.forEach(ws => {
    const rows = ws.getDataRange().getValues();
    result[ws.getName()] = rows;
  });
  return result;
}

// ─ アジェンダ生成（JSONで返す） ─
function createAgenda(storeName, format, memo) {
  const stores = getStores();
  const s = stores.find(st => st.name === storeName);
  if (!s) throw new Error("店舗が見つかりません: " + storeName);

  const allTasks = getTasks();
  const storeTasks = allTasks.filter(t => t.store === storeName && t.status !== "完了");

  // 4象限判定（BP基準: 施術単価¥5,020固定で高/低を判定）
  const total = (s.新規実績 || 0) + (s.再来実績 || 0);
  const newRatio = total > 0 ? s.新規実績 / total : 0.5;
  const unitPrice = s.客単価実績 || 0;
  const unitGoal  = s.客単価目標 || 5020;
  const BP_UNIT = 5020;
  const isHighPrice = unitPrice >= BP_UNIT;
  const isNewMajor  = newRatio > 0.5;
  let quadrant = "", quadrantMsg = "";
  if (isHighPrice && isNewMajor)       { quadrant = "優良新規"; quadrantMsg = "高単価×新規中心。VIPへの転換を促す施策が重要。"; }
  else if (isHighPrice && !isNewMajor) { quadrant = "VIP";      quadrantMsg = "高単価×再来中心。最良の状態。維持と口コミ促進を。"; }
  else if (!isHighPrice && isNewMajor) { quadrant = "お試し層"; quadrantMsg = "低単価×新規中心。次回予約率向上が最優先課題。"; }
  else                                 { quadrant = "リピーター"; quadrantMsg = "低単価×再来中心。単価アップの提案強化を。"; }

  const today = new Date();
  const ym = Utilities.formatDate(today, "Asia/Tokyo", "yyyy年M月");
  const dateStr = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd");
  const newPct = Math.round(newRatio * 100);

  return {
    ok: true,
    storeName, ym, dateStr, sv: s.sv, quadrant, quadrantMsg, newPct,
    unitPrice: Math.round(unitPrice),
    unitGoal: Math.round(unitGoal),
    summary: {
      sales: Math.round(s.実績売上 || 0), salesGoal: Math.round(s.売上目標 || 0), salesPct: s.達成率 || 0,
      forecastSales: Math.round(s.見込み売上 || 0), forecastPct: s.見込み達成率 || 0,
      royalty: Math.round(s.ロイヤリティ実績 || 0), royaltyGoal: Math.round(s.ロイヤリティ目標 || 0),
      svSales: Math.round(s.SV売上実績 || 0), svSalesGoal: Math.round(s.SV売上目標 || 0),
    },
    kpi: {
      total: Math.round(s.総客数実績 || 0), totalGoal: Math.round(s.総客数目標 || 0),
      newGuest: Math.round(s.新規実績 || 0), newGoal: Math.round(s.新規目標 || 0),
      repeat: Math.round(s.再来実績 || 0), repeatGoal: Math.round(s.再来目標 || 0),
      unitPrice: Math.round(s.客単価実績 || 0), unitPriceGoal: Math.round(s.客単価目標 || 0),
      nextRes: Math.round((s.次回予約率実績 || 0) * 100),
      ticket: Math.round(s.回数券売上実績 || 0), ticketGoal: Math.round(s.回数券売上目標 || 0),
      bussan: Math.round(s.物販売上実績 || 0),
    },
    tasks: storeTasks.map(t => ({category: t.category, taskName: t.taskName, priority: t.priority, status: t.status, memo: t.memo})),
    memo,
  };
}

// ─ アジェンダ外部ファイル生成（Docs/Slides） ─
function createAgendaExternal(storeName, format, memo) {
  const today = new Date();
  const curYM = today.getFullYear() * 100 + (today.getMonth() + 1);
  const prevDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const prevYM = prevDate.getFullYear() * 100 + (prevDate.getMonth() + 1);
  const prevYMStr = Utilities.formatDate(prevDate, "Asia/Tokyo", "yyyy年M月");
  // 2ヶ月前（比較用）
  const prevPrevDate = new Date(today.getFullYear(), today.getMonth() - 2, 1);
  const prevPrevYM = prevPrevDate.getFullYear() * 100 + (prevPrevDate.getMonth() + 1);
  const prevPrevYMStr = Utilities.formatDate(prevPrevDate, "Asia/Tokyo", "yyyy年M月");

  const stores = getStores(curYM);
  const s = stores.find(st => st.name === storeName);
  if (!s) throw new Error("店舗が見つかりません: " + storeName);

  const prevStores = getStores(prevYM);
  const sPrev = prevStores.find(st => st.name === storeName) || null;

  const prevPrevStores = getStores(prevPrevYM);
  const sPrevPrev = prevPrevStores.find(st => st.name === storeName) || null;

  const tasks = getTasks().filter(t => t.store === storeName && t.status !== "完了");
  const title = storeName + " オーナーMTGアジェンダ " + prevYMStr + "振り返り";
  // 分析は先月データ（sPrev）を優先。なければ当月で代替
  const sA = sPrev || s;
  const total = (sA.新規実績 || 0) + (sA.再来実績 || 0);
  const newRatio = total > 0 ? sA.新規実績 / total : 0.5;
  const unitPrice = sA.客単価実績 || 0;
  const unitGoal  = sA.客単価目標 || 5020;
  const BP_UNIT = 5020;
  const isHighPrice = unitPrice >= BP_UNIT;
  const isNewMajor  = newRatio > 0.5;
  let quadrant = "", quadrantMsg = "";
  if (isHighPrice && isNewMajor)       { quadrant = "優良新規"; quadrantMsg = "高単価×新規中心。VIPへの転換を促す施策が重要。"; }
  else if (isHighPrice && !isNewMajor) { quadrant = "VIP";      quadrantMsg = "高単価×再来中心。最良の状態。維持と口コミ促進を。"; }
  else if (!isHighPrice && isNewMajor) { quadrant = "お試し層"; quadrantMsg = "低単価×新規中心。次回予約率向上が最優先課題。"; }
  else                                 { quadrant = "リピーター"; quadrantMsg = "低単価×再来中心。単価アップの提案強化を。"; }
  // メニュー比率は先月YMで取得
  const menuData = getMenuRatios(storeName, prevYM);
  // 店舗レポート（売上明細から月別推移・スタッフ集計）スタッフは先月のみ
  const storeReport = getStoreReport(storeName, prevYM);
  if (format === "slides") {
    return createAgendaSlides(title, s, sPrev, sPrevPrev, prevYMStr, prevPrevYMStr, tasks, quadrant, quadrantMsg, memo, newRatio, unitPrice, unitGoal, menuData, storeReport);
  } else {
    return createAgendaDoc(title, s, sPrev, sPrevPrev, prevYMStr, prevPrevYMStr, tasks, quadrant, quadrantMsg, memo, newRatio, unitPrice, unitGoal, menuData, storeReport);
  }
}

// ─ メニュー構成比取得 ─
function getMenuRatios(storeName, targetYM) {
  try {
    const rows = _getSalesRows();
    if (!rows.length) return null;
    const counts = { matsu:0, mayu:0, matsuMayu:0, matsuMayuPerm:0, other:0 };
    const amounts = { matsu:0, mayu:0, matsuMayu:0, matsuMayuPerm:0, other:0 };
    // 略称 → 正式名称に変換して売上明細を検索
    const nameMap = _getSalesNameMap();
    const officialName = nameMap.abbrToOfficial[storeName] || storeName;
    for (let i = 1; i < rows.length; i++) {
      const store = String(rows[i][SC_NAME] || "").trim();
      if (store !== officialName) continue;
      const kubun = String(rows[i][SC_KUBUN] || "").trim();
      if (kubun !== "施術") continue;
      if (targetYM) {
        const dn = parseFloat(rows[i][SC_DATE]) || 0;
        if (Math.floor(dn / 100) !== targetYM) continue;
      }
      const mn = (String(rows[i][SC_CAT] || "") + " " + String(rows[i][SC_CAT2] || "")).trim();
      const cnt = parseFloat(rows[i][SC_COUNT]) || 1;
      const amt = parseFloat(rows[i][SC_AMT]) || 0;
      const mnL = mn.toLowerCase();
      const hasMatsu   = mnL.includes("まつ毛パーマ") || mnL.includes("マツパ") || mnL.includes("まつパ")
                      || mnL.includes("パリジェンヌ") || mnL.includes("ラッシュリフト") || mnL.includes("lash lift")
                      || mnL.includes("ラッシュ") || mnL.includes("lash") || mnL.includes("まつげパーマ")
                      || mnL.includes("まつ毛リフト") || mnL.includes("まつげリフト");
      const hasMayuWax = mnL.includes("眉毛ワックス") || mnL.includes("眉ワックス")
                      || mnL.includes("眉毛wax") || mnL.includes("眉wax")
                      || mnL.includes("眉毛wax") || mnL.includes("眉毛wax")
                      || mnL.includes("眉毛wax") || mnL.includes("まゆwax")
                      || (mnL.includes("眉") && (mnL.includes("wax") || mnL.includes("ワックス")));
      const hasMayuPerm= mnL.includes("眉毛パーマ") || mnL.includes("眉パーマ")
                      || mnL.includes("眉毛perm") || mnL.includes("ブロウラミネーション")
                      || mnL.includes("brow lamination");
      let k;
      if      (hasMatsu && hasMayuPerm) k = "matsuMayuPerm";
      else if (hasMatsu && hasMayuWax)  k = "matsuMayu";
      else if (hasMatsu)                k = "matsu";
      else if (hasMayuWax)              k = "mayu";
      else                              k = "other";
      counts[k] += cnt; amounts[k] += amt;
    }
    const total = Object.values(counts).reduce((a,b)=>a+b,0);
    const labels = {
      matsu:"まつ毛パーマ単品", mayu:"眉毛ワックス単品",
      matsuMayu:"まつ毛パーマ+眉毛ワックス", matsuMayuPerm:"まつ毛パーマ+眉毛パーマ", other:"その他"
    };
    return { counts, amounts, total, labels };
  } catch(e) {
    Logger.log("getMenuRatios error: " + e);
    return null;
  }
}

// ─ 実数値シートの店舗名診断（GASエディタから実行） ─
function diagnoseActualSheetNames() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const wsGoal = ss.getSheetByName(SHEET_GOAL);
  const wsReal = ss.getSheetByName(SHEET_REAL);
  if (!wsGoal || !wsReal) { Logger.log("シートが見つかりません"); return; }

  const goalRows = wsGoal.getDataRange().getValues();
  const realRows = wsReal.getDataRange().getValues();

  const officialNames = new Set();
  for (let i = 1; i < goalRows.length; i++) {
    const nm = String(goalRows[i][2] || "").trim();
    if (nm) officialNames.add(nm);
  }

  const realNames = new Set();
  for (let i = 1; i < realRows.length; i++) {
    const nm = String(realRows[i][2] || "").trim();
    if (nm) realNames.add(nm);
  }

  Logger.log("=== 加盟店目標の店舗名 ===");
  Array.from(officialNames).sort().forEach(n => Logger.log(" ○ " + n));

  Logger.log("\n=== 実数値シートで不一致の店舗名 ===");
  const unmatch = Array.from(realNames).filter(n => !officialNames.has(n)).sort();
  if (unmatch.length === 0) {
    Logger.log("  → 全て一致しています");
  } else {
    unmatch.forEach(n => {
      const norm = _normalizeForMatch(n);
      const candidates = Array.from(officialNames).filter(o => {
        const on = _normalizeForMatch(o);
        return on.includes(norm) || norm.includes(on);
      });
      Logger.log(" ✕ 「" + n + "」→ 候補: " + (candidates.length ? candidates.join(" / ") : "なし"));
    });
  }
}

// 実数値シートのC列店舗名を正式名称に一括置換
function fixActualSheetStoreNames(preview) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const wsGoal = ss.getSheetByName(SHEET_GOAL);
  const wsReal = ss.getSheetByName(SHEET_REAL);
  if (!wsGoal || !wsReal) { Logger.log("シートが見つかりません"); return; }

  const goalRows = wsGoal.getDataRange().getValues();
  const officialNames = [];
  for (let i = 1; i < goalRows.length; i++) {
    const nm = String(goalRows[i][2] || "").trim();
    if (nm && !officialNames.includes(nm)) officialNames.push(nm);
  }

  const realData = wsReal.getDataRange().getValues();
  let fixCount = 0;

  for (let i = 1; i < realData.length; i++) {
    const raw = String(realData[i][2] || "").trim();
    if (!raw || officialNames.includes(raw)) continue;

    const norm = _normalizeForMatch(raw);
    let best = null;
    for (const o of officialNames) {
      const on = _normalizeForMatch(o);
      if (on === norm || on.includes(norm) || norm.includes(on)) {
        if (!best || o.length > best.length) best = o;
      }
    }
    if (best) {
      Logger.log("行" + (i+1) + ": 「" + raw + "」→「" + best + "」");
      if (!preview) wsReal.getRange(i + 1, 3).setValue(best);
      fixCount++;
    } else {
      Logger.log("行" + (i+1) + ": 「" + raw + "」→ 候補なし");
    }
  }
  Logger.log((preview ? "[プレビュー] " : "[実行済み] ") + "変更件数: " + fixCount);
}

// ─ 店舗名正規化ユーティリティ ─
function _normalizeForMatch(name) {
  return name.replace(/\s/g, "").replace(/店$/,"").replace(/（.*?）/g,"").replace(/\(.*?\)/g,"").toLowerCase();
}

// 売上明細の店舗名と正式名称の差異を確認（GASエディタから実行→ログ参照）
function diagnoseStoreNames() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const wsGoal = ss.getSheetByName(SHEET_GOAL);
  const wsSales = ss.getSheetByName(SHEET_SALES);
  if (!wsGoal || !wsSales) { Logger.log("シートが見つかりません"); return; }

  // 正式名称一覧（加盟店目標 C列）
  const goalRows = wsGoal.getDataRange().getValues();
  const officialNames = new Set();
  for (let i = 1; i < goalRows.length; i++) {
    const nm = String(goalRows[i][2] || "").trim();
    if (nm) officialNames.add(nm);
  }

  // 売上明細の店舗名一覧（A列）
  const salesRows = wsSales.getDataRange().getValues();
  const salesNames = new Set();
  for (let i = 1; i < salesRows.length; i++) {
    const nm = String(salesRows[i][0] || "").trim();
    if (nm) salesNames.add(nm);
  }

  Logger.log("=== 正式名称（加盟店目標）===");
  Array.from(officialNames).sort().forEach(n => Logger.log(" ○ " + n));

  Logger.log("\n=== 売上明細の店舗名（正式名称と不一致のもの）===");
  const unmatch = Array.from(salesNames).filter(n => !officialNames.has(n)).sort();
  if (unmatch.length === 0) {
    Logger.log("  → 全て一致しています");
  } else {
    unmatch.forEach(n => {
      // 類似候補を探す
      const norm = _normalizeForMatch(n);
      const candidates = Array.from(officialNames).filter(o => {
        const on = _normalizeForMatch(o);
        return on.includes(norm) || norm.includes(on);
      });
      Logger.log(" ✕ 「" + n + "」→ 候補: " + (candidates.length ? candidates.join(" / ") : "なし"));
    });
  }
  Logger.log("\n正式名称にない売上明細の名前: " + unmatch.length + "件");
}

// 売上明細A列の店舗名を、同シート内で最も長い（正式な）名前に統一
// preview=true でログ確認のみ、引数なし or false で実際に書き込み
function fixSalesStoreNames(preview) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const wsSales = ss.getSheetByName(SHEET_SALES);
  if (!wsSales) { Logger.log("売上明細シートが見つかりません"); return; }

  const salesData = wsSales.getDataRange().getValues();

  // ① 売上明細内のユニーク店舗名を収集
  const allNames = [];
  for (let i = 1; i < salesData.length; i++) {
    const nm = String(salesData[i][0] || "").trim();
    if (nm && !allNames.includes(nm)) allNames.push(nm);
  }

  // ② ファジーマッチでグループ化し、最長名を正式名称とする
  const officialMap = {}; // 各名前 → 正式名称
  const processed = new Set();
  for (const name of allNames) {
    if (processed.has(name)) continue;
    const group = [name];
    const normName = _normalizeForMatch(name);
    for (const other of allNames) {
      if (other === name || processed.has(other)) continue;
      const normOther = _normalizeForMatch(other);
      if (normName === normOther || normName.includes(normOther) || normOther.includes(normName)) {
        group.push(other);
      }
    }
    // 最長の名前を正式名称に採用
    const official = group.slice().sort((a, b) => b.length - a.length)[0];
    Logger.log("グループ: [" + group.join(" / ") + "] → 正式名称:「" + official + "」");
    group.forEach(n => { officialMap[n] = official; processed.add(n); });
  }

  // ③ 略称行を正式名称に書き換え
  let fixCount = 0;
  for (let i = 1; i < salesData.length; i++) {
    const raw = String(salesData[i][0] || "").trim();
    if (!raw) continue;
    const official = officialMap[raw];
    if (official && official !== raw) {
      Logger.log("行" + (i+1) + ": 「" + raw + "」→「" + official + "」");
      if (!preview) wsSales.getRange(i + 1, 1).setValue(official);
      fixCount++;
    }
  }

  Logger.log((preview ? "[プレビュー] " : "[実行済み] ") + "変更件数: " + fixCount);
}

// ─ 店舗名マッピング（略称 ↔ 正式名称） ─
// ─ 売上明細 行キャッシュ（1実行につき1回だけシートを読む）─
let _salesRowsCache = null;
function _getSalesRows() {
  if (_salesRowsCache) return _salesRowsCache;
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_SALES);
  _salesRowsCache = ws ? ws.getDataRange().getValues() : [];
  return _salesRowsCache;
}

// SVポータル各シート: 略称  /  売上明細: 正式名称
// 実行ごとにキャッシュを保持
let _salesNameCache = null;
function _getSalesNameMap() {
  if (_salesNameCache) return _salesNameCache;
  const ss = SpreadsheetApp.openById(SS_ID);
  const wsGoal  = ss.getSheetByName(SHEET_GOAL);
  const wsSales = ss.getSheetByName(SHEET_SALES);
  const abbrToOfficial = {};
  const officialToAbbr = {};
  if (!wsGoal) {
    _salesNameCache = { abbrToOfficial, officialToAbbr };
    return _salesNameCache;
  }
  // 略称一覧（加盟店目標 C列）
  const abbrNames = [];
  wsGoal.getDataRange().getValues().slice(1).forEach(r => {
    const nm = String(r[2] || "").trim();
    if (nm && !abbrNames.includes(nm)) abbrNames.push(nm);
  });
  // 正式名称一覧（売上明細 B列 = SC_NAME）- キャッシュ経由
  const officialNames = new Set();
  _getSalesRows().slice(1).forEach(r => {
    const nm = String(r[SC_NAME] || "").trim();
    if (nm) officialNames.add(nm);
  });
  // 略称 → 正式名称マッピング（ファジーマッチ）
  for (const abbr of abbrNames) {
    if (officialNames.has(abbr)) {
      abbrToOfficial[abbr] = abbr;
      officialToAbbr[abbr] = abbr;
      continue;
    }
    const normAbbr = _normalizeForMatch(abbr);
    let best = null;
    for (const official of officialNames) {
      const normOfficial = _normalizeForMatch(official);
      if (normOfficial === normAbbr || normOfficial.includes(normAbbr) || normAbbr.includes(normOfficial)) {
        if (!best || official.length > best.length) best = official;
      }
    }
    if (best) {
      abbrToOfficial[abbr] = best;
      officialToAbbr[best] = abbr;
    }
  }
  _salesNameCache = { abbrToOfficial, officialToAbbr };
  return _salesNameCache;
}

// ─ 課題自動判定（コンサルタント目線） ─
function generateIssues(s, unitPrice, unitGoal, newRatio) {
  const issues = [];
  const fmt = n => Math.round(n || 0).toLocaleString();
  const forecastPct = s.見込み達成率 || 0;
  const newPct = Math.round(newRatio * 100);
  const nextResPct = Math.round((s.次回予約率実績 || 0) * 100);
  const guestPct = s.総客数目標 > 0 ? Math.round((s.総客数実績||0) / s.総客数目標 * 100) : 0;
  const ticketPct = s.回数券売上目標 > 0 ? Math.round((s.回数券売上実績||0) / s.回数券売上目標 * 100) : 0;

  if (forecastPct < 100) {
    issues.push({
      title: "月末着地が目標未達見込み",
      detail: "見込み達成率 " + forecastPct + "%（¥" + fmt(s.見込み売上) + " / 目標 ¥" + fmt(s.売上目標) + "）",
      comment: "残り期間でのフォロー強化と回数券・物販の積み増しが急務。"
    });
  }
  if (nextResPct < 35) {
    issues.push({
      title: "次回予約率が基準（35%）以下",
      detail: "次回予約率 " + nextResPct + "%",
      comment: "施術中のクロージングトークが不足している可能性。退店前に必ず次回日程を提案する習慣をつける。"
    });
  }
  if (newPct > 60) {
    issues.push({
      title: "新規依存度が高くリピート定着が課題",
      detail: "新規比率 " + newPct + "% / 再来比率 " + (100 - newPct) + "%",
      comment: "新規を獲得しても再来につながっていない状態。次回予約・回数券販売でリピート率向上を優先する。"
    });
  }
  if (newPct < 20) {
    issues.push({
      title: "新規流入が不足",
      detail: "新規比率 " + newPct + "%",
      comment: "SNS・口コミ・ホットペッパー露出を見直し、新規集客施策を強化する必要がある。"
    });
  }
  if (guestPct < 80 && s.総客数目標 > 0) {
    issues.push({
      title: "総客数が目標を大きく下回っている",
      detail: "達成率 " + guestPct + "%（" + fmt(s.総客数実績) + "人 / 目標 " + fmt(s.総客数目標) + "人）",
      comment: "新規・再来の両面で客数が不足。集客経路と離脱タイミングの分析が必要。"
    });
  }
  if (unitPrice < unitGoal && unitGoal > 0) {
    issues.push({
      title: "客単価が目標未達",
      detail: "¥" + fmt(unitPrice) + "（目標 ¥" + fmt(unitGoal) + "）",
      comment: "オプション・物販・コースメニューの提案が不十分。メニュー構成比と合わせて提案内容を見直す。"
    });
  }
  if (ticketPct < 70 && s.回数券売上目標 > 0) {
    issues.push({
      title: "回数券販売が低調",
      detail: "達成率 " + ticketPct + "%（¥" + fmt(s.回数券売上実績) + " / 目標 ¥" + fmt(s.回数券売上目標) + "）",
      comment: "LTV向上の最重要施策。施術後の自然な流れでメリットを伝えるトーク設計を見直す。"
    });
  }
  if ((s.物販売上実績 || 0) < 29700) {
    issues.push({
      title: "物販売上がBP基準（¥29,700）未達",
      detail: "物販売上 ¥" + fmt(s.物販売上実績),
      comment: "アフターケア商品の提案機会を増やす。施術内容と連動した商品提案（眉ケア・まつ毛美容液など）が有効。"
    });
  }

  return issues;
}

// ─ ベストプラクティス施策取得 ─
function getRelevantStrategies(quadrant, metrics) {
  try {
    const ss = SpreadsheetApp.openById(BP_SS_ID);
    const ws = ss.getSheetByName("施策一覧表");
    if (!ws) return [];
    const rows = ws.getDataRange().getValues();
    const strategies = [];
    for (let i = 1; i < rows.length; i++) {
      const targetQ = String(rows[i][0] || "").trim();
      const timing  = String(rows[i][1] || "").trim();
      // [2]=番号(空欄) [3]=施策名 [4]=施策対象 [5]=重要度
      const name    = String(rows[i][3] || "").trim();
      const target  = String(rows[i][4] || "").trim();
      const imp     = String(rows[i][5] || "").trim();
      if (!name) continue;
      const isGlobal = targetQ === "全対象";
      const matchQ   = isGlobal || targetQ.includes(quadrant);
      if (!matchQ) continue;
      let stars = (imp.match(/★/g) || []).length;
      // メトリクスベースで優先度を上げる
      if (metrics) {
        if (metrics.nextRes < 0.35 && target.includes("再来")) stars += 1;
        if (metrics.unitPrice < 5000 && target.includes("単価")) stars += 1;
        if (metrics.newGuest < 30 && target.includes("新規")) stars += 1;
      }
      strategies.push({ name, target, importance: imp, stars, timing, forQuadrant: targetQ });
    }
    strategies.sort((a, b) => {
      const aSpec = a.forQuadrant !== "全対象" ? 1 : 0;
      const bSpec = b.forQuadrant !== "全対象" ? 1 : 0;
      if (aSpec !== bSpec) return bSpec - aSpec;
      return b.stars - a.stars;
    });
    return strategies.slice(0, 8);
  } catch(e) {
    return [];
  }
}

function styleTableHeader(table, cols, bgColor) {
  const row = table.getRow(0);
  for (let c = 0; c < cols; c++) {
    row.getCell(c).setBackgroundColor(bgColor);
    row.getCell(c).editAsText().setBold(true).setFontSize(11);
  }
}

function createAgendaDoc(title, s, sPrev, sPrevPrev, prevYMStr, prevPrevYMStr, tasks, quadrant, quadrantMsg, memo, newRatio, unitPrice, unitGoal, menuData, storeReport) {
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  body.clear();

  const today = new Date();
  const dateStr = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd");
  const fmt = n => Math.round(n || 0).toLocaleString();
  const newPct = Math.round(newRatio * 100);
  const pct = (a, b) => b > 0 ? Math.round(a / b * 100) + "%" : "—";
  const sMain = sPrev || s;  // 振り返り対象月（先月）

  // タイトル
  const titlePara = body.appendParagraph(title);
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titlePara.editAsText().setForegroundColor("#1e40af").setFontSize(18).setBold(true);
  const metaPara = body.appendParagraph("担当SV: " + s.sv + "　作成日: " + dateStr + "　象限: 【" + quadrant + "】");
  metaPara.editAsText().setFontSize(11).setForegroundColor("#64748b");
  body.appendParagraph("");

  const trend = (cur, prev) => {
    if (cur == null || prev == null || prev === 0) return "—";
    const diff = cur - prev;
    const pct2 = Math.round(Math.abs(diff) / prev * 100);
    return diff > 0 ? "▲" + pct2 + "%" : diff < 0 ? "▼" + pct2 + "%" : "→";
  };

  // 1. 数値サマリー（振り返り月 vs 比較月）
  const h1 = body.appendParagraph("1. 数値サマリー（" + prevYMStr + " 振り返り）");
  h1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h1.editAsText().setForegroundColor("#1e40af").setBold(true);
  const prevPrevLabel = sPrevPrev ? prevPrevYMStr + "実績" : "2ヶ月前";
  const nextResPct = Math.round((sMain.次回予約率実績 || 0) * 100);
  const nextResAlert = nextResPct < 35 ? " ⚠" : " ✓";
  const unitAlert = unitPrice < 5020 ? " ⚠" : " ✓";
  const t1 = body.appendTable([
    ["指標", prevYMStr+"実績", "目標", "達成率", prevPrevLabel, "前月比"],
    ["売上",       "¥"+fmt(sMain.実績売上),       "¥"+fmt(sMain.売上目標),       (sMain.達成率||0)+"%",                           sPrevPrev?"¥"+fmt(sPrevPrev.実績売上):"—",         trend(sMain.実績売上, sPrevPrev?.実績売上)],
    ["総客数",     fmt(sMain.総客数実績)+"人",     fmt(sMain.総客数目標)+"人",    pct(sMain.総客数実績,sMain.総客数目標)+"%",       sPrevPrev?fmt(sPrevPrev.総客数実績)+"人":"—",      trend(sMain.総客数実績, sPrevPrev?.総客数実績)],
    ["新規客数",   fmt(sMain.新規実績)+"人",       fmt(sMain.新規目標)+"人",      pct(sMain.新規実績,sMain.新規目標)+"%",           sPrevPrev?fmt(sPrevPrev.新規実績)+"人":"—",        trend(sMain.新規実績, sPrevPrev?.新規実績)],
    ["再来客数",   fmt(sMain.再来実績)+"人",       fmt(sMain.再来目標)+"人",      pct(sMain.再来実績,sMain.再来目標)+"%",           sPrevPrev?fmt(sPrevPrev.再来実績)+"人":"—",        trend(sMain.再来実績, sPrevPrev?.再来実績)],
    ["客単価",     "¥"+fmt(sMain.客単価実績)+unitAlert, "¥"+fmt(sMain.客単価目標), pct(sMain.客単価実績,sMain.客単価目標)+"%",     sPrevPrev?"¥"+fmt(sPrevPrev.客単価実績):"—",       trend(sMain.客単価実績, sPrevPrev?.客単価実績)],
    ["次回予約率", nextResPct+"%"+nextResAlert,    "35%以上",                     "—",                                             sPrevPrev?Math.round((sPrevPrev.次回予約率実績||0)*100)+"%":"—", "—"],
    ["回数券売上", "¥"+fmt(sMain.回数券売上実績), "¥"+fmt(sMain.回数券売上目標), pct(sMain.回数券売上実績,sMain.回数券売上目標)+"%", sPrevPrev?"¥"+fmt(sPrevPrev.回数券売上実績):"—",  trend(sMain.回数券売上実績, sPrevPrev?.回数券売上実績)],
    ["物販売上",   "¥"+fmt(sMain.物販売上実績),   "BP:¥29,700",                  "—",                                             sPrevPrev?"¥"+fmt(sPrevPrev.物販売上実績):"—",     trend(sMain.物販売上実績, sPrevPrev?.物販売上実績)],
  ]);
  styleTableHeader(t1, 6, "#dbeafe");
  body.appendParagraph("");

  // 2. 顧客象限分析
  const h3 = body.appendParagraph("2. 顧客象限分析");
  h3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h3.editAsText().setForegroundColor("#1e40af").setBold(true);
  const qPara = body.appendParagraph("現在のポジション: 【" + quadrant + "】");
  qPara.editAsText().setFontSize(13).setBold(true).setForegroundColor("#1d4ed8");
  body.appendParagraph(quadrantMsg).editAsText().setFontSize(11);
  const t3 = body.appendTable([
    ["指標", "現状", "BP基準"],
    ["新規比率", newPct + "% (再来" + (100 - newPct) + "%)", "—"],
    ["客単価", (unitPrice >= unitGoal ? "▲ 目標達成 " : "▼ 目標未達 ") + "¥" + fmt(unitPrice) + " / 目標¥" + fmt(unitGoal), "¥5,020"],
    ["推奨フェーズ", quadrant, "VIP または 優良新規"],
  ]);
  styleTableHeader(t3, 3, "#ede9fe");
  body.appendParagraph("集客サイクル: お試し層 → リピーター → VIP → 優良新規 → 増員 → 循環").editAsText().setFontSize(10).setForegroundColor("#64748b");
  body.appendParagraph("");

  // 3. メニュー構成比
  const h4m = body.appendParagraph("3. メニュー構成比");
  h4m.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h4m.editAsText().setForegroundColor("#1e40af").setBold(true);
  if (menuData && menuData.total > 0) {
    const mKeys = ["matsu","mayu","matsuMayu","matsuMayuPerm","other"];
    const menuRows = [["メニュー", "件数", "比率", "売上金額"]];
    mKeys.forEach(k => {
      const cnt = menuData.counts[k];
      if (cnt === 0) return;
      const ratio = Math.round(cnt / menuData.total * 100) + "%";
      const amt = "¥" + Math.round(menuData.amounts[k]).toLocaleString();
      menuRows.push([menuData.labels[k], cnt + "件", ratio, amt]);
    });
    const totalAmt = "¥" + Math.round(Object.values(menuData.amounts).reduce((a,b)=>a+b,0)).toLocaleString();
    menuRows.push(["合計", menuData.total + "件", "100%", totalAmt]);
    const tm = body.appendTable(menuRows);
    styleTableHeader(tm, 4, "#dcfce7");
  } else {
    body.appendParagraph("メニューデータなし").editAsText().setFontSize(11);
  }
  body.appendParagraph("");

  // 4. 課題分析
  const h4i = body.appendParagraph("4. 課題分析");
  h4i.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h4i.editAsText().setForegroundColor("#1e40af").setBold(true);
  const issues = generateIssues(sPrev || s, unitPrice, unitGoal, newRatio);
  if (issues.length === 0) {
    body.appendParagraph("現時点で大きな課題は検出されませんでした。").editAsText().setFontSize(11);
  } else {
    const issueRows = [["課題", "現状", "コンサルコメント"]];
    issues.forEach(iss => issueRows.push([iss.title, iss.detail, iss.comment]));
    const ti = body.appendTable(issueRows);
    styleTableHeader(ti, 3, "#fee2e2");
  }
  body.appendParagraph("");

  // 5. 推奨アクション（ベストプラクティスより）
  const h4 = body.appendParagraph("5. 推奨アクション（ベストプラクティスより）");
  h4.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h4.editAsText().setForegroundColor("#1e40af").setBold(true);
  const strategies = getRelevantStrategies(quadrant, { nextRes: sMain.次回予約率実績 || 0, unitPrice, newGuest: sMain.新規実績 || 0 });
  if (strategies.length === 0) {
    body.appendParagraph("施策データなし");
  } else {
    const stratRows = [["施策名", "対象KPI", "重要度", "推奨タイミング"]];
    strategies.forEach(st => stratRows.push([st.name, st.target || "—", st.importance || "—", st.timing || "—"]));
    const t4 = body.appendTable(stratRows);
    styleTableHeader(t4, 4, "#fef9c3");
  }
  body.appendParagraph("");

  // 6. その他
  const h6 = body.appendParagraph("6. その他");
  h6.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h6.editAsText().setForegroundColor("#1e40af").setBold(true);
  body.appendParagraph(memo || "（なし）").editAsText().setFontSize(11);

  doc.saveAndClose();
  // 他アカウントからも閲覧できるようリンク共有を設定
  try {
    DriveApp.getFileById(doc.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  } catch(shareErr) { Logger.log("doc share error: " + shareErr); }
  return doc.getUrl();
}

function createAgendaSlides(title, s, sPrev, sPrevPrev, prevYMStr, prevPrevYMStr, tasks, quadrant, quadrantMsg, memo, newRatio, unitPrice, unitGoal, menuData, storeReport) {
  const pres = SlidesApp.create(title);
  const BG_DARK = "#0f172a";
  const BG_LIGHT = "#f8fafc";
  const ACCENT = "#2563eb";
  const fmt = n => Math.round(n || 0).toLocaleString();
  const newPct = Math.round(newRatio * 100);
  const dateStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  const pct = (a, g) => g > 0 ? Math.min(Math.round(a / g * 100), 999) : 0;

  function clearSlide(sl) { sl.getPageElements().forEach(el => el.remove()); }
  function addBox(sl, text, x, y, w, h, fontSize, bold, color) {
    const tb = sl.insertTextBox(text || " ", x, y, w, h);
    try {
      const runs = tb.getText().getRuns();
      if (runs.length > 0) {
        const ts = runs[0].getTextStyle();
        ts.setFontSize(fontSize || 14);
        if (bold) ts.setBold(true);
        if (color) ts.setForegroundColor(color);
      }
    } catch(e) { Logger.log("addBox style err: " + e); }
    return tb;
  }
  function setBg(sl, color) { try { sl.getBackground().setSolidFill(color); } catch(e) {} }

  // テーブル挿入ヘルパー
  function addTable(sl, rows, left, top, width, height, headerColor, fs) {
    const nr = rows.length, nc = rows[0].length;
    const tbl = sl.insertTable(nr, nc, left, top, width, height);
    rows.forEach((row, ri) => {
      row.forEach((txt, ci) => {
        const cell = tbl.getCell(ri, ci);
        const cellText = cell.getText();
        const content = String(txt != null ? txt : "");
        cellText.setText(content || " ");
        try {
          const ts = cellText.getTextStyle();
          ts.setFontSize(fs || 10);
          if (ri === 0) ts.setBold(true);
        } catch(e) {}
        if (ri === 0) cell.getFill().setSolidFill(headerColor || "#dbeafe");
      });
    });
    return tbl;
  }

  // 棒グラフ（達成率）挿入ヘルパー
  function addChart(sl, labels, values, chartTitle, left, top, w, h) {
    try {
      const dt = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, "KPI")
        .addColumn(Charts.ColumnType.NUMBER, "達成率(%)");
      labels.forEach((l, i) => dt.addRow([l, values[i] || 0]));
      const chart = Charts.newBarChart()
        .setDataTable(dt.build())
        .setTitle(chartTitle || "")
        .setColors(["#2563eb"])
        .setDimensions(w, h)
        .setOption("hAxis.minValue", 0)
        .setOption("hAxis.maxValue", 150)
        .setOption("legend.position", "none")
        .build();
      sl.insertImage(chart.getAs("image/png"), left, top, w, h);
    } catch(chartErr) { Logger.log("chart error: " + chartErr); }
  }

  const tr = (cur, prev) => {
    if (cur == null || prev == null || prev === 0) return "—";
    const d = cur - prev;
    const p = Math.round(Math.abs(d) / prev * 100);
    return d > 0 ? "▲" + p + "%" : d < 0 ? "▼" + p + "%" : "→";
  };
  const sMain = sPrev || s;  // 振り返り対象月（先月）
  const prevPrevL = sPrevPrev ? prevPrevYMStr : "2ヶ月前";

  // スライドサイズ取得（720×405 標準）
  const PW = pres.getPageWidth();
  const PH = pres.getPageHeight();
  const CW = PW - 30;  // 両端15ptマージン
  const TH = 38;       // タイトル高さ
  const TY = 10;       // タイトルY
  const CY = TY + TH + 5; // コンテンツ開始Y
  const CH = PH - CY - 10; // コンテンツ高さ

  // スライド1: タイトル
  const sl1 = pres.getSlides()[0];
  clearSlide(sl1);
  setBg(sl1, BG_DARK);
  addBox(sl1, title, 20, Math.round(PH*0.25), CW, Math.round(PH*0.22), 22, true, "#ffffff");
  addBox(sl1, "担当SV: " + s.sv + "　" + dateStr + "　象限:【" + quadrant + "】", 20, Math.round(PH*0.52), CW, 35, 14, false, "#94a3b8");
  addBox(sl1, quadrantMsg, 20, Math.round(PH*0.63), CW, 40, 12, false, "#60a5fa");

  // スライド2: 数値サマリー（テーブル左 + 達成率グラフ右）
  const sl2 = pres.appendSlide();
  clearSlide(sl2);
  setBg(sl2, BG_LIGHT);
  addBox(sl2, "1. 数値サマリー（" + prevYMStr + " 振り返り）", 15, TY, Math.round(CW*0.65), TH, 14, true, ACCENT);
  addBox(sl2, "比較: " + prevPrevL, Math.round(CW*0.67), TY, Math.round(CW*0.33), TH, 10, false, "#64748b");
  const salesPct = sMain.達成率 || 0;
  const nextResPct = Math.round((sMain.次回予約率実績 || 0) * 100);
  const prevPrevNextRes = sPrevPrev ? Math.round((sPrevPrev.次回予約率実績||0)*100) : null;
  const kpiRows = [
    ["指標", prevYMStr+"実績", "目標", "達成率", prevPrevL, "前月比"],
    ["売上",      "¥"+fmt(sMain.実績売上),       "¥"+fmt(sMain.売上目標),      salesPct+"%",                                sPrevPrev?"¥"+fmt(sPrevPrev.実績売上):"—",       tr(sMain.実績売上, sPrevPrev?.実績売上)],
    ["総客数",    fmt(sMain.総客数実績)+"人",     fmt(sMain.総客数目標)+"人",   pct(sMain.総客数実績,sMain.総客数目標)+"%",  sPrevPrev?fmt(sPrevPrev.総客数実績)+"人":"—",    tr(sMain.総客数実績, sPrevPrev?.総客数実績)],
    ["新規",      fmt(sMain.新規実績)+"人",       fmt(sMain.新規目標)+"人",     pct(sMain.新規実績,sMain.新規目標)+"%",      sPrevPrev?fmt(sPrevPrev.新規実績)+"人":"—",      tr(sMain.新規実績, sPrevPrev?.新規実績)],
    ["再来",      fmt(sMain.再来実績)+"人",       fmt(sMain.再来目標)+"人",     pct(sMain.再来実績,sMain.再来目標)+"%",      sPrevPrev?fmt(sPrevPrev.再来実績)+"人":"—",      tr(sMain.再来実績, sPrevPrev?.再来実績)],
    ["客単価",    "¥"+fmt(sMain.客単価実績),     "¥"+fmt(sMain.客単価目標),   pct(sMain.客単価実績,sMain.客単価目標)+"%",  sPrevPrev?"¥"+fmt(sPrevPrev.客単価実績):"—",    tr(sMain.客単価実績, sPrevPrev?.客単価実績)],
    ["次回予約率",nextResPct+"%" + (nextResPct<35?" ⚠":""), "35%以上", "—",                                                  prevPrevNextRes!==null?prevPrevNextRes+"%":"—",   "—"],
    ["回数券",    "¥"+fmt(sMain.回数券売上実績), "¥"+fmt(sMain.回数券売上目標),pct(sMain.回数券売上実績,sMain.回数券売上目標)+"%", sPrevPrev?"¥"+fmt(sPrevPrev.回数券売上実績):"—", tr(sMain.回数券売上実績, sPrevPrev?.回数券売上実績)],
    ["物販",      "¥"+fmt(sMain.物販売上実績),   "BP:¥29,700",                 "—",                                          sPrevPrev?"¥"+fmt(sPrevPrev.物販売上実績):"—",   tr(sMain.物販売上実績, sPrevPrev?.物販売上実績)],
  ];
  addTable(sl2, kpiRows, 15, CY, CW, CH, "#dbeafe", 8);

  // スライド3: メニュー構成比
  const sl4m = pres.appendSlide();
  clearSlide(sl4m);
  setBg(sl4m, BG_LIGHT);
  addBox(sl4m, "2. メニュー構成比", 15, TY, CW, TH, 15, true, ACCENT);
  if (menuData && menuData.total > 0) {
    const pieW = Math.round(CW * 0.46);
    const tableW = CW - pieW - 5;
    try {
      const pieDt = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, "メニュー")
        .addColumn(Charts.ColumnType.NUMBER, "件数");
      const pieKeys = ["matsu","mayu","matsuMayu","matsuMayuPerm"];
      pieKeys.forEach(k => {
        if (menuData.counts[k] > 0) pieDt.addRow([menuData.labels[k], menuData.counts[k]]);
      });
      if (menuData.counts.other > 0) pieDt.addRow(["その他", menuData.counts.other]);
      const pieChart = Charts.newPieChart()
        .setDataTable(pieDt.build())
        .setTitle("メニュー構成比")
        .setDimensions(pieW, CH)
        .setOption("legend.position", "bottom")
        .build();
      sl4m.insertImage(pieChart.getAs("image/png"), 15, CY, pieW, CH);
    } catch(pe) { Logger.log("pie error: " + pe); }
    const mKeys = ["matsu","mayu","matsuMayu","matsuMayuPerm","other"];
    const mRows = [["メニュー", "件数", "比率", "売上"]];
    mKeys.forEach(k => {
      const cnt = menuData.counts[k];
      if (cnt === 0) return;
      const ratio = menuData.total > 0 ? Math.round(cnt / menuData.total * 100) + "%" : "—";
      const amt = Math.round(menuData.amounts[k] / 10000) + "万円";
      mRows.push([menuData.labels[k], cnt + "件", ratio, amt]);
    });
    mRows.push(["合計", menuData.total + "件", "100%", Math.round(Object.values(menuData.amounts).reduce((a,b)=>a+b,0)/10000) + "万円"]);
    addTable(sl4m, mRows, 15 + pieW + 5, CY, tableW, CH, "#dbeafe", 9);
  } else {
    addBox(sl4m, "メニューデータなし（売上明細シートに施術データが必要です）", 15, CY, CW, 50, 12, false, "#64748b");
  }

  // スライド4: 顧客象限
  const sl4 = pres.appendSlide();
  clearSlide(sl4);
  setBg(sl4, BG_DARK);
  addBox(sl4, "3. 顧客象限分析", 15, TY, CW, TH, 17, true, "#60a5fa");
  addBox(sl4, "現在のポジション: 【" + quadrant + "】", 15, CY, CW, 42, 20, true, "#fbbf24");
  addBox(sl4, quadrantMsg, 15, CY+47, CW, 38, 13, false, "#e2e8f0");
  addBox(sl4,
    "新規比率: " + newPct + "% / 再来: " + (100 - newPct) + "%\n" +
    "客単価: ¥" + fmt(unitPrice) + "  （" + (unitPrice >= unitGoal ? "▲ 目標達成" : "▼ 目標未達") + "）\n\n" +
    "集客サイクル: お試し層 → リピーター → VIP → 優良新規 → 増員 → 循環",
    15, CY+92, CW, CH-92, 13, false, "#94a3b8");

  // スライド5: 課題分析
  const slIssue = pres.appendSlide();
  clearSlide(slIssue);
  setBg(slIssue, BG_LIGHT);
  addBox(slIssue, "4. 課題分析", 15, TY, CW, TH, 15, true, ACCENT);
  const issues = generateIssues(sPrev || s, unitPrice, unitGoal, newRatio);
  if (issues.length === 0) {
    addBox(slIssue, "現時点で大きな課題は検出されませんでした。", 15, CY, CW, 45, 12, false, "#64748b");
  } else {
    const issueRows = [["課題", "現状", "コンサルコメント"]];
    issues.forEach(iss => issueRows.push([iss.title, iss.detail, iss.comment]));
    addTable(slIssue, issueRows, 15, CY, CW, CH, "#fee2e2", 9);
  }

  // スライド6: 推奨アクション
  const sl5 = pres.appendSlide();
  clearSlide(sl5);
  setBg(sl5, BG_LIGHT);
  addBox(sl5, "5. 推奨アクション（ベストプラクティスより）", 15, TY, CW, TH, 14, true, ACCENT);
  const strategies = getRelevantStrategies(quadrant, { nextRes: sMain.次回予約率実績 || 0, unitPrice, newGuest: sMain.新規実績 || 0 });
  if (strategies.length > 0) {
    const stratRows = [["施策名", "対象KPI", "重要度", "推奨タイミング"]];
    strategies.slice(0, 8).forEach(st => stratRows.push([st.name||"", st.target||"—", st.importance||"—", st.timing||"—"]));
    addTable(sl5, stratRows, 15, CY, CW, CH, "#fef9c3", 9);
  } else {
    addBox(sl5, "施策データなし", 15, CY, CW, 45, 12, false, "#64748b");
  }

  // スライド7: その他
  const sl7 = pres.appendSlide();
  clearSlide(sl7);
  setBg(sl7, BG_DARK);
  addBox(sl7, "6. その他", 15, TY, CW, TH, 16, true, "#60a5fa");
  addBox(sl7, memo || "（なし）", 15, CY, CW, CH, 13, false, "#e2e8f0");

  pres.saveAndClose();
  // 他アカウントからも閲覧できるようリンク共有を設定
  try {
    DriveApp.getFileById(pres.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  } catch(shareErr) { Logger.log("slides share error: " + shareErr); }
  return pres.getUrl();
}

// =============================================
// 店舗レポート
// =============================================

/**
 * 指定店舗の売上明細から月別・メニュー・スタッフの集計を返す
 */
function getStoreReport(storeName, staffYM) {
  // staffYM: スタッフ集計を絞り込む月（省略時は全月累計）
  if (!storeName) return { error: "店舗名が必要です" };

  const nameMap = _getSalesNameMap();
  const officialName = nameMap.abbrToOfficial[storeName] || storeName;

  const rows = _getSalesRows();
  if (!rows.length) return { error: "売上明細シートなし" };

  const monthlyMap = {};
  const menuMap = {};
  const staffMap = {};
  const kubunMap = {};   // 区分ごとの集計（クーポン・割引含む）

  for (let i = 1; i < rows.length; i++) {
    const store = String(rows[i][SC_NAME] || "").trim();
    if (store !== officialName) continue;

    const dn = parseFloat(rows[i][SC_DATE]) || 0;
    const ym = Math.floor(dn / 100);
    if (!ym || ym < 200001) continue;

    const kubun    = String(rows[i][SC_KUBUN] || "").trim();
    const category = String(rows[i][SC_CAT]   || "").trim();
    const amount   = parseFloat(rows[i][SC_AMT])   || 0;
    const staff    = String(rows[i][SC_STAFF] || "").trim();
    const cnt      = parseFloat(rows[i][SC_COUNT]) || 1;

    // 区分ごとに集計（全種類）
    const rawKubun = kubun || String(rows[i][SC_KUBUN] || "").trim();
    if (rawKubun) {
      if (!kubunMap[rawKubun]) kubunMap[rawKubun] = { count: 0, sales: 0 };
      kubunMap[rawKubun].count += cnt;
      kubunMap[rawKubun].sales += amount;
    }

    if (!monthlyMap[ym]) monthlyMap[ym] = { sales: 0, serviceCount: 0 };

    if (kubun === "施術") {
      monthlyMap[ym].sales += amount;
      monthlyMap[ym].serviceCount += cnt;

      if (category) {
        if (!menuMap[category]) menuMap[category] = { count: 0, sales: 0 };
        menuMap[category].count += cnt;
        menuMap[category].sales += amount;
      }
      // スタッフは staffYM 指定時はその月のみ、省略時は全月
      if (staff && (!staffYM || ym === staffYM)) {
        if (!staffMap[staff]) staffMap[staff] = { count: 0, sales: 0 };
        staffMap[staff].count += cnt;
        staffMap[staff].sales += amount;
      }
    }
  }

  const months = Object.keys(monthlyMap).map(Number).sort();
  const monthly = months.map(ym => ({
    ym,
    label: (ym % 100) + "月",
    sales: monthlyMap[ym].sales,
    serviceCount: monthlyMap[ym].serviceCount,
    unitPrice: monthlyMap[ym].serviceCount > 0
      ? Math.round(monthlyMap[ym].sales / monthlyMap[ym].serviceCount) : 0
  }));

  const menus = Object.entries(menuMap)
    .sort((a, b) => b[1].sales - a[1].sales)
    .map(([name, d]) => ({ name, count: d.count, sales: d.sales }));

  const staff = Object.entries(staffMap)
    .sort((a, b) => b[1].sales - a[1].sales)
    .slice(0, 15)
    .map(([name, d]) => ({ name, count: d.count, sales: d.sales }));

  // クーポン・割引など施術以外の区分をまとめる
  const kubunSummary = Object.entries(kubunMap)
    .sort((a, b) => b[1].count - a[1].count)
    .map(([kubun, d]) => ({ kubun, count: d.count, sales: d.sales }));

  return { storeName, officialName, monthly, menus, staff, kubunSummary };
}

// =============================================
// 売上明細 略称列挿入（一回だけ実行）
// =============================================

/**
 * 売上明細シートの一番左に略称列(A)を挿入する。
 * 既に挿入済みの場合は何もしない。
 * GASエディタまたは ?action=addAbbrColumn から実行。
 */
function addAbbrColumnToSales() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_SALES);
  if (!ws) throw new Error("売上明細シートなし");

  // 既に挿入済みかチェック
  const headerA = String(ws.getRange(1, 1).getValue()).trim();
  if (headerA === "略称") {
    Logger.log("略称列は既に挿入済みです");
    return "already_done";
  }

  // キャッシュをリセットして正式名称→略称マップを構築（挿入前の旧構造で）
  _salesRowsCache = null;
  _salesNameCache = null;
  const nameMap = _getSalesNameMap();

  // 列Aの前に新列を挿入してヘッダー設定
  ws.insertColumnBefore(1);
  ws.getRange(1, 1).setValue("略称");

  // B列（挿入後、旧A列=正式名称）を読んで略称をA列に書き込む
  const lastRow = ws.getLastRow();
  if (lastRow >= 2) {
    const officials = ws.getRange(2, 2, lastRow - 1, 1).getValues();
    const abbrs = officials.map(([nm]) => {
      const official = String(nm || "").trim();
      return [nameMap.officialToAbbr[official] || official];
    });
    ws.getRange(2, 1, abbrs.length, 1).setValues(abbrs);
  }

  // キャッシュをリセット（新構造で再読み込みされるように）
  _salesRowsCache = null;
  _salesNameCache = null;

  Logger.log("略称列の挿入完了: " + (lastRow - 1) + "行");
  return "done";
}

// =============================================
// 売上情報CSV 自動同期
// =============================================

const SYNC_SOURCE_SS_ID  = "1B2eQ8K4oN7DgvTU3-mWF8ZShfDDVPXM8aU6GuxlWwMI";
const SYNC_SOURCE_GID    = 50056376;
const SYNC_DEST_GID      = 2067713084;  // sv-portal SS内の同期先シート

/**
 * gidでシートを取得するヘルパー
 */
function _getSheetByGid(ss, gid) {
  return ss.getSheets().find(s => s.getSheetId() === gid) || null;
}

/**
 * 売上情報CSVシートの内容をsv-portalシートに全量コピー
 * GASエディタまたは時間トリガーから実行する
 */
function syncSalesData() {
  const srcSS  = SpreadsheetApp.openById(SYNC_SOURCE_SS_ID);
  const srcSht = _getSheetByGid(srcSS, SYNC_SOURCE_GID);
  if (!srcSht) throw new Error("同期元シートが見つかりません (gid=" + SYNC_SOURCE_GID + ")");

  const destSS  = SpreadsheetApp.openById(SS_ID);
  const destSht = _getSheetByGid(destSS, SYNC_DEST_GID);
  if (!destSht) throw new Error("同期先シートが見つかりません (gid=" + SYNC_DEST_GID + ")");

  const srcData = srcSht.getDataRange().getValues();
  if (!srcData || srcData.length === 0) {
    Logger.log("同期元にデータがありません");
    return;
  }

  const numRows = srcData.length;
  const numCols = srcData[0].length;

  // シートのサイズをデータに合わせて縮小（セル上限対策）
  const curRows = destSht.getMaxRows();
  const curCols = destSht.getMaxColumns();
  if (curRows > numRows) destSht.deleteRows(numRows + 1, curRows - numRows);
  if (curCols > numCols) destSht.deleteColumns(numCols + 1, curCols - numCols);

  // 行・列が足りなければ追加
  if (destSht.getMaxRows() < numRows) destSht.insertRowsAfter(destSht.getMaxRows(), numRows - destSht.getMaxRows());
  if (destSht.getMaxColumns() < numCols) destSht.insertColumnsAfter(destSht.getMaxColumns(), numCols - destSht.getMaxColumns());

  // クリアして書き込み
  destSht.clearContents();
  destSht.getRange(1, 1, numRows, numCols).setValues(srcData);

  Logger.log("同期完了: " + numRows + "行 × " + numCols + "列");
}

/**
 * 毎日自動実行するトリガーを登録（GASエディタで1回だけ実行する）
 * 既存トリガーが重複しないようチェックしてから登録
 */
function setupSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const alreadySet = triggers.some(t => t.getHandlerFunction() === "syncSalesData");
  if (alreadySet) {
    Logger.log("トリガーは既に登録済みです");
    return;
  }
  ScriptApp.newTrigger("syncSalesData")
    .timeBased()
    .everyDays(1)
    .atHour(6)  // 毎朝6時に実行
    .create();
  Logger.log("毎日6時に syncSalesData を実行するトリガーを登録しました");
}

// ─ 権限認証用（GASエディタで1回だけ実行する） ─
function authorizeDocAccess() {
  const doc = DocumentApp.create("_権限テスト_削除してください");
  doc.saveAndClose();
  const pres = SlidesApp.create("_権限テスト_削除してください");
  pres.saveAndClose();
  Logger.log("認証完了。Googleドライブに「_権限テスト_削除してください」が2件できているので手動で削除してください。");
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
