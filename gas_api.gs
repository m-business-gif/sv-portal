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
    // ベストプラクティス読み込み（デバッグ用）
    if (e && e.parameter && e.parameter.action === "getBestPractices") {
      return json(getBestPractices());
    }
    // 指定月の店舗データ取得
    if (e && e.parameter && e.parameter.action === "getStores") {
      const ym = parseInt(e.parameter.ym) || null;
      return json({ stores: getStores(ym) });
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

  const actualMap = {};
  for (let i = 1; i < rowsR.length; i++) {
    const type = String(rowsR[i][0] || "").trim();
    if (type && type !== "加盟") continue;
    const nm = String(rowsR[i][2] || "").trim();
    if (!nm) continue;
    const dv = Math.round(parseFloat(rowsR[i][3]) || 0);
    if (dv >= 200001 && dv !== curYM) continue;
    actualMap[nm] = rowsR[i];
  }

  const mikomiMap = {};
  for (let i = 1; i < rowsM.length; i++) {
    const nm = String(rowsM[i][2] || "").trim();
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
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName(SHEET_SALES);
  if (!ws) return [];
  const rows = ws.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const kubun    = String(rows[i][5] || "").trim();
    const category = String(rows[i][7] || "").trim();
    const menuName = String(rows[i][8] || "").trim();
    const staffRaw = String(rows[i][13] || "").trim();
    const amt      = parseFloat(rows[i][12]) || 0;
    const cnt      = parseFloat(rows[i][11]) || 1;
    const dateNum  = parseFloat(rows[i][1]) || 0;
    const store    = String(rows[i][0] || "").trim();
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

  const stores = getStores(curYM);
  const s = stores.find(st => st.name === storeName);
  if (!s) throw new Error("店舗が見つかりません: " + storeName);

  const prevStores = getStores(prevYM);
  const sPrev = prevStores.find(st => st.name === storeName) || null;

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
  if (format === "slides") {
    return createAgendaSlides(title, s, sPrev, prevYMStr, tasks, quadrant, quadrantMsg, memo, newRatio, unitPrice, unitGoal, menuData);
  } else {
    return createAgendaDoc(title, s, sPrev, prevYMStr, tasks, quadrant, quadrantMsg, memo, newRatio, unitPrice, unitGoal, menuData);
  }
}

// ─ メニュー構成比取得 ─
function getMenuRatios(storeName, targetYM) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const ws = ss.getSheetByName(SHEET_SALES);
    if (!ws) return null;
    const rows = ws.getDataRange().getValues();
    const counts = { matsu:0, mayu:0, matsuMayu:0, matsuMayuPerm:0, other:0 };
    const amounts = { matsu:0, mayu:0, matsuMayu:0, matsuMayuPerm:0, other:0 };
    for (let i = 1; i < rows.length; i++) {
      const store = String(rows[i][0] || "").trim();
      if (store !== storeName) continue;
      const kubun = String(rows[i][5] || "").trim();
      if (kubun !== "施術") continue;
      if (targetYM) {
        const dn = parseFloat(rows[i][1]) || 0;
        if (Math.floor(dn / 100) !== targetYM) continue;
      }
      const mn = (String(rows[i][7] || "") + " " + String(rows[i][8] || "")).trim();
      const cnt = parseFloat(rows[i][11]) || 1;
      const amt = parseFloat(rows[i][12]) || 0;
      const hasMatsu   = mn.includes("まつ毛パーマ") || mn.includes("マツパ") || mn.includes("まつパ");
      const hasMayuWax = mn.includes("眉毛ワックス") || mn.includes("眉ワックス");
      const hasMayuPerm= mn.includes("眉毛パーマ")  || mn.includes("眉パーマ");
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

function createAgendaDoc(title, s, sPrev, prevYMStr, tasks, quadrant, quadrantMsg, memo, newRatio, unitPrice, unitGoal, menuData) {
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  body.clear();

  const today = new Date();
  const dateStr = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd");
  const fmt = n => Math.round(n || 0).toLocaleString();
  const newPct = Math.round(newRatio * 100);
  const pct = (a, b) => b > 0 ? Math.round(a / b * 100) + "%" : "—";

  // タイトル
  const titlePara = body.appendParagraph(title);
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titlePara.editAsText().setForegroundColor("#1e40af").setFontSize(18).setBold(true);
  const metaPara = body.appendParagraph("担当SV: " + s.sv + "　作成日: " + dateStr + "　象限: 【" + quadrant + "】");
  metaPara.editAsText().setFontSize(11).setForegroundColor("#64748b");
  body.appendParagraph("");

  const trend = (cur, prev) => {
    if (!prev || prev === 0) return "";
    const diff = cur - prev;
    const pct2 = Math.round(Math.abs(diff) / prev * 100);
    return diff > 0 ? " ▲" + pct2 + "%" : diff < 0 ? " ▼" + pct2 + "%" : " →";
  };

  // 1. 数値サマリー（売上・KPI まとめて一表）
  const h1 = body.appendParagraph("1. 数値サマリー");
  h1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h1.editAsText().setForegroundColor("#1e40af").setBold(true);
  const prevLabel = sPrev ? prevYMStr + "実績" : "前月なし";
  const nextResPct = Math.round((s.次回予約率実績 || 0) * 100);
  const nextResAlert = nextResPct < 35 ? " ⚠" : " ✓";
  const unitAlert = unitPrice < 5020 ? " ⚠" : " ✓";
  const t1 = body.appendTable([
    ["指標", "当月実績", "目標", "達成率", prevLabel, "前月比"],
    ["売上",       "¥"+fmt(s.実績売上),       "¥"+fmt(s.売上目標),       (s.達成率||0)+"%",                         sPrev?"¥"+fmt(sPrev.実績売上):"—",                trend(s.実績売上,sPrev?.実績売上)],
    ["月末見込",   "¥"+fmt(s.見込み売上),     "¥"+fmt(s.売上目標),       (s.見込み達成率||0)+"%",                   "—",                                              "—"],
    ["総客数",     fmt(s.総客数実績)+"人",     fmt(s.総客数目標)+"人",    pct(s.総客数実績,s.総客数目標)+"%",       sPrev?fmt(sPrev.総客数実績)+"人":"—",             trend(s.総客数実績,sPrev?.総客数実績)],
    ["新規客数",   fmt(s.新規実績)+"人",       fmt(s.新規目標)+"人",      pct(s.新規実績,s.新規目標)+"%",           sPrev?fmt(sPrev.新規実績)+"人":"—",               trend(s.新規実績,sPrev?.新規実績)],
    ["再来客数",   fmt(s.再来実績)+"人",       fmt(s.再来目標)+"人",      pct(s.再来実績,s.再来目標)+"%",           sPrev?fmt(sPrev.再来実績)+"人":"—",               trend(s.再来実績,sPrev?.再来実績)],
    ["客単価",     "¥"+fmt(s.客単価実績)+unitAlert, "¥"+fmt(s.客単価目標), pct(s.客単価実績,s.客単価目標)+"%",     sPrev?"¥"+fmt(sPrev.客単価実績):"—",             trend(s.客単価実績,sPrev?.客単価実績)],
    ["次回予約率", nextResPct+"%"+nextResAlert,"35%以上",                 "—",                                       sPrev?Math.round((sPrev.次回予約率実績||0)*100)+"%":"—", "—"],
    ["回数券売上", "¥"+fmt(s.回数券売上実績), "¥"+fmt(s.回数券売上目標), pct(s.回数券売上実績,s.回数券売上目標)+"%",sPrev?"¥"+fmt(sPrev.回数券売上実績):"—",         trend(s.回数券売上実績,sPrev?.回数券売上実績)],
    ["物販売上",   "¥"+fmt(s.物販売上実績),   "BP:¥29,700",              "—",                                       sPrev?"¥"+fmt(sPrev.物販売上実績):"—",            trend(s.物販売上実績,sPrev?.物販売上実績)],
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
  const strategies = getRelevantStrategies(quadrant, { nextRes: s.次回予約率実績 || 0, unitPrice, newGuest: s.新規実績 || 0 });
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
    DriveApp.getFileById(doc.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(shareErr) { Logger.log("doc share error: " + shareErr); }
  return doc.getUrl();
}

function createAgendaSlides(title, s, sPrev, prevYMStr, tasks, quadrant, quadrantMsg, memo, newRatio, unitPrice, unitGoal, menuData) {
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
    if (!prev || prev === 0) return "";
    const d = cur - prev;
    const p = Math.round(Math.abs(d) / prev * 100);
    return d > 0 ? " ▲" + p + "%" : d < 0 ? " ▼" + p + "%" : " →";
  };
  const prevL = sPrev ? prevYMStr : "";

  // スライド1: タイトル
  const sl1 = pres.getSlides()[0];
  clearSlide(sl1);
  setBg(sl1, BG_DARK);
  addBox(sl1, title, 40, 140, 880, 100, 26, true, "#ffffff");
  addBox(sl1, "担当SV: " + s.sv + "　" + dateStr + "　象限:【" + quadrant + "】", 40, 260, 880, 40, 16, false, "#94a3b8");
  addBox(sl1, quadrantMsg, 40, 310, 880, 50, 13, false, "#60a5fa");

  // スライド2: 数値サマリー（テーブル左 + 達成率グラフ右）
  const sl2 = pres.appendSlide();
  clearSlide(sl2);
  setBg(sl2, BG_LIGHT);
  addBox(sl2, "1. 数値サマリー", 40, 12, 500, 38, 18, true, ACCENT);
  if (prevL) addBox(sl2, "比較: " + prevL, 560, 12, 360, 32, 11, false, "#64748b");
  const salesPct = s.達成率 || 0;
  const mkPct2 = s.見込み達成率 || 0;
  const nextResPct = Math.round((s.次回予約率実績 || 0) * 100);
  const prevNextRes = sPrev ? Math.round((sPrev.次回予約率実績||0)*100) : null;
  const kpiRows = [
    ["指標", "当月実績", "目標", "達成率", prevL||"前月", "前月比"],
    ["売上",      "¥"+fmt(s.実績売上),       "¥"+fmt(s.売上目標),      salesPct+"%",                          sPrev?"¥"+fmt(sPrev.実績売上):"—",       sPrev?tr(s.実績売上,sPrev.実績売上):"—"],
    ["月末見込",  "¥"+fmt(s.見込み売上),     "¥"+fmt(s.売上目標),      mkPct2+"%",                            "—",                                     "—"],
    ["総客数",    fmt(s.総客数実績)+"人",     fmt(s.総客数目標)+"人",   pct(s.総客数実績,s.総客数目標)+"%",   sPrev?fmt(sPrev.総客数実績)+"人":"—",    sPrev?tr(s.総客数実績,sPrev.総客数実績):"—"],
    ["新規",      fmt(s.新規実績)+"人",       fmt(s.新規目標)+"人",     pct(s.新規実績,s.新規目標)+"%",       sPrev?fmt(sPrev.新規実績)+"人":"—",      sPrev?tr(s.新規実績,sPrev.新規実績):"—"],
    ["再来",      fmt(s.再来実績)+"人",       fmt(s.再来目標)+"人",     pct(s.再来実績,s.再来目標)+"%",       sPrev?fmt(sPrev.再来実績)+"人":"—",      sPrev?tr(s.再来実績,sPrev.再来実績):"—"],
    ["客単価",    "¥"+fmt(s.客単価実績),     "¥"+fmt(s.客単価目標),   pct(s.客単価実績,s.客単価目標)+"%",   sPrev?"¥"+fmt(sPrev.客単価実績):"—",    sPrev?tr(s.客単価実績,sPrev.客単価実績):"—"],
    ["次回予約率",nextResPct+"%" + (nextResPct<35?" ⚠":""), "35%以上", "—",                                    prevNextRes!==null?prevNextRes+"%":"—",   "—"],
    ["回数券",    "¥"+fmt(s.回数券売上実績), "¥"+fmt(s.回数券売上目標),pct(s.回数券売上実績,s.回数券売上目標)+"%",sPrev?"¥"+fmt(sPrev.回数券売上実績):"—",sPrev?tr(s.回数券売上実績,sPrev.回数券売上実績):"—"],
    ["物販",      "¥"+fmt(s.物販売上実績),   "BP:¥29,700",             "—",                                    sPrev?"¥"+fmt(sPrev.物販売上実績):"—",   sPrev?tr(s.物販売上実績,sPrev.物販売上実績):"—"],
  ];
  addTable(sl2, kpiRows, 20, 56, 530, 450, "#dbeafe", 9);
  // 達成率グラフ（右側）
  const chartLabels = ["売上", "見込み", "総客数", "新規", "再来", "客単価", "回数券"];
  const chartVals = [
    salesPct, mkPct2,
    pct(s.総客数実績,s.総客数目標), pct(s.新規実績,s.新規目標),
    pct(s.再来実績,s.再来目標),    pct(s.客単価実績,s.客単価目標),
    pct(s.回数券売上実績,s.回数券売上目標)
  ];
  addChart(sl2, chartLabels, chartVals, "達成率(%)", 565, 56, 390, 450);

  // スライド4: メニュー構成比
  const sl4m = pres.appendSlide();
  clearSlide(sl4m);
  setBg(sl4m, BG_LIGHT);
  addBox(sl4m, "2. メニュー構成比", 40, 15, 500, 40, 18, true, ACCENT);
  if (menuData && menuData.total > 0) {
    // 円グラフ（左側）
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
        .setTitle("メニュー構成比（件数）")
        .setDimensions(460, 420)
        .setOption("legend.position", "bottom")
        .build();
      sl4m.insertImage(pieChart.getAs("image/png"), 20, 60, 460, 420);
    } catch(pe) { Logger.log("pie error: " + pe); }
    // テーブル（右側）
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
    addTable(sl4m, mRows, 500, 60, 440, 420, "#dbeafe", 10);
  } else {
    addBox(sl4m, "メニューデータなし（売上明細シートに施術データが必要です）", 40, 80, 880, 60, 13, false, "#64748b");
  }

  // スライド5: 顧客象限
  const sl4 = pres.appendSlide();
  clearSlide(sl4);
  setBg(sl4, BG_DARK);
  addBox(sl4, "3. 顧客象限分析", 40, 20, 880, 45, 20, true, "#60a5fa");
  addBox(sl4, "現在のポジション: 【" + quadrant + "】", 40, 80, 880, 50, 22, true, "#fbbf24");
  addBox(sl4, quadrantMsg, 40, 145, 880, 50, 14, false, "#e2e8f0");
  addBox(sl4,
    "新規比率: " + newPct + "% / 再来: " + (100 - newPct) + "%\n" +
    "客単価: ¥" + fmt(unitPrice) + "  （" + (unitPrice >= unitGoal ? "▲ 目標達成" : "▼ 目標未達") + "）\n\n" +
    "集客サイクル: お試し層 → リピーター → VIP → 優良新規 → 増員 → 循環",
    40, 210, 880, 180, 14, false, "#94a3b8");

  // スライド6: 推奨アクション
  // 課題分析スライド
  const slIssue = pres.appendSlide();
  clearSlide(slIssue);
  setBg(slIssue, BG_LIGHT);
  addBox(slIssue, "4. 課題分析", 40, 15, 880, 40, 18, true, ACCENT);
  const issues = generateIssues(sPrev || s, unitPrice, unitGoal, newRatio);
  if (issues.length === 0) {
    addBox(slIssue, "現時点で大きな課題は検出されませんでした。", 40, 80, 880, 60, 13, false, "#64748b");
  } else {
    const issueRows = [["課題", "現状", "コンサルコメント"]];
    issues.forEach(iss => issueRows.push([iss.title, iss.detail, iss.comment]));
    addTable(slIssue, issueRows, 20, 62, 920, Math.min(62 + issues.length * 50 + 30, 450), "#fee2e2", 10);
  }

  // 推奨アクションスライド
  const sl5 = pres.appendSlide();
  clearSlide(sl5);
  setBg(sl5, BG_LIGHT);
  addBox(sl5, "5. 推奨アクション（ベストプラクティスより）", 40, 15, 880, 40, 18, true, ACCENT);
  const strategies = getRelevantStrategies(quadrant, { nextRes: s.次回予約率実績 || 0, unitPrice, newGuest: s.新規実績 || 0 });
  if (strategies.length > 0) {
    const stratRows = [["施策名", "対象KPI", "重要度", "推奨タイミング"]];
    strategies.slice(0, 8).forEach(st => stratRows.push([st.name||"", st.target||"—", st.importance||"—", st.timing||"—"]));
    addTable(sl5, stratRows, 20, 62, 920, 440, "#fef9c3", 10);
  } else {
    addBox(sl5, "施策データなし", 40, 80, 880, 60, 13, false, "#64748b");
  }

  // スライド: その他
  const sl7 = pres.appendSlide();
  clearSlide(sl7);
  setBg(sl7, BG_DARK);
  addBox(sl7, "6. その他", 40, 20, 880, 45, 18, true, "#60a5fa");
  addBox(sl7, memo || "（なし）", 40, 80, 880, 300, 14, false, "#e2e8f0");

  pres.saveAndClose();
  // 他アカウントからも閲覧できるようリンク共有を設定
  try {
    DriveApp.getFileById(pres.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(shareErr) { Logger.log("slides share error: " + shareErr); }
  return pres.getUrl();
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
