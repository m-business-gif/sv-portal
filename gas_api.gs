// =============================================
// SV加盟店管理ポータル - GAS API v4
// =============================================

const SS_ID        = "1K-4ub8YvFh__JrseNKGiCkGigDYykraIwocOhLQLevY";
const SHEET_GOAL   = "加盟店目標";
const SHEET_REAL   = "実数値";
const SHEET_TASK   = "タスクボード";
const SHEET_MIKOMI = "見込み数値";
const SHEET_STAFF  = "スタッフランク";
const SHEET_SALES  = "スタッフ売上（9~2月）";
const SHEET_CONFIG = "【眉毛】加盟店管理集計";

// タスクボード列定義
// A:店舗名 B:担当SV C:カテゴリ D:タスク名 E:ステータス F:優先度 G:メモ H:完了日

function doGet(e) {
  try {
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
    return json({ stores: getStores(), tasks: getTasks(), staffRanks: getStaffRanks(), staffSales: getStaffSales(), config: getConfig() });
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

// ─ 店舗データ ─
function getStores() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const rowsG = ss.getSheetByName(SHEET_GOAL).getDataRange().getValues();
  const rowsR = ss.getSheetByName(SHEET_REAL).getDataRange().getValues();
  const rowsM = ss.getSheetByName(SHEET_MIKOMI).getDataRange().getValues();

  const today = new Date();
  const curYM = today.getFullYear() * 100 + (today.getMonth() + 1);

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

  const elapsed = today.getDate() - 1;
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

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

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
