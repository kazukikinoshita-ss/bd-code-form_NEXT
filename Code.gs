// ===============================================
// Code.gs v20.0 (サジェスト完全復旧・検索エラー解消版)
// ===============================================

const SPREADSHEET_ID = '12xNumVKAx5pp4eHMxn1iDZsBKw9_CjNqnducZeAbNqU'; // ★実際のIDに置き換えてください
const SETTINGS_SHEET_NAME = '設定';
const PROP_KEY_MAX_PARENT = 'CURRENT_MAX_PARENT_ID';

const PROJECT_MASTER_SHEET_NAME = '[マスタ] Project Code';
const ORG_BU_MASTER_SHEET_NAME = '[マスタ] 組織/BU（FY25）';
const ACCOUNT_MASTER_SHEET_NAME = '[マスタ] 勘定科目';

const MP_LINK_CELL = 'L1';
const MP_SHEET_NAME_CELL = 'L2';
const DATE_HEADER_ROW = 6; 
const DATA_START_ROW = 7;
const DEFAULT_MAX_COL_LIMIT = 'DL';

const SEARCHABLE_COLUMN_NAMES = ["MPコード", "BU", "計上部門", "勘定科目", "対象PJC", "仕入先"];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MPコード管理')
    .addItem('ID採番の同期 (手動修正後に実行)', 'syncIdCache')
    .addSeparator()
    .addItem('金額データ検索・更新ツール', 'openWebAppSidebar')
    .addSeparator()
    .addItem('【重要】初回設定(権限承認)', 'runAuthCheck')
    .addItem('キャッシュクリア(再読込)', 'clearAppCache')
    .addToUi();
}

function runAuthCheck() {
  try {
    Sheets.Spreadsheets.Values.get(SPREADSHEET_ID, 'A1');
    SpreadsheetApp.getUi().alert("権限の確認が完了しました。\nWebアプリをご利用いただけます。");
  } catch (e) {
    SpreadsheetApp.getUi().alert("エラーが発生しました: " + e.message);
  }
}

function openWebAppSidebar() {
  const html = HtmlService.createTemplateFromFile('WebApp').evaluate().setTitle('金額データ検索・更新');
  SpreadsheetApp.getUi().showSidebar(html);
}

function doGet(e) {
  const page = e.parameter.page;
  let template, title;
  const baseUrl = ScriptApp.getService().getUrl();
  
  if (page === 'tool' || page === 'WebApp') { template = HtmlService.createTemplateFromFile('WebApp'); title = '金額データ検索・更新'; } 
  else if (page === 'new') { template = HtmlService.createTemplateFromFile('index'); title = 'MPコード新規登録フォーム'; } 
  else if (page === 'update') { template = HtmlService.createTemplateFromFile('update'); title = '情報マスタ更新フォーム'; } 
  else { template = HtmlService.createTemplateFromFile('home'); title = '販管費管理メニュー'; }
  
  template.topUrl = baseUrl;
  return template.evaluate().setTitle(title).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function getWebAppUrl() { return ScriptApp.getService().getUrl(); }

function getSheetNames() {
  try {
    const settingsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) throw new Error(`「${SETTINGS_SHEET_NAME}」シートが見つかりません。`);
    const values = settingsSheet.getRange('B2:C5').getValues();
    return { budgetMaster: values[0][0], infoMaster: values[1][0], externalInputUrl: values[2][1], errorLog: values[3][0] };
  } catch (e) {
    return { budgetMaster: 'MPコードマスタ_raw', infoMaster: '情報マスタ_raw', externalInputUrl: '', errorLog: 'error_log' };
  }
}

// 💡 10行目からのRoutingMap読み込み
function getRouteMap() {
  try {
    const settingsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SETTINGS_SHEET_NAME);
    const lastRow = settingsSheet.getLastRow();
    if (lastRow < 10) return [];
    const data = settingsSheet.getRange(10, 1, lastRow - 9, 5).getValues();
    return data.filter(row => row[1] && String(row[1]).trim() !== ""); 
  } catch(e) { return []; }
}

function getManagerSpreadsheet() {
  var url = getSheetNames().externalInputUrl;
  if (url) {
    try { return SpreadsheetApp.openByUrl(url); } 
    catch(e) {
      var match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (match) return SpreadsheetApp.openById(match[1]);
    }
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ===============================================
// Utils & 大容量キャッシュ
// ===============================================
function formatDateFast(date) {
  if (!(date instanceof Date) || isNaN(date)) return "";
  const y = date.getFullYear();
  const m = ('0' + (date.getMonth() + 1)).slice(-2);
  const d = ('0' + date.getDate()).slice(-2);
  return `${y}/${m}/${d}`;
}

function getYearMonth(val) {
  if (!val) return null;
  let d = val;
  if (!(d instanceof Date)) {
    const t = new Date(val);
    if (!isNaN(t.getTime())) d = t;
    else return null;
  }
  return d.getFullYear() * 100 + (d.getMonth() + 1);
}

function safeString(val) {
  if (val instanceof Date) return formatDateFast(val);
  return (val === null || val === undefined) ? "" : String(val);
}

function safeNumber(val) { return (typeof val === 'number' && !isNaN(val)) ? val : 0; }

function putCacheLarge(key, data, expirationInSeconds) {
  const cache = CacheService.getScriptCache();
  const json = JSON.stringify(data);
  const chunkSize = 95000;
  if (json.length <= chunkSize) { cache.put(key, json, expirationInSeconds); return; }
  const chunks = {};
  let chunkIndex = 0;
  for (let i = 0; i < json.length; i += chunkSize) {
    chunks[`${key}_${chunkIndex}`] = json.substr(i, chunkSize);
    chunkIndex++;
  }
  chunks[`${key}_count`] = String(chunkIndex);
  cache.putAll(chunks, expirationInSeconds);
}

function getCacheLarge(key) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const countStr = cache.get(`${key}_count`);
  if (!countStr) return null;
  const count = parseInt(countStr, 10);
  const keys = [];
  for (let i = 0; i < count; i++) { keys.push(`${key}_${i}`); }
  const chunks = cache.getAll(keys);
  let json = "";
  for (let i = 0; i < count; i++) {
    const chunk = chunks[`${key}_${i}`];
    if (!chunk) return null;
    json += chunk;
  }
  return JSON.parse(json);
}

function clearAppCache() {
  const cache = CacheService.getScriptCache();
  const keys = ["initialData_v20", "initialData_v20_count"];
  try {
    const routes = getRouteMap();
    routes.forEach(r => keys.push("input_all_headers_v20_" + r[1]));
  } catch(e) {}
  try { cache.removeAll(keys); } catch(e) {}
  return "キャッシュをクリアしました。";
}

function getCachedHeaders(sheet) {
  const cacheKey = "input_all_headers_v20_" + sheet.getName();
  const cached = getCacheLarge(cacheKey);
  if (cached) return cached;
  
  const lastColNum = Math.max(sheet.getLastColumn(), 50);
  const searchData = sheet.getRange(1, 1, Math.min(sheet.getLastRow() || 20, 20), lastColNum).getValues();
  let headerRowIdx = -1;
  
  for (let i = 0; i < searchData.length; i++) {
    const rowStr = searchData[i].join("");
    if (rowStr.includes("MPコード") && (rowStr.includes("一般科目") || rowStr.includes("グループ名") || rowStr.includes("商品ファミリ"))) { 
      headerRowIdx = i + 1; 
      break; 
    }
  }
  if(headerRowIdx === -1) headerRowIdx = DATE_HEADER_ROW; 

  const mainHeaders = sheet.getRange(headerRowIdx, 1, 1, lastColNum).getValues()[0].map(h => String(h).trim());
  const dateHeaders = sheet.getRange(headerRowIdx, 1, 1, lastColNum).getValues()[0];
  const dateHeadersInfo = dateHeaders.map(d => {
      if (d instanceof Date && !isNaN(d)) return { isDate: true, time: d.getTime() };
      return { isDate: false };
  });
  const result = { mainHeaders, dateHeadersInfo, headerRowIndex: headerRowIdx };
  putCacheLarge(cacheKey, result, 1200); 
  return result;
}

// ===============================================
// WebApp用 API群
// ===============================================
function getInitialSettings() {
  let defStart = "", defEnd = "";
  try {
      const settingsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SETTINGS_SHEET_NAME);
      // 💡 B7・B8からの日付取得に対応
      const vals = settingsSheet.getRange('B6:B8').getValues();
      const sVal = vals[0][0] || vals[1][0]; // B6かB7にある開始日
      const eVal = vals[1][0] || vals[2][0]; // B7かB8にある終了日

      if (sVal instanceof Date) defStart = formatDateFast(sVal).slice(0, 7).replace(/\//g, '-');
      else if (typeof sVal === 'string') defStart = sVal.replace(/\//g, '-'); 
      if (eVal instanceof Date) defEnd = formatDateFast(eVal).slice(0, 7).replace(/\//g, '-');
      else if (typeof eVal === 'string') defEnd = eVal.replace(/\//g, '-');
  } catch(e) {}

  const routes = getRouteMap();
  const routeList = routes.map((r, idx) => ({ index: idx, bu: String(r[0]).trim(), sheet: String(r[1]).trim(), url: String(r[2]).trim() }));
  return { routes: routeList, defaultStartDate: defStart, defaultEndDate: defEnd };
}

function getMasterDataByRouteIndex(routeIndex) {
  const routes = getRouteMap();
  const route = routes[Number(routeIndex)];
  if (!route) throw new Error("指定されたシートが見つかりません。");

  const sheetName = String(route[1]).trim();
  const ssUrl = route[2];

  const accountMap = new Map();
  try {
      const defaultSS = SpreadsheetApp.openById(SPREADSHEET_ID);
      const accSheet = defaultSS.getSheetByName(ACCOUNT_MASTER_SHEET_NAME);
      if (accSheet && accSheet.getLastRow() >= 2) {
          const accData = accSheet.getRange(2, 2, accSheet.getLastRow() - 1, 2).getValues(); 
          accData.forEach(r => {
             if(String(r[1]).trim() && String(r[0]).trim()) accountMap.set(String(r[1]).trim(), String(r[0]).trim());
          });
      }
  } catch (e) { }

  const searchIndex = [];
  const groupSet = new Set(), subjectSet = new Set(), globalDatesSet = new Set();
       
  try {
    let targetSS = (ssUrl && ssUrl.includes("spreadsheets")) ? SpreadsheetApp.openByUrl(ssUrl) : SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = targetSS.getSheetByName(sheetName);
    if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。`);

    const { mainHeaders, dateHeadersInfo, headerRowIndex } = getCachedHeaders(sheet);
    
    const bdCodeCol = mainHeaders.findIndex(h => h && h.includes("MPコード")); 
    const deptCol = mainHeaders.findIndex(h => h && (h.includes("グループ名") || h.includes("商品ファミリ") || h.includes("計上部門"))); 
    const accountCol = mainHeaders.findIndex(h => h && (h.includes("一般科目") || h.includes("勘定科目"))); 
    const pjcCol = mainHeaders.findIndex(h => h && (h.includes("プロジェクト名") || h.includes("対象PJC") || h.includes("PJC")));
    const vendorCol = mainHeaders.findIndex(h => h && h.includes("仕入先"));

    if (bdCodeCol === -1) throw new Error(`シート内に「MPコード」の列が見つかりません。`);
    dateHeadersInfo.forEach(d => { if (d.isDate) globalDatesSet.add(formatDateFast(new Date(d.time)).slice(0, 7).replace(/\//g, '-')); });

    const START_ROW = headerRowIndex + 1;
    const lastRow = sheet.getLastRow();
    if (lastRow >= START_ROW) {
      const data = sheet.getRange(START_ROW, 1, lastRow - START_ROW + 1, sheet.getLastColumn()).getValues();

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const bdCode = String(row[bdCodeCol] || "").trim();
        if (!bdCode) continue;

        const groupVal = deptCol !== -1 ? String(row[deptCol] || "").trim() : "";
        if(groupVal) groupSet.add(groupVal);

        const accountVal = accountCol !== -1 ? String(row[accountCol] || "").trim() : "";
        const subjectVal = accountMap.get(accountVal) || "その他"; 
        subjectSet.add(subjectVal);

        const globalId = `${routeIndex}___${i + START_ROW}`; 
        const item = { id: globalId, g: groupVal, s: subjectVal, o: {} };
        
        if (pjcCol !== -1) item.o["対象PJC"] = String(row[pjcCol]).trim();
        if (vendorCol !== -1) item.o["仕入先"] = String(row[vendorCol]).trim();
        if (accountCol !== -1) item.o["一般科目"] = accountVal;
        
        searchIndex.push(item);
      }
    }
  } catch(e) { throw new Error(`シート読込失敗: ${e.message}`); }

  const datesArray = Array.from(globalDatesSet).sort();
  return { groups: Array.from(groupSet).sort(), subjects: Array.from(subjectSet).sort(), otherHeaders: ["対象PJC", "仕入先", "一般科目"], dateOptions: datesArray, searchIndex: searchIndex };
}

function searchDataByMpCode(budgetCode, customDateRange) {
  const routes = getRouteMap();
  const results = [];
  const globalMonthlyHeadersSet = new Set();
  const targetBdCode = String(budgetCode).trim().toLowerCase();
  
  for (let r = 0; r < routes.length; r++) {
    const sheetName = String(routes[r][1]).trim();
    const ssUrl = routes[r][2];

    try {
      let targetSS = (ssUrl && ssUrl.includes("spreadsheets")) ? SpreadsheetApp.openByUrl(ssUrl) : SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = targetSS.getSheetByName(sheetName);
      if (!sheet) continue;

      const { mainHeaders, dateHeadersInfo, headerRowIndex } = getCachedHeaders(sheet);
      const bdCodeCol = mainHeaders.findIndex(h => h && h.includes("MPコード")); 
      if (bdCodeCol === -1) continue;

      const allValues = sheet.getDataRange().getValues();
      const START_ROW = headerRowIndex + 1;

      for (let i = START_ROW - 1; i < allValues.length; i++) {
        if (String(allValues[i][bdCodeCol]).trim().toLowerCase() === targetBdCode) {
           const item = parseRowDataToItem(allValues[i], r, i + 1, mainHeaders, dateHeadersInfo, customDateRange, globalMonthlyHeadersSet);
           results.push(item);
        }
      }
    } catch(e) { console.warn(`Error searching in row ${r}: ${e.message}`); }
  }

  const sortedMonthlyHeaders = Array.from(globalMonthlyHeadersSet).sort((a,b) => a - b);
  return JSON.stringify({ searchResults: results, monthlyHeaders: sortedMonthlyHeaders, basicInfo: { readOnly: SEARCHABLE_COLUMN_NAMES, editable: [] }, remarksInfo: { readOnly: ["備考5"], editable: ["備考1", "備考2", "備考3", "備考4"] } });
}

function parseRowDataToItem(rowValues, routeIndex, rowNum, mainHeaders, dateHeadersInfo, customDateRange, globalMonthlyHeadersSet) {
    const bdCodeCol = mainHeaders.findIndex(h => h && h.includes("MPコード")); 
    const buCol = mainHeaders.findIndex(h => h && h.includes("BU")); 
    const deptCol = mainHeaders.findIndex(h => h && (h.includes("グループ名") || h.includes("商品ファミリ") || h.includes("計上部門"))); 
    const accountCol = mainHeaders.findIndex(h => h && (h.includes("一般科目") || h.includes("勘定科目"))); 
    const pjcCol = mainHeaders.findIndex(h => h && (h.includes("プロジェクト名") || h.includes("対象PJC") || h.includes("PJC")));
    const vendorCol = mainHeaders.findIndex(h => h && h.includes("仕入先"));

    const dateColMap = new Map();
    dateHeadersInfo.forEach((d, idx) => {
      if (d.isDate) {
        const ym = formatDateFast(new Date(d.time)).slice(0, 7).replace(/\//g, '-');
        let startF = customDateRange ? customDateRange.startDate : "";
        let endF = customDateRange ? customDateRange.endDate : "";
        if (!customDateRange || (ym >= startF && ym <= endF)) {
          globalMonthlyHeadersSet.add(d.time);
          dateColMap.set(d.time, idx);
        }
      }
    });

    const item = { rowNumber: `${routeIndex}___${rowNum}` };
    if (bdCodeCol !== -1) item["MPコード"] = safeString(rowValues[bdCodeCol]);
    if (buCol !== -1) item["BU"] = safeString(rowValues[buCol]);
    if (deptCol !== -1) item["計上部門"] = safeString(rowValues[deptCol]);
    if (accountCol !== -1) item["勘定科目"] = safeString(rowValues[accountCol]);
    if (pjcCol !== -1) item["対象PJC"] = safeString(rowValues[pjcCol]);
    if (vendorCol !== -1) item["仕入先"] = safeString(rowValues[vendorCol]);
    
    mainHeaders.forEach((h, c) => { if (h && h.startsWith("備考")) item[h] = safeString(rowValues[c]); });
    Array.from(dateColMap.entries()).forEach(([timeMs, colIdx]) => { item[timeMs] = safeNumber(rowValues[colIdx]); });
    return item;
}

function fetchPageDataForWebApp(globalRowIds, customDateRange) {
  const routes = getRouteMap();
  const results = [];
  const globalMonthlyHeadersSet = new Set();
  
  const groupsByRoute = {};
  globalRowIds.forEach(id => {
    const parts = id.split("___");
    if (parts.length < 2) return;
    if (!groupsByRoute[parts[0]]) groupsByRoute[parts[0]] = [];
    groupsByRoute[parts[0]].push(parseInt(parts[1], 10));
  });

  Object.keys(groupsByRoute).forEach(rIdxStr => {
    const rIdx = parseInt(rIdxStr, 10);
    const route = routes[rIdx];
    if (!route) return;
    
    const sheetName = String(route[1]).trim();
    const ssUrl = route[2];

    try {
      let targetSS = (ssUrl && ssUrl.includes("spreadsheets")) ? SpreadsheetApp.openByUrl(ssUrl) : SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = targetSS.getSheetByName(sheetName);
      if (!sheet) return;

      const { mainHeaders, dateHeadersInfo } = getCachedHeaders(sheet);
      const sortedIndices = groupsByRoute[rIdxStr].sort((a,b) => a - b);
      const groups = getRowGroupsWithTolerance(sortedIndices, 10);
      const ssId = targetSS.getId();
      
      const ranges = groups.map(g => `'${sheetName}'!A${g.start}:${DEFAULT_MAX_COL_LIMIT}${g.start + g.numRows - 1}`);
      const response = Sheets.Spreadsheets.Values.batchGet(ssId, { ranges: ranges, valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'SERIAL_NUMBER' });
      const valueRanges = response.valueRanges || [];

      groups.forEach((g, idx) => {
        const blockValues = valueRanges[idx] ? valueRanges[idx].values : [];
        if (!blockValues) return;
        sortedIndices.forEach(targetRow => {
            if (targetRow >= g.start && targetRow < g.start + g.numRows) {
                const relativeIdx = targetRow - g.start;
                // 💡 空行スキップによるエラーを防止
                if (relativeIdx < blockValues.length && blockValues[relativeIdx]) {
                    const item = parseRowDataToItem(blockValues[relativeIdx], rIdx, targetRow, mainHeaders, dateHeadersInfo, customDateRange, globalMonthlyHeadersSet);
                    results.push(item);
                }
            }
        });
      });
    } catch(e) { console.warn("Fetch Error: " + e.message); }
  });

  const sortedMonthlyHeaders = Array.from(globalMonthlyHeadersSet).sort((a,b) => a - b);
  return JSON.stringify({ searchResults: results, monthlyHeaders: sortedMonthlyHeaders, basicInfo: { readOnly: SEARCHABLE_COLUMN_NAMES, editable: [] }, remarksInfo: { readOnly: ["備考5"], editable: ["備考1", "備考2", "備考3", "備考4"] } });
}

function getRowGroupsWithTolerance(indices, tolerance) {
  if (!indices || !indices.length) return [];
  indices.sort((a, b) => a - b);
  const groups = []; let start = indices[0], prev = indices[0];
  for (let i = 1; i < indices.length; i++) {
    if (indices[i] <= prev + 1 + tolerance) prev = indices[i];
    else { groups.push({ start: start, numRows: prev - start + 1 }); start = indices[i]; prev = indices[i]; }
  }
  groups.push({ start: start, numRows: prev - start + 1 }); return groups;
}

function getMpDataOnly(globalRowIds, bdCodes) {
  const res = {};
  if (!globalRowIds || !globalRowIds.length) return { mpDataMap: res, mpSheetName: 'MP', error: null };
  try {
    const parts = globalRowIds[0].split("___");
    const rIdx = parseInt(parts[0], 10);
    const routes = getRouteMap();
    const route = routes[rIdx]; 
    if (!route) throw new Error("ルートマップに該当する行が存在しません。");
    
    const mpSheetName = String(route[3] || "").trim(); 
    const mpUrl = String(route[4] || "").trim();      
    
    if (!mpUrl || !mpSheetName) {
      return { mpDataMap: res, mpSheetName: 'MP', error: `設定シートのD/E列にMPマスタ情報が入力されていません。` };
    }

    const match = mpUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) return { mpDataMap: res, mpSheetName: mpSheetName, error: `設定シートのMP参照用URLの形式が不正です。` };
    const mpSSId = match[1];
    
    const mpSheet = SpreadsheetApp.openById(mpSSId).getSheetByName(mpSheetName);
    if (!mpSheet) return { mpDataMap: res, mpSheetName: mpSheetName, error: `MPマスタ「${mpSheetName}」が見つかりません。` };

    const cBd = 0; 
    const allBd = Sheets.Spreadsheets.Values.get(mpSSId, `'${mpSheetName}'!A${DATA_START_ROW}:A${mpSheet.getLastRow()}`).values.flat().map(String);
    const bdMapObj = {};
    allBd.forEach((bd, i) => { if(bd) bdMapObj[String(bd).trim().toLowerCase()] = i + DATA_START_ROW; });
    
    const targetIndices = [];
    bdCodes.forEach(b => {
        const mpRowIdx = bdMapObj[String(b).trim().toLowerCase()];
        if (mpRowIdx) targetIndices.push(mpRowIdx);
    });

    const dateColMap = new Map();
    const rawDateHeaders = mpSheet.getRange(DATE_HEADER_ROW, 1, 1, mpSheet.getLastColumn()).getValues()[0];
    rawDateHeaders.forEach((d, idx) => { if(d instanceof Date && !isNaN(d)) dateColMap.set(d.getTime(), idx); });

    const fetchedRows = new Map();
    if (targetIndices.length > 0) {
        const groups = getRowGroupsWithTolerance([...new Set(targetIndices)].sort((a,b)=>a-b), 200);
        const mpResp = Sheets.Spreadsheets.Values.batchGet(mpSSId, { ranges: groups.map(g => `'${mpSheetName}'!A${g.start}:${DEFAULT_MAX_COL_LIMIT}${g.start + g.numRows - 1}`), valueRenderOption: 'UNFORMATTED_VALUE' });
        (mpResp.valueRanges || []).forEach(vr => (vr.values || []).forEach(row => { 
            if (row.length > cBd && row[cBd]) fetchedRows.set(String(row[cBd]).trim().toLowerCase(), row);
        }));
    }

    globalRowIds.forEach((gId, i) => {
        const bd = String(bdCodes[i]).trim().toLowerCase();
        const row = fetchedRows.get(bd);
        if (!row) { res[gId] = "__NOT_FOUND__"; return; }

        const m = {};
        Array.from(dateColMap.entries()).forEach(([timeMs, colIdx]) => {
          if (colIdx < row.length) {
              const val = row[colIdx];
              m[timeMs] = (val === "" || val === null || val === undefined) ? null : Number(val);
          } else { m[timeMs] = null; }
        });
        res[gId] = m;
    });
    return { mpDataMap: res, mpSheetName: mpSheetName, error: null };
  } catch (e) { return { error: e.message }; }
}

function updateMultipleData(dataArray) {
  if (!dataArray.length) return 'No Data';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('他の人が編集中です。しばらくお待ちください。');
  const routes = getRouteMap();
  try {
    const itemsByRoute = {};
    dataArray.forEach(item => {
      const parts = String(item.rowNumber).split("___");
      const rIdx = parts[0];
      if(!itemsByRoute[rIdx]) itemsByRoute[rIdx] = [];
      const clonedItem = JSON.parse(JSON.stringify(item));
      clonedItem.rowNumber = parseInt(parts[1], 10); 
      itemsByRoute[rIdx].push(clonedItem);
    });

    Object.keys(itemsByRoute).forEach(rIdxStr => {
      const rIdx = parseInt(rIdxStr, 10);
      const route = routes[rIdx];
      if(!route) return;
      
      const sheetName = String(route[1]).trim();
      const ssUrl = route[2];

      let targetSS = (ssUrl && ssUrl.includes("spreadsheets")) ? SpreadsheetApp.openByUrl(ssUrl) : SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = targetSS.getSheetByName(sheetName);
      if(!sheet) return;

      const { mainHeaders, dateHeadersInfo } = getCachedHeaders(sheet);
      const dateColIndices = {};
      dateHeadersInfo.forEach((d, i) => { if (d.isDate) dateColIndices[d.time] = i; });
      const b1Idx = mainHeaders.findIndex(h => h && h.includes("備考1"));
      const b2Idx = mainHeaders.findIndex(h => h && h.includes("備考2"));
      const b3Idx = mainHeaders.findIndex(h => h && h.includes("備考3"));
      const b4Idx = mainHeaders.findIndex(h => h && h.includes("備考4"));
      
      const allowedCols = [b1Idx, b2Idx, b3Idx, b4Idx].filter(idx => idx !== -1);
      const minAllowedCol = Math.min(...allowedCols, ...Object.values(dateColIndices));
      const numCols = mainHeaders.length - minAllowedCol;
      const colMap = { "備考1": b1Idx, "備考2": b2Idx, "備考3": b3Idx, "備考4": b4Idx };

      const subItems = itemsByRoute[rIdxStr].sort((a, b) => a.rowNumber - b.rowNumber);
      let currentGroup = [subItems[0]];
      for(let i=1; i<subItems.length; i++) {
        if(subItems[i].rowNumber === subItems[i-1].rowNumber + 1) currentGroup.push(subItems[i]);
        else { processUpdateGroup(sheet, currentGroup, minAllowedCol + 1, numCols, minAllowedCol, colMap, dateColIndices); currentGroup = [subItems[i]]; }
      }
      processUpdateGroup(sheet, currentGroup, minAllowedCol + 1, numCols, minAllowedCol, colMap, dateColIndices);
    });

    return `${dataArray.length}件のデータを更新しました。`;
  } finally { lock.releaseLock(); }
}

function processUpdateGroup(sheet, items, startCol, numCols, minAllowedCol, colMap, dateColIndices) {
  const range = sheet.getRange(items[0].rowNumber, startCol, items.length, numCols), values = range.getValues();
  items.forEach((item, i) => {
    Object.keys(item).forEach(key => {
      if (key === 'rowNumber') return;
      let targetColIdx = !isNaN(key) ? dateColIndices[parseInt(key, 10)] : colMap[key];
      if (targetColIdx !== undefined && targetColIdx !== -1) {
        const rel = targetColIdx - minAllowedCol;
        if (rel >= 0 && rel < numCols) values[i][rel] = (!isNaN(key) && item[key] !== "") ? Number(item[key]) : item[key];
      }
    });
  });
  range.setValues(values);
}

function searchBudgetCode(budgetCode) {
  try {
    const sheetNames = getSheetNames();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetNames.infoMaster);
    const data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();
    const formatToYYYYMMDD = (val) => {
      if (!val) return '';
      if (val instanceof Date) return `${val.getFullYear()}-${String(val.getMonth() + 1).padStart(2, '0')}-${String(val.getDate()).padStart(2, '0')}`;
      return String(val);
    };
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === budgetCode) {
        return {
          row: i + 2,
          data: {
            bu: data[i][1], department: data[i][2], account: data[i][3], pjc: data[i][4], vendor: data[i][5], summary: data[i][6], amount: data[i][10], notes: data[i][11],
            startDate: formatToYYYYMMDD(data[i][12]), endDate: formatToYYYYMMDD(data[i][13])
          }
        };
      }
    }
    return { error: "指定されたMPコードが見つかりませんでした。" };
  } catch (e) { return { error: "検索中にエラーが発生しました: " + e.message }; }
}

function updateInfoData(formData, rowNumber) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return { success: false, message: '【アクセス混雑】時間をおいて再度お試しください。' }; }
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const sheetNames = getSheetNames();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetNames.infoMaster);
    const updateTime = new Date();
    const newRowData = [formData.bu, formData.department, formData.account, formData.pjc, formData.vendor, formData.summary, "", updateTime, userEmail, formData.amount, formData.notes, formData.startDate, formData.endDate];
    sheet.getRange(rowNumber, 2, 1, 13).setValues([newRowData]);
    return { success: true, message: `MPコード ${formData.budgetCode} の情報を更新しました。` };
  } catch (e) { return { success: false, message: "更新中にエラーが発生しました: " + e.message }; } finally { lock.releaseLock(); }
}

// 💡 修正：サジェスト構築時のガード条件を完全復旧（クラッシュ防止）
function getBuSuggestions() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ]BUインポート');
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange('E2:E' + sheet.getLastRow()).getValues().map(row => row[0]).filter((v, i, s) => v && s.indexOf(v) === i);
  } catch (e) { return []; }
}
function getPjcSuggestions() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ] Project Code');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const targetColIndex = headers.indexOf('ジョブ名'), flagColIndex = headers.indexOf('有効/無効フラグ');
    if(targetColIndex === -1 || flagColIndex === -1) return [];
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    return values.filter(row => row[flagColIndex] === true).map(row => row[targetColIndex]).map(item => String(item).trim()).filter((v, i, s) => v && s.indexOf(v) === i);
  } catch (e) { return []; }
}
function getAccountSuggestions() {
  try {
    const targetSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ] 勘定科目');
    if (!targetSheet || targetSheet.getLastRow() < 6) return [];
    const values = targetSheet.getRange('B6:F' + targetSheet.getLastRow()).getValues();
    return values.filter(row => row[4] === true).map(row => row[1]).filter((v, i, s) => v && s.indexOf(v) === i);
  } catch (e) { return []; }
}
function getDepartmentSuggestions(selectedBu) {
  if (!selectedBu) return [];
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ]項目データ');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const headers = sheet.getRange('1:1').getValues()[0];
    const buColumnIndex = headers.indexOf(selectedBu);
    if (buColumnIndex === -1) return [];
    const columnLetter = String.fromCharCode('A'.charCodeAt(0) + buColumnIndex);
    return sheet.getRange(`${columnLetter}2:${columnLetter}`).getValues().map(row => row[0]).filter((v, i, s) => v && s.indexOf(v) === i);
  } catch (e) { return []; }
}
function getVendorSuggestions() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ] 取引先');
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange('A2:A' + sheet.getLastRow()).getValues().map(row => row[0]).filter((v, i, s) => v && s.indexOf(v) === i);
  } catch (e) { return []; }
}
function getParentIdSuggestions() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('MPコードマスタ_raw');
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange('B2:B' + sheet.getLastRow()).getValues().map(row => row[0]).filter((v, i, s) => v && s.indexOf(v) === i).sort((a, b) => b - a);
  } catch (e) { return []; }
}

function getInitialFormData() {
  try {
    const [parentIds, bus, pjcs, accounts, vendors] = [getParentIdSuggestions(), getBuSuggestions(), getPjcSuggestions(), getAccountSuggestions(), getVendorSuggestions()];
    return { parentIds: parentIds, bus: bus, pjcs: pjcs, accounts: accounts, vendors: vendors, routeMap: getRouteMap() };
  } catch (e) { return { error: e.message }; }
}

function processForm(formObject) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return { success: false, message: '【アクセス混雑】順番待ちが制限時間を超えました。' }; }
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetNames = getSheetNames();
    const budgetSheet = ss.getSheetByName(sheetNames.budgetMaster);
    const infoSheet = ss.getSheetByName(sheetNames.infoMaster);
    const registrationTime = new Date();
    const calculateEndDate = (src) => new Date(src.getFullYear() + (src.getMonth() + 1 >= 4 ? 2 : 1), 2, 31);
    const endDate = calculateEndDate(registrationTime);

    let newParentId, newChildId;
    let isSlide = formObject.slideParentId && formObject.slideParentId.trim() !== '';
    if (isSlide) {
      const parentId = formObject.slideParentId.trim();
      const lastRow = budgetSheet.getLastRow();
      const data = lastRow > 1 ? budgetSheet.getRange(2, 2, lastRow - 1, 2).getValues() : [];
      let maxChildId = 0, parentFound = false;
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString() == parentId) {
          parentFound = true;
          const childId = parseInt(data[i][1], 10) || 0;
          if (childId > maxChildId) maxChildId = childId;
        }
      }
      if (!parentFound) return { success: false, message: 'エラー: 指定された親IDが見つかりません。' };
      newParentId = parentId;
      newChildId = maxChildId + 1;
    } else {
      const scriptProps = PropertiesService.getScriptProperties();
      let cachedMax = parseInt(scriptProps.getProperty(PROP_KEY_MAX_PARENT), 10) || 0;
      if (cachedMax === 0) {
        const lastRow = budgetSheet.getLastRow();
        const data = lastRow > 1 ? budgetSheet.getRange(2, 2, lastRow - 1, 1).getValues() : [];
        for (let i = 0; i < data.length; i++) {
          const pid = parseInt(data[i][0], 10);
          if (pid > cachedMax) cachedMax = pid;
        }
      }
      newParentId = cachedMax + 1;
      newChildId = 1;
      scriptProps.setProperty(PROP_KEY_MAX_PARENT, newParentId.toString());
    }

    const formattedParentId = ('000000' + newParentId).slice(-6);
    const formattedChildId = ('000' + newChildId).slice(-3);
    const budgetCode = `${formattedParentId}_${formattedChildId}`;
    
    budgetSheet.appendRow(["'" + budgetCode, "'" + formattedParentId, "'" + formattedChildId, registrationTime, endDate, userEmail]);
    infoSheet.appendRow(["'" + budgetCode, formObject.bu, formObject.department, formObject.account, formObject.pjc, formObject.vendor, formObject.summary, isSlide ? "Yes" : "", registrationTime, userEmail, formObject.amount, formObject.notes, formObject.startDate, formObject.endDate]);

    const targetUrl = formObject.targetUrl;
    const targetSheetName = formObject.targetSheetName;
    if (!targetSheetName) throw new Error("書き込み先のInputシート名が選択されていません。");

    let targetSS = (targetUrl && targetUrl.includes("spreadsheets")) ? SpreadsheetApp.openByUrl(targetUrl) : ss;
    let inputSheet = targetSS.getSheetByName(targetSheetName);
    if (!inputSheet) throw new Error(`対象のスプレッドシート内に「${targetSheetName}」が見つかりません。`);

    const { mainHeaders, dateHeadersInfo } = getCachedHeaders(inputSheet);
    let newRowData = new Array(mainHeaders.length).fill("");
    
    const fill = (kw, val) => {
        const idx = mainHeaders.findIndex(h => h && String(h).includes(kw));
        if (idx !== -1) newRowData[idx] = val;
    };

    fill("MPコード", "'" + budgetCode);               
    fill("グループ名", formObject.department);         
    fill("一般科目", formObject.account);              
    fill("プロジェクト名", formObject.pjc);            
    fill("仕入先", formObject.vendor);                 
    fill("備考1", formObject.summary); 
    fill("備考3", formObject.notes);   

    const monthlyAmounts = formObject.monthlyAmounts || {};
    dateHeadersInfo.forEach((dInfo, colIdx) => {
      if (dInfo.isDate) {
        const dt = new Date(dInfo.time);
        const y = dt.getFullYear();
        const m = ('0' + (dt.getMonth() + 1)).slice(-2);
        const ym = `${y}/${m}`; 
        if (monthlyAmounts[ym] !== undefined && monthlyAmounts[ym] !== "") {
          newRowData[colIdx] = Number(monthlyAmounts[ym]);
        }
      }
    });

    let insertRow = inputSheet.getLastRow() + 1; 
    const maxRows = inputSheet.getMaxRows();
    if (maxRows > 0) {
      const colData = inputSheet.getRange(1, 3, maxRows, 1).getValues();
      for (let i = colData.length - 1; i >= 0; i--) {
        if (colData[i][0] !== "") { insertRow = i + 2; break; }
      }
    }
    
    inputSheet.getRange(insertRow, 1, 1, newRowData.length).setValues([newRowData]);
    SpreadsheetApp.flush();
    return { success: true, message: `登録完了しました。\n新しいMPコード: ${budgetCode}` };
  } catch (e) { return { success: false, message: 'エラーが発生しました: ' + e.message };
  } finally { lock.releaseLock(); }
}
