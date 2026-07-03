// ===============================================
// Code.gs v8.1 (クロススプレッドシート URL取得バグ解消版)
// ===============================================

// ==== フォーム用設定 ====
const SPREADSHEET_ID = '18n6mElgnIYF1L1wG4q5O9jNQepci4WPHoQ-OKmPYZUA';
const SETTINGS_SHEET_NAME = '設定';
const PROP_KEY_MAX_PARENT = 'CURRENT_MAX_PARENT_ID';

// ==== マネージャー(一括ツール)用設定 ====
const INPUT_ALL_SHEET_NAME = 'input_ALL';
const DEPRECIATION_AC_SHEET_NAME = 'Input_減価償却費(既存)';
const DEPRECIATION_INVESTMENT_SHEET_NAME = 'Input_減価償却費(設備投資)';
const ASSET_EXPENSE_SHEET_NAME = 'Input_資産化・損金';
const ASSETIZATION_SHEET_NAME = '資産化比率入力シート';
const PROJECT_MASTER_SHEET_NAME = '[マスタ] Project Code';
const ORG_BU_MASTER_SHEET_NAME = '[マスタ] 組織/BU（FY25）';
const ACCOUNT_MASTER_SHEET_NAME = '[マスタ] 勘定科目';

const MP_LINK_CELL = 'L1';
const MP_SHEET_NAME_CELL = 'L2';
const DATE_HEADER_ROW = 5;
const MAIN_HEADER_ROW = 5;
const DATA_START_ROW = 7;

const DEFAULT_MAX_COL_LIMIT = 'DK';
const SEARCH_DATA_COL_LIMIT = 'L';
const COLUMN_MAPPING = {
  "計上部門": "グループ名",
  "集約科目": "集約科目",
  "勘定科目": "一般科目",
  "対象PJC": "プロジェクト名"
};
const SEARCHABLE_COLUMN_NAMES = [
  "MPコード", "BU", "計上部門", "集約科目", "勘定科目", "対象PJC", "仕入先"
];
const DEPARTMENT_ORDER = [
  "人件費", "業務委託費等", "採用費", "地代家賃等", "サーバ費",
  "広告宣伝費", "交際費・会議費", "旅費交通費", "租税公課", "その他", "設備投資", "減価償却費"
];
const PERSONNEL_ACCOUNTS = new Set([
  '給与手当_月給社員', '給与手当_時給社員', '給与手当_出向社員',
  '残業手当_月給社員', '残業手当_時給社員', '残業手当_出向社員',
  '通勤手当_月給社員', '通勤手当_時給社員', '通勤手当_出向社員',
  '法定福利_月給社員', '法定福利_時給社員', '法定福利_出向社員'
]);
const OUTSOURCING_ACCOUNTS = new Set(['業務委託_開発']);

const DEPRECIATION_TARGET_ACCOUNTS = new Set(['設備投資', ...PERSONNEL_ACCOUNTS, ...OUTSOURCING_ACCOUNTS]);
const ASSET_EXPENSE_TARGET_ACCOUNTS = new Set([...PERSONNEL_ACCOUNTS, ...OUTSOURCING_ACCOUNTS]);

// ===============================================
// UI / Menu / Routing
// ===============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MPコード管理')
    .addItem('ID採番の同期 (手動修正後に実行)', 'syncIdCache')
    .addSeparator()
    .addItem('金額データ検索・更新ツール', 'openWebAppSidebar')
    .addSeparator()
    .addItem('減価償却 再計算ツール', 'openDepreciationPage')
    .addSeparator()
    .addItem('資産化比率入力ツール', 'openAssetizationTool')
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
  let template;
  let title;
  
  const baseUrl = ScriptApp.getService().getUrl();
  
  if (page === 'tool' || page === 'WebApp') {
    template = HtmlService.createTemplateFromFile('WebApp');
    title = '金額データ検索・更新';
  } else if (page === 'depreciation') {
    template = HtmlService.createTemplateFromFile('DepreciationTool');
    title = '減価償却 再計算ツール';
  } else if (page === 'assetization') {
    template = HtmlService.createTemplateFromFile('AssetizationTool');
    title = '資産化比率入力ツール';
  } else if (page === 'new') {
    template = HtmlService.createTemplateFromFile('index');
    title = 'MPコード新規登録フォーム';
  } else if (page === 'update') {
    template = HtmlService.createTemplateFromFile('update');
    title = '情報マスタ更新フォーム';
  } else {
    template = HtmlService.createTemplateFromFile('home');
    title = '販管費管理メニュー';
  }
  
  template.topUrl = baseUrl;
  return template.evaluate().setTitle(title).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function getWebAppUrl() { return ScriptApp.getService().getUrl(); }


// ===============================================
// 💡 司令塔関数: クロススプレッドシート接続
// ===============================================

// ★修正: お客様の元のコード通り、B2:C5 の範囲から正確に取得するように完全復元
function getSheetNames() {
  try {
    const settingsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) throw new Error(`「${SETTINGS_SHEET_NAME}」シートが見つかりません。`);
    
    const values = settingsSheet.getRange('B2:C5').getValues();
    return {
      budgetMaster: values[0][0],
      infoMaster: values[1][0],
      externalInputUrl: values[2][1], // ← ここ(C4セル)に一括ツールのURLがある
      errorLog: values[3][0]
    };
  } catch (e) {
    console.error(e.message);
    return { budgetMaster: 'MPコードマスタ_raw', infoMaster: '情報マスタ_raw', externalInputUrl: '', errorLog: 'error_log' };
  }
}

function getManagerSpreadsheet() {
  var url = getSheetNames().externalInputUrl;
  if (url) {
    try {
      return SpreadsheetApp.openByUrl(url);
    } catch(e) {
      var match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (match) {
        return SpreadsheetApp.openById(match[1]);
      }
    }
  }
  // URLが取れない場合は手前のSSを返す
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}


// ===============================================
// Utils & Helpers
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
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  return y * 100 + m;
}

function safeString(val) {
  if (val instanceof Date) return formatDateFast(val);
  return (val === null || val === undefined) ? "" : String(val);
}

function safeNumber(val) { return (typeof val === 'number' && !isNaN(val)) ? val : 0; }

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

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
  const keys = ["initialData_v61", "initialData_v61_count", "input_all_headers_v61", "mp_settings_v61", "mp_index_v90"];
  try {
      const sheet = getManagerSpreadsheet().getSheetByName(INPUT_ALL_SHEET_NAME);
      if (sheet) {
          const urlVal = sheet.getRange(MP_LINK_CELL).getValue();
          if (urlVal && typeof urlVal === 'string') {
              const match = urlVal.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
              if (match) {
                  const id = match[1];
                  keys.push(`mp_index_v90_${id}`);
              }
          }
      }
  } catch(e) { console.warn("Cache key identification failed: " + e.message); }
  try { cache.removeAll(keys); } catch(e) { console.warn("Cache removeAll failed: " + e.message); }
  return "キャッシュをクリアしました。";
}

function getCachedHeaders(sheet) {
  const cacheKey = "input_all_headers_v61";
  const cached = getCacheLarge(cacheKey);
  if (cached) return cached;
  const lastColLetter = DEFAULT_MAX_COL_LIMIT;
  const mainHeaders = sheet.getRange(`${INPUT_ALL_SHEET_NAME}!A${MAIN_HEADER_ROW}:${lastColLetter}${MAIN_HEADER_ROW}`).getValues()[0].map(h => String(h).trim());
  const dateHeaders = sheet.getRange(`${INPUT_ALL_SHEET_NAME}!A${DATE_HEADER_ROW}:${lastColLetter}${DATE_HEADER_ROW}`).getValues()[0];
  const dateHeadersInfo = dateHeaders.map(d => {
      if (d instanceof Date && !isNaN(d)) return { isDate: true, time: d.getTime() };
      return { isDate: false };
  });
  const result = { mainHeaders, dateHeadersInfo };
  putCacheLarge(cacheKey, result, 1200); 
  return result;
}


// ===============================================
// フォーム側機能 (手前のSSを見に行く)
// ===============================================
function syncIdCache() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const budgetSheet = ss.getSheetByName('MPコードマスタ_raw');
    if (!budgetSheet) throw new Error("MPコードマスタが見つかりません。");

    const lastRow = budgetSheet.getLastRow();
    if (lastRow < 2) {
      PropertiesService.getScriptProperties().setProperty(PROP_KEY_MAX_PARENT, '0');
      ui.alert("データが存在しないため、IDカウンターを 0 にリセットしました。");
      return;
    }
    const data = budgetSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    let maxParentId = 0;
    for (let i = 0; i < data.length; i++) {
      const pid = parseInt(data[i][0], 10);
      if (!isNaN(pid) && pid > maxParentId) { maxParentId = pid; }
    }
    PropertiesService.getScriptProperties().setProperty(PROP_KEY_MAX_PARENT, maxParentId.toString());
    ui.alert(`同期完了しました。\n現在の最大親ID: ${maxParentId}\n\n次は ${maxParentId + 1} から採番されます。`);
  } catch (e) { ui.alert("同期に失敗しました: " + e.message); }
}

function getInitialFormData() {
  try {
    const [parentIds, bus, pjcs, accounts, vendors] = [getParentIdSuggestions(), getBuSuggestions(), getPjcSuggestions(), getAccountSuggestions(), getVendorSuggestions()];
    return { parentIds: parentIds, bus: bus, pjcs: pjcs, accounts: accounts, vendors: vendors };
  } catch (e) { return { error: e.message }; }
}

function getBuSuggestions() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ]BUインポート');
    if(!sheet) return [];
    const values = sheet.getRange('E2:E').getValues();
    return values.map(row => row[0]).filter((value, index, self) => value && self.indexOf(value) === index);
  } catch (e) { return []; }
}

function getPjcSuggestions() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ] Project Code');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const targetColIndex = headers.indexOf('ジョブ名');
    const flagColIndex = headers.indexOf('有効/無効フラグ');
    if (targetColIndex === -1 || flagColIndex === -1) return [];
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const suggestions = values.filter(row => row[flagColIndex] === true).map(row => row[targetColIndex]);
    return suggestions.map(item => String(item).trim()).filter((v, i, self) => v && self.indexOf(v) === i);
  } catch (e) { return []; }
}

function getAccountSuggestions() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ] 勘定科目');
    if (!sheet || sheet.getLastRow() < 6) return [];
    const values = sheet.getRange('B6:F' + sheet.getLastRow()).getValues();
    const filteredValues = values.filter(row => row[4] === true).map(row => [row[1]]);
    return filteredValues.map(row => row[0]).filter((value, index, self) => value && self.indexOf(value) === index);
  } catch (e) { return []; }
}

function getDepartmentSuggestions(selectedBu) {
  if (!selectedBu) return [];
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ]項目データ');
    if(!sheet) return [];
    const headers = sheet.getRange('1:1').getValues()[0];
    const buColumnIndex = headers.indexOf(selectedBu);
    if (buColumnIndex === -1) return [];
    const columnLetter = String.fromCharCode('A'.charCodeAt(0) + buColumnIndex);
    const values = sheet.getRange(`${columnLetter}2:${columnLetter}`).getValues();
    return values.map(row => row[0]).filter((value, index, self) => value && self.indexOf(value) === index);
  } catch (e) { return []; }
}

function getVendorSuggestions() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ] 取引先');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const values = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
    return values.map(row => row[0]).filter((value, index, self) => value && self.indexOf(value) === index);
  } catch (e) { return []; }
}

function getParentIdSuggestions() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('MPコードマスタ_raw');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const values = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();
    return values.map(row => row[0]).filter((value, index, self) => value && self.indexOf(value) === index).sort((a, b) => b - a);
  } catch (e) { return []; }
}

function searchBudgetCode(budgetCode) {
  try {
    const sheetNames = getSheetNames();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetNames.infoMaster);
    const data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();
    const formatToYYYYMMDD = (val) => {
      if (!val) return '';
      if (val instanceof Date) {
        return `${val.getFullYear()}-${String(val.getMonth() + 1).padStart(2, '0')}-${String(val.getDate()).padStart(2, '0')}`;
      }
      return String(val);
    };
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === budgetCode) {
        return {
          row: i + 2,
          data: {
            bu: data[i][1], department: data[i][2], account: data[i][3],
            pjc: data[i][4], vendor: data[i][5], summary: data[i][6],
            amount: data[i][10], notes: data[i][11],
            startDate: formatToYYYYMMDD(data[i][12]),
            endDate: formatToYYYYMMDD(data[i][13])
          }
        };
      }
    }
    return { error: "指定されたMPコードが見つかりませんでした。" };
  } catch (e) { return { error: "検索中にエラーが発生しました: " + e.message }; }
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
      if (!parentFound) return { success: false, message: 'エラー: 指定された親ID ' + parentId + ' が見つかりません。' };
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

    budgetSheet.appendRow(["'" + budgetCode, "='" + formattedParentId, "='" + formattedChildId, registrationTime, endDate, userEmail]);
    infoSheet.appendRow(["'" + budgetCode, formObject.bu, formObject.department, formObject.account, formObject.pjc, formObject.vendor, formObject.summary, isSlide ? "Yes" : "", registrationTime, userEmail, formObject.amount, formObject.notes, formObject.startDate, formObject.endDate]);

    SpreadsheetApp.flush();
    return { success: true, message: `登録完了しました。新しいMPコード: ${budgetCode}` };
  } catch (e) { return { success: false, message: 'エラーが発生しました: ' + e.message }; } finally { lock.releaseLock(); }
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


// ===============================================
// 一括ツール・資産化処理関連 (奥のSSを見に行く)
// ===============================================
function getInitialMasterData() {
  const cacheKey = "initialData_v61";
  const cached = getCacheLarge(cacheKey);
  if (cached) return cached;

  const targetSS = getManagerSpreadsheet();
  const sheet = targetSS.getSheetByName(INPUT_ALL_SHEET_NAME);
  if (!sheet) return { error: 'シートが見つかりません。設定シートのURLを確認してください。' };

  const settingsRange = sheet.getRange("J1:L2").getValues();
  const settings = {
    startDate: settingsRange[0][0] instanceof Date ? formatDateFast(settingsRange[0][0]).slice(0,7) : String(settingsRange[0][0]),
    endDate: settingsRange[1][0] instanceof Date ? formatDateFast(settingsRange[1][0]).slice(0,7) : String(settingsRange[1][0])
  };
  const { mainHeaders, dateHeadersInfo } = getCachedHeaders(sheet);
  const lastRow = sheet.getLastRow();
  let data = [];
  if (lastRow >= DATA_START_ROW) {
    const ssId = targetSS.getId();
    const lastColLetter = SEARCH_DATA_COL_LIMIT;
    const resp = Sheets.Spreadsheets.Values.get(ssId, `'${INPUT_ALL_SHEET_NAME}'!A${DATA_START_ROW}:${lastColLetter}${lastRow}`, {
        valueRenderOption: 'UNFORMATTED_VALUE'
    });
    data = resp.values || [];
  }

  let busList = [];
  try {
      const buSheet = targetSS.getSheetByName(ORG_BU_MASTER_SHEET_NAME);
      if (buSheet) {
          const buLastRow = buSheet.getLastRow();
          if (buLastRow >= 6) {
              const buData = buSheet.getRange(6, 5, buLastRow - 5, 1).getValues();
              busList = buData.flat().map(r => String(r).trim()).filter(Boolean);
          }
      }
  } catch (e) { console.warn("BU Master load failed: " + e.message); }
  
  let subjectsList = [];
  try {
      const accSheet = targetSS.getSheetByName(ACCOUNT_MASTER_SHEET_NAME);
      if (accSheet) {
          const accLastRow = accSheet.getLastRow();
          if (accLastRow >= 2) {
             const accData = accSheet.getRange(2, 2, accLastRow - 1, 5).getValues();
             subjectsList = accData.filter(r => {
                 return r[4] === true || String(r[4]).toUpperCase() === 'TRUE';
             }).map(r => String(r[0]).trim()).filter(Boolean);
          }
      }
  } catch (e) { console.warn("Account Master load failed: " + e.message); }

  const buCol = mainHeaders.indexOf('BU');
  const subjectCol = mainHeaders.indexOf('集約科目');
  const bdCodeCol = mainHeaders.indexOf('MPコード');
  const otherColsMap = {};
  ["勘定科目", "対象PJC", "仕入先"].forEach(h => {
      const sheetName = COLUMN_MAPPING[h] || h;
      const idx = mainHeaders.indexOf(sheetName);
      if (idx !== -1) otherColsMap[h] = idx;
  });
  const bus = new Set(busList);
  const subjects = new Set(subjectsList);
  
  const searchIndex = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const bu = String(row[buCol] || "");
    const subj = String(row[subjectCol] || "");
    const item = { id: i, b: String(row[bdCodeCol] || "").toLowerCase(), u: bu, s: subj, o: {} };
    Object.keys(otherColsMap).forEach(key => {
       const colIdx = otherColsMap[key];
       if (row[colIdx]) item.o[key] = String(row[colIdx]);
    });
    searchIndex.push(item);
  }

  const dates = [];
  dateHeadersInfo.forEach(d => { if (d.isDate) dates.push(formatDateFast(new Date(d.time)).slice(0, 7)); });
  const result = {
    bus: Array.from(bus).sort(),
    departments: Array.from(subjects).sort((a,b) => {
        const ia = DEPARTMENT_ORDER.indexOf(a), ib = DEPARTMENT_ORDER.indexOf(b);
        return (ia !== -1 && ib !== -1) ? ia - ib : a.localeCompare(b);
    }),
    otherHeaders: Object.keys(otherColsMap),
    dateOptions: dates,
    defaultStartDate: settings.startDate,
    defaultEndDate: settings.endDate,
    searchIndex: searchIndex,
    generatedAt: new Date().getTime() 
  };
  putCacheLarge(cacheKey, result, 1200);
  return result;
}

function fetchPageDataForWebApp(rowIds, isDebug, customDateRange) {
  const targetSS = getManagerSpreadsheet();
  const sheet = targetSS.getSheetByName(INPUT_ALL_SHEET_NAME);
  const { mainHeaders, dateHeadersInfo } = getCachedHeaders(sheet);
  const results = [];
  const monthlyHeaders = [];
  const settingsRange = sheet.getRange("J1:L2").getValues();
  const defStart = settingsRange[0][0] instanceof Date ? formatDateFast(settingsRange[0][0]).slice(0,7) : String(settingsRange[0][0]);
  const defEnd = settingsRange[1][0] instanceof Date ? formatDateFast(settingsRange[1][0]).slice(0,7) : String(settingsRange[1][0]);
  const targetStart = customDateRange ? customDateRange.startDate : defStart;
  const targetEnd = customDateRange ? customDateRange.endDate : defEnd;
  dateHeadersInfo.forEach(d => {
    if (d.isDate) {
      const ym = formatDateFast(new Date(d.time)).slice(0, 7);
      if (ym >= targetStart && ym <= targetEnd) monthlyHeaders.push(d.time);
    }
  });
  if (rowIds.length > 0) {
    const sortedIndices = rowIds.sort((a,b) => a - b);
    const groups = getRowGroupsWithTolerance(sortedIndices, 1);
    const ssId = targetSS.getId();
    const lastColLetter = DEFAULT_MAX_COL_LIMIT;
    const ranges = groups.map(g => `'${INPUT_ALL_SHEET_NAME}'!A${g.start + DATA_START_ROW}:${lastColLetter}${g.start + g.numRows + DATA_START_ROW - 1}`);
    const response = Sheets.Spreadsheets.Values.batchGet(ssId, { ranges: ranges, valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'SERIAL_NUMBER' });
    const valueRanges = response.valueRanges || [];
    groups.forEach((g, idx) => {
        const blockValues = valueRanges[idx] ? valueRanges[idx].values : [];
        if (!blockValues) return;
        sortedIndices.forEach(targetIdx => {
            if (targetIdx >= g.start && targetIdx < g.start + g.numRows) {
                const relativeIdx = targetIdx - g.start;
                if (relativeIdx < blockValues.length) {
                    const rowValues = blockValues[relativeIdx];
                    const item = { rowNumber: targetIdx + DATA_START_ROW };
                    mainHeaders.forEach((h, c) => {
                        const key = Object.keys(COLUMN_MAPPING).find(k => COLUMN_MAPPING[k] === h) || h;
                        if (SEARCHABLE_COLUMN_NAMES.includes(key) || ["備考1","備考2","備考3","備考4","備考5"].includes(key)) item[key] = safeString(rowValues[c]);
                    });
                    dateHeadersInfo.forEach((d, c) => { if (d.isDate) item[d.time] = safeNumber(rowValues[c]); });
                    results.push(item);
                }
            }
        });
    });
  }
  return JSON.stringify({ searchResults: results, monthlyHeaders: monthlyHeaders, basicInfo: { readOnly: SEARCHABLE_COLUMN_NAMES, editable: [] }, remarksInfo: { readOnly: ["備考5"], editable: ["備考1", "備考2", "備考3", "備考4"] }, mpDataMap: {}, mpSheetName: '', mpDataError: null });
}

function updateMultipleData(dataArray) {
  if (!dataArray.length) return 'No Data';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('他の人が編集中です。');
  try {
    const sheet = getManagerSpreadsheet().getSheetByName(INPUT_ALL_SHEET_NAME);
    const { mainHeaders, dateHeadersInfo } = getCachedHeaders(sheet);
    const lastColNum = mainHeaders.length;
    const allowedCols = [];
    ["備考1", "備考2", "備考3", "備考4"].forEach(name => {
      const idx = mainHeaders.indexOf(name);
      if (idx !== -1) allowedCols.push(idx);
    });
    const dateColIndices = {};
    dateHeadersInfo.forEach((d, i) => { if (d.isDate) dateColIndices[d.time] = i; });
    const minAllowedCol = Math.min(...allowedCols, ...Object.values(dateColIndices));
    const startCol = minAllowedCol + 1, numCols = lastColNum - minAllowedCol;
    dataArray.sort((a, b) => a.rowNumber - b.rowNumber);
    let currentGroup = [dataArray[0]];
    for(let i=1; i<dataArray.length; i++) {
      if(dataArray[i].rowNumber === dataArray[i-1].rowNumber + 1) currentGroup.push(dataArray[i]);
      else { processUpdateGroup(sheet, currentGroup, startCol, numCols, minAllowedCol, mainHeaders, dateColIndices); currentGroup = [dataArray[i]]; }
    }
    processUpdateGroup(sheet, currentGroup, startCol, numCols, minAllowedCol, mainHeaders, dateColIndices);
    return `${dataArray.length}件 更新完了`;
  } finally { lock.releaseLock(); }
}

function processUpdateGroup(sheet, items, startCol, numCols, minAllowedCol, mainHeaders, dateColIndices) {
  const range = sheet.getRange(items[0].rowNumber, startCol, items.length, numCols), values = range.getValues();
  items.forEach((item, i) => {
    Object.keys(item).forEach(key => {
      if (key === 'rowNumber') return;
      let targetColIdx = !isNaN(key) ? dateColIndices[parseInt(key, 10)] : mainHeaders.indexOf(COLUMN_MAPPING[key] || key);
      if (targetColIdx !== -1) {
        const rel = targetColIdx - minAllowedCol;
        if (rel >= 0 && rel < numCols) values[i][rel] = (!isNaN(key) && item[key] !== "") ? Number(item[key]) : item[key];
      }
    });
  });
  range.setValues(values);
}

function getMpDataOnly(rowNumbers, bdCodes) {
  const res = {};
  if (!rowNumbers || !rowNumbers.length) return { mpDataMap: res, mpSheetName: 'MP', error: null };
  try {
    const sheet = getManagerSpreadsheet().getSheetByName(INPUT_ALL_SHEET_NAME);
    const infoRange = sheet.getRange(MP_LINK_CELL + ':' + MP_SHEET_NAME_CELL).getValues();
    const infoObj = { url: infoRange[0][0], name: String(infoRange[1][0]).trim() };
    const id = infoObj.url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)[1];
    const indexCacheKey = `mp_index_v90_${id}`;
    const mpSheet = SpreadsheetApp.openById(id).getSheetByName(infoObj.name);
    if (!mpSheet) {
        return { mpDataMap: res, mpSheetName: infoObj.name, error: `シート「${infoObj.name}」が見つかりません。` };
    }
    const cBd = 0;
    let bdMapObj = getCacheLarge(indexCacheKey);
    if (!bdMapObj) {
        const allBd = Sheets.Spreadsheets.Values.get(id, `'${infoObj.name}'!${columnToLetter(cBd+1)}${DATA_START_ROW}:${columnToLetter(cBd+1)}${mpSheet.getLastRow()}`).values.flat().map(String);
        bdMapObj = {};
        allBd.forEach((bd, i) => { if(bd) bdMapObj[String(bd).trim().toLowerCase()] = i + DATA_START_ROW; });
        putCacheLarge(indexCacheKey, bdMapObj, 1800);
    }
    
    const targetIndices = [];
    bdCodes.forEach((b, i) => {
        const mpRowIdx = bdMapObj[String(b).trim().toLowerCase()];
        if (mpRowIdx) {
            targetIndices.push(mpRowIdx);
        }
    });
    const { dateHeadersInfo } = getCachedHeaders(sheet);
    const dateColMap = new Map();
    const rawDateHeaders = mpSheet.getRange(DATE_HEADER_ROW, 1, 1, mpSheet.getLastColumn()).getValues()[0];
    rawDateHeaders.forEach((d, idx) => {
        const ym = getYearMonth(d);
        if (ym) dateColMap.set(ym, idx);
    });
    const fetchedRows = new Map();
    if (targetIndices.length > 0) {
        const groups = getRowGroupsWithTolerance([...new Set(targetIndices)].sort((a,b)=>a-b), 200);
        const mpResp = Sheets.Spreadsheets.Values.batchGet(id, { ranges: groups.map(g => `'${infoObj.name}'!A${g.start}:${DEFAULT_MAX_COL_LIMIT}${g.start + g.numRows - 1}`), valueRenderOption: 'UNFORMATTED_VALUE' });
        (mpResp.valueRanges || []).forEach(vr => (vr.values || []).forEach(row => { 
            if (row.length > cBd && row[cBd]) fetchedRows.set(String(row[cBd]).trim().toLowerCase(), row);
        }));
    }

    rowNumbers.forEach((rNum, i) => {
        const bd = String(bdCodes[i]).trim().toLowerCase();
        const row = fetchedRows.get(bd);
        
        if (!row) {
            res[rNum] = "__NOT_FOUND__";
            return;
        }

        const m = {};
        dateHeadersInfo.forEach(dInfo => { 
            if (dInfo.isDate) {
              const ym = getYearMonth(new Date(dInfo.time));
              const colIdx = dateColMap.get(ym);
              if (colIdx !== undefined && colIdx < row.length) {
                  const val = row[colIdx];
                  m[dInfo.time] = (val === "" || val === null || val === undefined) ? null : Number(val);
              } else {
                  m[dInfo.time] = null;
              }
            }
        });
        res[rNum] = m;
    });
    return { mpDataMap: res, mpSheetName: infoObj.name, error: null };
  } catch (e) { return { error: e.message }; }
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

function openDepreciationPage() {
  const url = `${ScriptApp.getService().getUrl()}?page=depreciation`;
  const htmlOutput = HtmlService.createHtmlOutput(`<script>window.open('${url}', '_blank');google.script.host.close();</script>`).setWidth(100).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ツール起動中...');
}

function openAssetizationTool() {
  const url = ScriptApp.getService().getUrl() + '?page=assetization';
  const htmlOutput = HtmlService.createHtmlOutput(`<script>window.open('${url}', '_blank');google.script.host.close();</script>`).setWidth(100).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ツール起動中...');
}

function handleDepreciationAcEdit(e) {
  const r = e.range, s = r.getSheet();
  if (r.getRow() < 2 || r.getColumn() < 2) return;
  const d = s.getDataRange().getValues();
  const h = d[0].map(x => String(x).trim());
  const cBd = h.indexOf('MPコード');
  if (cBd === -1) return;
  const bd = d[r.getRow() - 1][cBd];
  const date = d[0][r.getColumn() - 1];
  if (!bd || !(date instanceof Date)) return;
  const ts = e.source.getSheetByName(DEPRECIATION_INVESTMENT_SHEET_NAME);
  if (!ts) return;
  const td = ts.getDataRange().getValues();
  const th = td[MAIN_HEADER_ROW - 1].map(x => String(x).trim());
  const tDates = td[DATE_HEADER_ROW - 1];
  const tcBd = th.indexOf('MPコード'), tcAc = th.indexOf('一般科目');
  let tRow = -1;
  for (let i = DATA_START_ROW - 1; i < td.length; i++) {
    if (td[i][tcBd] === bd && String(td[i][tcAc]).trim() === '減価償却費') { tRow = i; break; }
  }
  if (tRow === -1) return;
  const tCol = tDates.findIndex(d => d instanceof Date && d.getTime() === date.getTime());
  if (tCol !== -1) ts.getRange(tRow + 1, tCol + 1).setValue(0);
}

function getProjectSuggestions() {
  try {
    const sheet = getManagerSpreadsheet().getSheetByName(PROJECT_MASTER_SHEET_NAME);
    if (!sheet) return [];
    const startRow = 6;
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return [];
    const values = sheet.getRange(startRow, 3, lastRow - startRow + 1, 5).getValues();
    return values.map(row => {
      const name = row[0];
      const isAssetizable = row[4];
      if (!name) return null;
      return { name: name, type: isAssetizable === true ? '資産化対象' : '資産化対象外' };
    }).filter(p => p !== null).sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) { return []; }
}

function getAssetizationFilterOptions() {
  try {
    const sheet = getManagerSpreadsheet().getSheetByName(INPUT_ALL_SHEET_NAME);
    const { mainHeaders } = getCachedHeaders(sheet);
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, mainHeaders.length).getValues();
    const buCol = mainHeaders.indexOf('BU');
    const accountCol = mainHeaders.indexOf('一般科目');
    const targetAccounts = new Set([...PERSONNEL_ACCOUNTS, ...OUTSOURCING_ACCOUNTS]);
    const bus = new Set();
    for (let i = 0; i < data.length; i++) {
      const account = String(data[i][accountCol]).trim();
      if (targetAccounts.has(account) && data[i][buCol]) bus.add(data[i][buCol]);
    }
    return { bus: [...bus].sort() };
  } catch (e) { return { error: e.message }; }
}

function fetchAssetizationData(filters) {
  try {
    const ss = getManagerSpreadsheet();
    const inputAllSheet = ss.getSheetByName(INPUT_ALL_SHEET_NAME);
    const assetizationSheet = ss.getSheetByName(ASSETIZATION_SHEET_NAME);
    if (!inputAllSheet || !assetizationSheet) return { error: 'シートが見つかりません。' };
    const { budgetCode, selectedCategories, selectedBUs } = filters;
    const targetAccounts = new Set([...PERSONNEL_ACCOUNTS, ...OUTSOURCING_ACCOUNTS]);
    const lastRow = inputAllSheet.getLastRow();
    const { mainHeaders } = getCachedHeaders(inputAllSheet);
    const allData = inputAllSheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, mainHeaders.length).getValues();
    const bdCodeCol = mainHeaders.indexOf('MPコード');
    const accountCol = mainHeaders.indexOf('一般科目');
    const buCol = mainHeaders.indexOf('BU');
    const extractedData = new Map();
    for (let i = 0; i < allData.length; i++) {
      const row = allData[i];
      const account = String(row[accountCol]).trim();
      const bdCode = String(row[bdCodeCol]).trim();
      const bu = String(row[buCol]).trim();
      const category = OUTSOURCING_ACCOUNTS.has(account) ? '業務委託' : '人件費';
      if (!bdCode || !targetAccounts.has(account)) continue;
      let isMatch = false;
      if (budgetCode) {
        if (bdCode.toLowerCase().includes(budgetCode.toLowerCase())) isMatch = true;
      } else {
        const categoryMatch = selectedCategories.length > 0 ? selectedCategories.includes(category) : true;
        const buMatch = selectedBUs.length > 0 ? selectedBUs.includes(bu) : true;
        if (categoryMatch && buMatch) isMatch = true;
      }
      if (isMatch) {
        const key = `${bdCode}|${account}`;
        if (!extractedData.has(key)) extractedData.set(key, { bdCode, account, category, projects: [] });
      }
    }
    const ratioData = getAssetizationRatioData();
    extractedData.forEach((data, key) => {
      if (ratioData.ratioMap.has(key)) data.projects = ratioData.ratioMap.get(key);
    });
    return { data: Array.from(extractedData.values()) };
  } catch (e) { return { error: e.message }; }
}

function saveAssetizationData(dataToSave) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('他の処理が実行中です。');
  try {
    const ss = getManagerSpreadsheet();
    const sheet = ss.getSheetByName(ASSETIZATION_SHEET_NAME);
    const projectSuggestions = getProjectSuggestions();
    const projectTypeMap = new Map(projectSuggestions.map(p => [p.name, p.type]));
    const headers = ["MPコード", "一般科目", "分類", "プロジェクト", "比率(%)", "資産化対象/対象外", "最終更新者", "最終更新日時"];
    const updaterEmail = Session.getActiveUser().getEmail();
    const updateTimestamp = new Date();
    const updatedKeys = new Set(dataToSave.map(item => `${item.bdCode}|${item.account}`));
    const existingData = sheet.getDataRange().getValues();
    const preservedRows = [];
    if (existingData.length > 1) {
      for (let i = 1; i < existingData.length; i++) {
        if (!updatedKeys.has(`${existingData[i][0]}|${existingData[i][1]}`)) preservedRows.push(existingData[i]);
      }
    }
    const newRows = [];
    dataToSave.forEach(item => {
      if (item.projects.length > 0) {
        item.projects.forEach(p => {
          newRows.push([item.bdCode, item.account, item.category, p.project, p.ratio, projectTypeMap.get(p.project) || '資産化対象外', updaterEmail, updateTimestamp]);
        });
      } else {
        newRows.push([item.bdCode, item.account, item.category, '', '', '', updaterEmail, updateTimestamp]);
      }
    });
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    const finalData = preservedRows.concat(newRows.filter(r => r[3] !== ''));
    if (finalData.length > 0) sheet.getRange(2, 1, finalData.length, headers.length).setValues(finalData);
    return `データを ${dataToSave.length} 件更新しました。`;
  } catch (e) { throw e; } finally { lock.releaseLock(); }
}

function getDepreciationFilterOptions() {
  try {
    const sheet = getManagerSpreadsheet().getSheetByName(INPUT_ALL_SHEET_NAME);
    const { mainHeaders } = getCachedHeaders(sheet);
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, mainHeaders.length).getValues();
    const buCol = mainHeaders.indexOf('BU');
    const accountCol = mainHeaders.indexOf('一般科目');
    const allRelevantAccounts = new Set([...DEPRECIATION_TARGET_ACCOUNTS, ...ASSET_EXPENSE_TARGET_ACCOUNTS]);
    const bus = new Set(), departments = new Set();
    for (let i = 0; i < data.length; i++) {
      const account = String(data[i][accountCol]).trim();
      if (allRelevantAccounts.has(account)) {
        if (data[i][buCol]) bus.add(data[i][buCol]);
      }
    }
    return { bus: [...bus].sort() };
  } catch (e) { return { error: e.message }; }
}

function getAssetizationRatiosAndTypes() {
  const assetData = getAssetizationRatioData();
  const assetRatioMap = new Map(), expenseRatioMap = new Map();
  if (assetData.error) return { assetRatioMap, expenseRatioMap, error: assetData.error };
  assetData.ratioMap.forEach((projects, key) => {
    let currentAssetRatio = 0;
    projects.forEach(p => {
      if (assetData.typeMap.get(p.project) === '資産化対象') currentAssetRatio += (p.ratio / 100);
    });
    currentAssetRatio = Math.min(1, Math.max(0, currentAssetRatio));
    assetRatioMap.set(key, currentAssetRatio);
    expenseRatioMap.set(key, 1 - currentAssetRatio);
  });
  return { assetRatioMap, expenseRatioMap };
}

function getFilteredInvestmentBdCodes(searchConditions) {
  try {
    const sheet = getManagerSpreadsheet().getSheetByName(INPUT_ALL_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const { mainHeaders } = getCachedHeaders(sheet);
    const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, mainHeaders.length).getValues();
    const bdCodeCol = mainHeaders.indexOf('MPコード');
    const buCol = mainHeaders.indexOf('BU');
    const accountCol = mainHeaders.indexOf('一般科目');
    const { budgetCode, bus, departments, generalAccounts } = searchConditions;
    const matchingItemsMap = new Map();
    const relevantAccounts = new Set([...DEPRECIATION_TARGET_ACCOUNTS, ...ASSET_EXPENSE_TARGET_ACCOUNTS]);

    if (budgetCode) {
      const lower = budgetCode.toLowerCase();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (String(row[bdCodeCol]).toLowerCase().includes(lower) && relevantAccounts.has(String(row[accountCol]).trim())) {
          matchingItemsMap.set(row[bdCodeCol], { bdCode: row[bdCodeCol], originalIndex: i + DATA_START_ROW - 1 });
        }
      }
    } else {
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const acc = String(row[accountCol]).trim();
        if (!relevantAccounts.has(acc) || !row[bdCodeCol]) continue;
        const buMatch = bus.length === 0 || bus.includes(row[buCol]);
        const accMatch = generalAccounts.length === 0 || generalAccounts.includes(acc);
        if (buMatch && accMatch) {
          matchingItemsMap.set(row[bdCodeCol], { bdCode: row[bdCodeCol], originalIndex: i + DATA_START_ROW - 1 });
        }
      }
    }
    const items = Array.from(matchingItemsMap.values());
    items.sort((a, b) => (a.bdCode || '').localeCompare(b.bdCode || ''));
    return { items: items };
  } catch (e) { return { error: e.message }; }
}

function processDepreciationBatch(pageItems, depreciationPeriod) {
  if (!pageItems || pageItems.length === 0) return { depreciationResults: [], assetExpenseResults: [], monthlyHeaders: [] };
  const ss = getManagerSpreadsheet();
  const sheet = ss.getSheetByName(INPUT_ALL_SHEET_NAME);
  const rowIndices = pageItems.map(item => item.originalIndex + 1).sort((a, b) => a - b);
  const minRow = rowIndices[0];
  const maxRow = rowIndices[rowIndices.length - 1];
  const { mainHeaders, dateHeadersInfo } = getCachedHeaders(sheet);
  const blockValues = sheet.getRange(minRow, 1, maxRow - minRow + 1, mainHeaders.length).getValues();
  const pageItemData = pageItems.map(item => {
    const relativeIndex = (item.originalIndex + 1) - minRow;
    return { ...item, sourceRow: blockValues[relativeIndex] || [] };
  }).filter(item => item.sourceRow.length > 0);
  const accountCol = mainHeaders.indexOf('一般科目');
  const ratioData = getAssetizationRatiosAndTypes();
  if (ratioData.error) return { error: ratioData.error };
  const { assetRatioMap, expenseRatioMap } = ratioData;
  const depSheet = ss.getSheetByName(DEPRECIATION_INVESTMENT_SHEET_NAME);
  const exDepMap = new Map();
  if (depSheet && depSheet.getLastRow() >= DATA_START_ROW) {
    const exData = depSheet.getDataRange().getValues();
    const exHeaders = exData[MAIN_HEADER_ROW - 1].map(h => String(h).trim());
    const [cBd, cAc, cR1, cR2] = [exHeaders.indexOf('MPコード'), exHeaders.indexOf('一般科目'), exHeaders.indexOf('備考1'), exHeaders.indexOf('備考2')];
    if (cBd !== -1) {
      for (let i = DATA_START_ROW - 1; i < exData.length; i++) {
        if (exData[i][cBd]) exDepMap.set(`${exData[i][cBd]}|${exData[i][cAc]}`, { updater: exData[i][cR1], updateTimestamp: exData[i][cR2] });
      }
    }
  }

  const dateRange = getDateRange();
  const monthlyHeaders = [];
  dateHeadersInfo.forEach(d => {
      if(d.isDate) {
          const ym = formatDateFast(new Date(d.time)).slice(0, 7);
          if(ym >= dateRange.startDate && ym <= dateRange.endDate) monthlyHeaders.push(d);
      }
  });
  const depResults = [], aeResults = [];

  pageItemData.forEach(item => {
    const row = item.sourceRow;
    const account = String(row[accountCol]).trim();
    const key = `${item.bdCode}|${account}`;
    if (DEPRECIATION_TARGET_ACCOUNTS.has(account)) {
      const depSums = new Map();
      dateHeadersInfo.forEach((d, c) => {
        if (d.isDate) {
            const amt = parseFloat(row[c]) || 0;
            if (amt > 0) {
              const mDep = amt / depreciationPeriod;
              for (let k = 1; k <= depreciationPeriod; k++) {
                const dt = new Date(d.time); dt.setMonth(dt.getMonth() + k);
                const ym = formatDateFast(dt).slice(0, 7);
                depSums.set(ym, (depSums.get(ym) || 0) + mDep);
              }
            }
        }
      });
      const ratio = account === '設備投資' ? 1 : (assetRatioMap.get(key) || 0);
      const mVals = {};
      monthlyHeaders.forEach(d => { mVals[d.time] = (depSums.get(formatDateFast(new Date(d.time)).slice(0, 7)) || 0) * ratio; });
      const dAcc = account === '設備投資' ? '減価償却費' : '減価償却_SW/ﾌﾟﾛﾀﾞｸﾄ';
      const log = exDepMap.get(`${item.bdCode}|${dAcc}`) || {};
      depResults.push({ bdCode: item.bdCode, account: dAcc, originalAccount: account, updater: log.updater, updateTimestamp: log.updateTimestamp, monthlyValues: mVals, originalIndex: item.originalIndex });
    }
    if (ASSET_EXPENSE_TARGET_ACCOUNTS.has(account)) {
      const aVals = {}, eVals = {};
      const aRatio = assetRatioMap.get(key) || 0, eRatio = expenseRatioMap.get(key) || 0;
      monthlyHeaders.forEach(d => {
          const colIdx = dateHeadersInfo.findIndex(info => info.time === d.time);
          const cost = parseFloat(row[colIdx]) || 0;
          aVals[d.time] = cost !== 0 ? cost * aRatio * -1 : 0;
          eVals[d.time] = cost !== 0 ? cost * eRatio : 0;
      });
      let aAcc = '', eAcc = '';
      if (PERSONNEL_ACCOUNTS.has(account)) { aAcc = '資産化_人件費'; eAcc = '損金_人件費'; }
      else { aAcc = '資産化_業務委託費等'; eAcc = '損金_業務委託費'; }
      if (aAcc) aeResults.push({ bdCode: item.bdCode, originalAccount: account, assetAccount: aAcc, expenseAccount: eAcc, assetMonthlyValues: aVals, expenseMonthlyValues: eVals, originalIndex: item.originalIndex });
    }
  });
  return { depreciationResults: depResults, assetExpenseResults: aeResults, monthlyHeaders: monthlyHeaders.map(d => d.time) };
}

function executeDepreciationUpdate(checkedBdCodes, uncheckedBdCodes, defaultDepreciationPeriod) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error('処理中です。');
  try {
    const ss = getManagerSpreadsheet();
    const sheetIn = ss.getSheetByName(INPUT_ALL_SHEET_NAME);
    const sheetDep = ss.getSheetByName(DEPRECIATION_INVESTMENT_SHEET_NAME);
    const sheetAE = ss.getSheetByName(ASSET_EXPENSE_SHEET_NAME) || ss.insertSheet(ASSET_EXPENSE_SHEET_NAME);
    const sheetAc = ss.getSheetByName(DEPRECIATION_AC_SHEET_NAME);
    if (sheetAE.getLastRow() === 0) {
      const h = sheetDep.getRange(1, 1, DATA_START_ROW - 1, sheetDep.getLastColumn()).getValues();
      sheetAE.getRange(1, 1, h.length, h[0].length).setValues(h);
    }

    const { mainHeaders, dateHeadersInfo } = getCachedHeaders(sheetIn);
    const dataIn = sheetIn.getRange(DATA_START_ROW, 1, sheetIn.getLastRow()-DATA_START_ROW+1, mainHeaders.length).getValues();
    
    const cBd = mainHeaders.indexOf('MPコード');
    const cAc = mainHeaders.indexOf('一般科目');
    const cR1 = mainHeaders.indexOf('備考1');
    const cR2 = mainHeaders.indexOf('備考2');
    
    const { assetRatioMap, expenseRatioMap } = getAssetizationRatiosAndTypes();
    const actuals = new Map();
    if (sheetAc.getLastRow() > 1) {
        const dAc = sheetAc.getDataRange().getValues();
        const hAc = dAc[0].map(s=>String(s).trim());
        const cAcBd = hAc.indexOf('MPコード');
        if (cAcBd !== -1) {
            for (let i=1; i<dAc.length; i++) {
                if (dAc[i][cAcBd]) {
                    if (!actuals.has(dAc[i][cAcBd])) actuals.set(dAc[i][cAcBd], new Set());
                    for (let j=1; j<cAcBd; j++) {
                        if (dAc[i][j] !== '' && dAc[0][j] instanceof Date) actuals.get(dAc[i][cAcBd]).add(formatDateFast(dAc[0][j]).slice(0, 7));
                    }
                }
            }
        }
    }

    const mapDep = new Map(), mapAE = new Map();
    [sheetDep, sheetAE].forEach((s, idx) => {
        if (s.getLastRow() >= DATA_START_ROW) {
            const d = s.getRange(DATA_START_ROW, 1, s.getLastRow()-DATA_START_ROW+1, s.getLastColumn()).getValues();
            d.forEach(r => {
                if (r[cBd] && uncheckedBdCodes.includes(r[cBd])) (idx===0 ? mapDep : mapAE).set(`${r[cBd]}|${r[cAc]}`, r);
            });
        }
    });
    const user = Session.getActiveUser().getEmail();
    const time = new Date();
    const checkedSet = new Set(checkedBdCodes);
    for (let i=0; i<dataIn.length; i++) {
        const row = dataIn[i];
        const bd = row[cBd], acc = String(row[cAc]).trim();
        if (!checkedSet.has(bd)) continue;
        const key = `${bd}|${acc}`;
        if (DEPRECIATION_TARGET_ACCOUNTS.has(acc)) {
            const sums = new Map();
            dateHeadersInfo.forEach((d, c) => {
               if(d.isDate) {
                   const amt = parseFloat(row[c]) || 0;
                   if(amt > 0) {
                       const m = amt / defaultDepreciationPeriod;
                       for(let k=1; k<=defaultDepreciationPeriod; k++) {
                           const dt = new Date(d.time); dt.setMonth(dt.getMonth()+k);
                           const ym = formatDateFast(dt).slice(0, 7);
                           sums.set(ym, (sums.get(ym)||0)+m);
                       }
                   }
               }
            });
            const ratio = acc === '設備投資' ? 1 : (assetRatioMap.get(key)||0);
            const dAcc = acc === '設備投資' ? '減価償却費' : '減価償却_SW/ﾌﾟﾛﾀﾞｸﾄ';
            const newRow = row.slice(0, cR1).concat(new Array(row.length - cR1).fill(null)); 
            newRow[cAc] = dAcc;
            newRow[cR1] = user; newRow[cR2] = time;
            
            const acts = actuals.get(bd) || new Set();
            dateHeadersInfo.forEach((d, c) => {
               if(d.isDate) {
                   const ym = formatDateFast(new Date(d.time)).slice(0, 7);
                   newRow[c] = acts.has(ym) ? 0 : (sums.get(ym)||0) * ratio;
               }
            });
            mapDep.set(`${bd}|${dAcc}`, newRow);
        }
        
        if (ASSET_EXPENSE_TARGET_ACCOUNTS.has(acc)) {
            const ar = assetRatioMap.get(key) || 0, er = expenseRatioMap.get(key) || 0;
            let aa = '', ea = '';
            if (PERSONNEL_ACCOUNTS.has(acc)) { aa = '資産化_人件費'; ea = '損金_人件費'; }
            else { aa = '資産化_業務委託費等'; ea = '損金_業務委託費'; }
            if (aa && ar > 0) {
                const nr = row.slice(0, cR1).concat(new Array(row.length - cR1).fill(0));
                nr[cAc] = aa; nr[cR1] = user; nr[cR2] = time;
                dateHeadersInfo.forEach((d, c) => { if(d.isDate) nr[c] = (parseFloat(row[c])||0) * ar * -1; });
                mapAE.set(`${bd}|${aa}`, nr);
            }
            if (ea && er > 0) {
                const nr = row.slice(0, cR1).concat(new Array(row.length - cR1).fill(0));
                nr[cAc] = ea; nr[cR1] = user; nr[cR2] = time;
                dateHeadersInfo.forEach((d, c) => { if(d.isDate) nr[c] = (parseFloat(row[c])||0) * er; });
                mapAE.set(`${bd}|${ea}`, nr);
            }
        }
    }
    
    [ {s: sheetDep, m: mapDep}, {s: sheetAE, m: mapAE} ].forEach(obj => {
        if (obj.s.getLastRow() >= DATA_START_ROW) obj.s.getRange(DATA_START_ROW, 1, obj.s.getLastRow()-DATA_START_ROW+1, obj.s.getLastColumn()).clearContent();
        const rows = Array.from(obj.m.values());
        if (rows.length > 0) {
            const maxCol = Math.max(obj.s.getLastColumn(), ...rows.map(r=>r.length));
            const padded = rows.map(r => [...r, ...new Array(maxCol - r.length).fill(null)]);
            obj.s.getRange(DATA_START_ROW, 1, padded.length, maxCol).setValues(padded);
        }
    });
    return `完了: ${checkedBdCodes.length}件のMPコードを計算・更新しました。`;
  } catch (e) { throw e; } finally { lock.releaseLock(); }
}

function getAssetizationRatioData() {
  const ss = getManagerSpreadsheet();
  const sheet = ss.getSheetByName(ASSETIZATION_SHEET_NAME);
  const ratioMap = new Map();
  const typeMap = new Map();
  if (!sheet) return { ratioMap, typeMap };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ratioMap, typeMap };
  const headers = data[0].map(h => String(h).trim());
  const bdCodeCol = headers.indexOf("MPコード");
  const accountCol = headers.indexOf("一般科目");
  const projectCol = headers.indexOf("プロジェクト");
  const ratioCol = headers.indexOf("比率(%)");
  const assetizableCol = headers.indexOf("資産化対象/対象外");
  if ([bdCodeCol, accountCol, projectCol, ratioCol, assetizableCol].includes(-1)) return { ratioMap, typeMap };
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[bdCodeCol] && row[accountCol] && row[projectCol] && row[ratioCol]) {
      const key = `${row[bdCodeCol]}|${row[accountCol]}`;
      if (!ratioMap.has(key)) ratioMap.set(key, []);
      ratioMap.get(key).push({ project: row[projectCol], ratio: parseFloat(row[ratioCol]) || 0 });
      if (!typeMap.has(row[projectCol])) typeMap.set(row[projectCol], row[assetizableCol]);
    }
  }
  return { ratioMap, typeMap };
}

function findDateIndexByYearMonth(headers, targetDateStr) {
  if (!targetDateStr) return -1;
  const parts = targetDateStr.split(/[/.]/);
  const y = parseInt(parts[0], 10), m = parseInt(parts[1], 10);
  if (isNaN(y) || isNaN(m)) return -1;
  return headers.findIndex(c => c instanceof Date && c.getFullYear() === y && c.getMonth() === m - 1);
}
