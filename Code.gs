// ===============================================
// Code.gs v6.9 (日付検索バグ修正版)
// ===============================================

// 設定
const SPREADSHEET_ID = '18n6mElgnIYF1L1wG4q5O9jNQepci4WPHoQ-OKmPYZUA';
const SETTINGS_SHEET_NAME = '設定';

// ★TOP画面（home.html）のURL。URLが変わった際はこちらを修正してください。
const TOP_URL = 'https://script.google.com/a/macros/supership.jp/s/AKfycbwoOYnxicBr54Z50w_GdCCkDyl8MT1XndAEFyf9MQ7-y4IMJeu0j9fjIdaaM37T_0Lq/exec';

// キャッシュするプロパティのキー名
const PROP_KEY_MAX_PARENT = 'CURRENT_MAX_PARENT_ID';

/**
 * スプレッドシートを開いた時に実行される関数
 * 管理メニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MPコード管理')
    .addItem('ID採番の同期 (手動修正後に実行)', 'syncIdCache')
    .addToUi();
}

/**
 * 【管理者用】ID採番の同期（リセット）
 */
function syncIdCache() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheetNames = getSheetNames();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const budgetSheet = ss.getSheetByName(sheetNames.budgetMaster);

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
      if (!isNaN(pid) && pid > maxParentId) {
        maxParentId = pid;
      }
    }

    PropertiesService.getScriptProperties().setProperty(PROP_KEY_MAX_PARENT, maxParentId.toString());
    ui.alert(`同期完了しました。\n現在の最大親ID: ${maxParentId}\n\n次は ${maxParentId + 1} から採番されます。`);
  } catch (e) {
    ui.alert("同期に失敗しました: " + e.message);
  }
}

/**
 * 「設定」シートから各種シート名とURLを取得します。
 */
function getSheetNames() {
  try {
    const settingsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) {
      throw new Error(`「${SETTINGS_SHEET_NAME}」シートが見つかりません。`);
    }
    const values = settingsSheet.getRange('B2:C5').getValues();
    return {
      budgetMaster: values[0][0],
      infoMaster: values[1][0],
      externalInputUrl: values[2][1],
      errorLog: values[3][0]
    };
  } catch (e) {
    console.error(e.message);
    return { budgetMaster: 'MPコードマスタ_raw', infoMaster: '情報マスタ_raw', externalInputUrl: '', errorLog: 'error_log' };
  }
}

// ===============================================
// Webアプリのルーティング
// ===============================================
function doGet(e) {
  let template;
  let output;
  if (e.parameter.page === 'new') {
    template = HtmlService.createTemplateFromFile('index');
    template.topUrl = TOP_URL;
    output = template.evaluate().setTitle('MPコード新規登録フォーム');
  } else if (e.parameter.page === 'update') {
    template = HtmlService.createTemplateFromFile('update');
    template.topUrl = TOP_URL;
    output = template.evaluate().setTitle('情報マスタ更新フォーム');
  } else {
    output = HtmlService.createTemplateFromFile('home').evaluate().setTitle('管理メニュー');
  }
  return output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() { return ScriptApp.getService().getUrl(); }

function getInitialFormData() {
  try {
    const [parentIds, bus, pjcs, accounts, vendors] = [getParentIdSuggestions(), getBuSuggestions(), getPjcSuggestions(), getAccountSuggestions(), getVendorSuggestions()];
    return { parentIds: parentIds, bus: bus, pjcs: pjcs, accounts: accounts, vendors: vendors };
  } catch (e) { return { error: e.message }; }
}

function calculateEndDate(startDate) {
  const startMonth = startDate.getMonth() + 1;
  let targetYear = startDate.getFullYear() + (startMonth >= 4 ? 2 : 1);
  return new Date(targetYear, 2, 31);
}

// ===============================================
// メイン処理
// ===============================================
function processForm(formObject) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
  } catch (e) {
    return { success: false, message: '【アクセス混雑】順番待ちが制限時間を超えました。' };
  }

  try {
    const userEmail = Session.getActiveUser().getEmail();
    const sheetNames = getSheetNames();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const budgetSheet = ss.getSheetByName(sheetNames.budgetMaster);
    const infoSheet = ss.getSheetByName(sheetNames.infoMaster);
    if (!budgetSheet || !infoSheet) throw new Error("ローカルシートが見つかりません。");

    const registrationTime = new Date();
    const endDate = calculateEndDate(registrationTime);

    let newParentId, newChildId;
    let isSlide = formObject.slideParentId && formObject.slideParentId.trim() !== '';

    // --- ID採番ロジック ---
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
    
    // 1. MPコードマスタ
    const budgetMasterRow = ["'" + budgetCode, "'" + formattedParentId, "'" + formattedChildId, registrationTime, endDate, userEmail];
    budgetSheet.appendRow(budgetMasterRow);

    // 2. 情報マスタ (K列: amount, L列: notes, M列: startDate, N列: endDate)
    const infoMasterRow = [
      "'" + budgetCode, formObject.bu, formObject.department, formObject.account,
      formObject.pjc, formObject.vendor, formObject.summary,
      isSlide ? "Yes" : "", registrationTime, userEmail,
      formObject.amount, formObject.notes, formObject.startDate, formObject.endDate
    ];
    infoSheet.appendRow(infoMasterRow);

    SpreadsheetApp.flush(); 
    lock.releaseLock();

    // ★自動通知機能の呼び出し
    try {
      sendCompletionEmail('new', userEmail, budgetCode, formObject);
      postToSlack('new', userEmail, budgetCode, formObject);
    } catch (e) {
      console.error("通知送信処理でエラー: " + e.message);
    }

    return { success: true, message: `登録完了しました。新しいMPコード: ${budgetCode}` };
  } catch (e) {
    console.error('Error in processForm: ' + e.toString());
    return { success: false, message: 'エラーが発生しました: ' + e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// ===============================================
// 更新・検索・サジェスト
// ===============================================

function searchBudgetCode(budgetCode) {
  try {
    const sheetNames = getSheetNames();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetNames.infoMaster);
    const data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();

    // ★追加: GASのDateオブジェクトを安全な文字列に変換する関数
    const formatToYYYYMMDD = (val) => {
      if (!val) return '';
      if (val instanceof Date) {
        const yyyy = val.getFullYear();
        const mm = String(val.getMonth() + 1).padStart(2, '0');
        const dd = String(val.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
      }
      return String(val); // すでに文字列ならそのまま返す
    };

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === budgetCode) {
        return {
          data: { 
            bu: data[i][1], department: data[i][2], account: data[i][3], 
            pjc: data[i][4], vendor: data[i][5], summary: data[i][6],
            amount: data[i][10], notes: data[i][11],
            // ★変更: サーバー側で変換してから渡すことで通信エラーを回避
            startDate: formatToYYYYMMDD(data[i][12]), 
            endDate: formatToYYYYMMDD(data[i][13])
          },
          row: i + 2
        };
      }
    }
    return { error: "指定されたMPコードが見つかりませんでした。" };
  } catch (e) { return { error: "検索中にエラーが発生しました: " + e.message };
  }
}

function updateInfoData(formData, rowNumber) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000);
  } catch (e) { return { success: false, message: '【アクセス混雑】時間をおいて再度お試しください。' };
  }
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const sheetNames = getSheetNames();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetNames.infoMaster);
    const updateTime = new Date();
    
    const newRowData = [
      formData.bu, formData.department, formData.account, formData.pjc, 
      formData.vendor, formData.summary, "", updateTime, userEmail, 
      formData.amount, formData.notes, formData.startDate, formData.endDate
    ];
    sheet.getRange(rowNumber, 2, 1, 13).setValues([newRowData]);
    
    try {
      sendCompletionEmail('update', userEmail, formData.budgetCode, formData);
      postToSlack('update', userEmail, formData.budgetCode, formData);
    } catch (e) {
      console.error("通知送信処理でエラー: " + e.message);
    }

    return { success: true, message: `MPコード ${formData.budgetCode} の情報を更新しました。` };
  } catch (e) { return { success: false, message: "更新中にエラーが発生しました: " + e.message }; } finally { lock.releaseLock();
  }
}

function getUniqueValues(data) { return data.map(row => row[0]).filter((value, index, self) => value && self.indexOf(value) === index); }
function getParentIdSuggestions() { try { const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(getSheetNames().budgetMaster); if (sheet.getLastRow() < 2) return []; const values = sheet.getRange('B2:B' + sheet.getLastRow()).getValues(); return values.map(row => row[0]).filter((value, index, self) => value && self.indexOf(value) === index).sort((a, b) => b - a); } catch (e) { return []; } }
function getBuSuggestions() { try { const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ]BUインポート'); const values = sheet.getRange('E2:E').getValues(); return getUniqueValues(values); } catch (e) { return []; } }
function getPjcSuggestions() { try { const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ] Project Code'); if (sheet.getLastRow() < 2) return []; const values = sheet.getRange('C2:F' + sheet.getLastRow()).getValues(); const suggestions = values.filter(row => row[3] === true).map(row => row[0]); return getUniqueValues(suggestions.map(item => [item])); } catch (e) { return []; } }
function getAccountSuggestions() { try { const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ] 勘定科目'); if (sheet.getLastRow() < 6) return []; const values = sheet.getRange('B6:F' + sheet.getLastRow()).getValues(); const filteredValues = values.filter(row => row[4] === true).map(row => [row[1]]); return getUniqueValues(filteredValues); } catch (e) { return []; } }
function getDepartmentSuggestions(selectedBu) { if (!selectedBu) return []; try { const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ]項目データ'); const headers = sheet.getRange('1:1').getValues()[0]; const buColumnIndex = headers.indexOf(selectedBu); if (buColumnIndex === -1) return []; const columnLetter = String.fromCharCode('A'.charCodeAt(0) + buColumnIndex); const range = sheet.getRange(`${columnLetter}2:${columnLetter}`); const values = range.getValues(); return getUniqueValues(values); } catch (e) { return []; } }
function getVendorSuggestions() { try { const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('[マスタ] 取引先'); if (sheet.getLastRow() < 2) return []; const values = sheet.getRange('A2:A' + sheet.getLastRow()).getValues(); return getUniqueValues(values); } catch (e) { return []; } }
