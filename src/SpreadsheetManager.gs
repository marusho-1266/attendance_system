/**
 * スプレッドシート操作管理
 * TDD実装による段階的な機能追加
 */

/**
 * アクティブなスプレッドシートを取得
 * @return {Spreadsheet} アクティブなスプレッドシート
 */
function getActiveSpreadsheet() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('アクティブなスプレッドシートが見つかりません');
    }
    return spreadsheet;
  } catch (error) {
    throw new Error('スプレッドシートの取得に失敗しました: ' + error.message);
  }
}

/**
 * 指定された名前のシートを取得
 * @param {string} sheetName - シート名
 * @return {Sheet} 指定されたシート
 */
function getSheet(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') {
    throw new Error('Invalid sheet name parameter');
  }
  
  try {
    var spreadsheet = getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('シート "' + sheetName + '" が見つかりません');
    }
    
    return sheet;
  } catch (error) {
    throw new Error('シートの取得に失敗しました: ' + error.message);
  }
}

/**
 * 新しいシートを作成
 * @param {string} sheetName - 作成するシート名
 * @return {Sheet} 作成されたシート
 */
function createSheet(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') {
    throw new Error('Invalid sheet name parameter');
  }
  
  try {
    var spreadsheet = getActiveSpreadsheet();
    
    // 既存のシートをチェック
    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      throw new Error('シート "' + sheetName + '" は既に存在します');
    }
    
    var sheet = spreadsheet.insertSheet(sheetName);
    return sheet;
  } catch (error) {
    throw new Error('シートの作成に失敗しました: ' + error.message);
  }
}

/**
 * シートが存在するかどうかを確認
 * @param {string} sheetName - 確認するシート名
 * @return {boolean} 存在する場合true
 */
function sheetExists(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') {
    return false;
  }
  
  try {
    var spreadsheet = getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);
    return sheet !== null;
  } catch (error) {
    return false;
  }
}

/**
 * シートまたは作成して取得
 * @param {string} sheetName - シート名
 * @return {Sheet} 既存または新規作成されたシート
 */
function getOrCreateSheet(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') {
    throw new Error('Invalid sheet name parameter');
  }
  
  try {
    if (sheetExists(sheetName)) {
      return getSheet(sheetName);
    } else {
      return createSheet(sheetName);
    }
  } catch (error) {
    throw new Error('シートの取得または作成に失敗しました: ' + error.message);
  }
}

/**
 * Master_Employee シートのヘッダーを設定
 * @param {Sheet} sheet - 対象シート
 */
function setupEmployeeSheetHeader(sheet) {
  if (!sheet) {
    throw new Error('Invalid sheet parameter');
  }
  
  var headers = [
    '社員ID',
    '氏名', 
    'Gmail',
    '所属',
    '雇用区分',
    '上司Gmail',
    '基準始業',
    '基準終業'
  ];
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // ヘッダー行のフォーマット設定
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e6e6e6');
  
  return sheet;
}

/**
 * Master_Holiday シートのヘッダーを設定
 * @param {Sheet} sheet - 対象シート
 */
function setupHolidaySheetHeader(sheet) {
  if (!sheet) {
    throw new Error('Invalid sheet parameter');
  }
  
  var headers = [
    '日付',
    '名称',
    '法定休日?',
    '会社休日?'
  ];
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // ヘッダー行のフォーマット設定
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e6e6e6');
  
  return sheet;
}

/**
 * Log_Raw シートのヘッダーを設定
 * @param {Sheet} sheet - 対象シート
 */
function setupLogRawSheetHeader(sheet) {
  if (!sheet) {
    throw new Error('Invalid sheet parameter');
  }
  
  var headers = [
    'タイムスタンプ',
    '社員ID',
    '氏名',
    'Action',
    '端末IP',
    '備考'
  ];
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // ヘッダー行のフォーマット設定
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e6e6e6');
  
  // Log_Raw シートは編集禁止に設定
  sheet.protect().setWarningOnly(true);
  
  return sheet;
}

/**
 * Daily_Summary シートのヘッダーを設定
 * @param {Sheet} sheet - 対象シート
 */
function setupDailySummarySheetHeader(sheet) {
  if (!sheet) {
    throw new Error('Invalid sheet parameter');
  }
  
  var headers = [
    '日付',
    '社員ID',
    '出勤',
    '退勤',
    '休憩',
    '実働',
    '残業',
    '遅刻/早退',
    '承認状態'
  ];
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // ヘッダー行のフォーマット設定
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e6e6e6');
  
  return sheet;
}

/**
 * Monthly_Summary シートのヘッダーを設定
 * @param {Sheet} sheet - 対象シート
 */
function setupMonthlySummarySheetHeader(sheet) {
  if (!sheet) {
    throw new Error('Invalid sheet parameter');
  }
  
  var headers = [
    '年月',
    '社員ID',
    '勤務日数',
    '総労働',
    '残業',
    '有給',
    '備考'
  ];
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // ヘッダー行のフォーマット設定
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e6e6e6');
  
  return sheet;
}

/**
 * Request_Responses シートのヘッダーを設定
 * @param {Sheet} sheet - 対象シート
 */
function setupRequestResponsesSheetHeader(sheet) {
  if (!sheet) {
    throw new Error('Invalid sheet parameter');
  }
  
  var headers = [
    'タイムスタンプ',
    '社員ID',
    '種別(修正/残業/休暇)',
    '詳細',
    '希望値',
    '承認者',
    'Status'
  ];
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // ヘッダー行のフォーマット設定
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e6e6e6');
  
  return sheet;
}

/**
 * すべての必要なシートを初期化
 * TDD段階的実装：6枚のシートを1枚ずつ追加
 */
function initializeAllSheets() {
  try {
    console.log('シート初期化を開始します...');
    
    // 1. Master_Employee シート
    var employeeSheet = getOrCreateSheet(getSheetName('MASTER_EMPLOYEE'));
    setupEmployeeSheetHeader(employeeSheet);
    console.log('✓ Master_Employee シートを初期化しました');
    
    // 2. Master_Holiday シート
    var holidaySheet = getOrCreateSheet(getSheetName('MASTER_HOLIDAY'));
    setupHolidaySheetHeader(holidaySheet);
    console.log('✓ Master_Holiday シートを初期化しました');
    
    // 3. Log_Raw シート
    var logRawSheet = getOrCreateSheet(getSheetName('LOG_RAW'));
    setupLogRawSheetHeader(logRawSheet);
    console.log('✓ Log_Raw シートを初期化しました');
    
    // 4. Daily_Summary シート
    var dailySummarySheet = getOrCreateSheet(getSheetName('DAILY_SUMMARY'));
    setupDailySummarySheetHeader(dailySummarySheet);
    console.log('✓ Daily_Summary シートを初期化しました');
    
    // 5. Monthly_Summary シート
    var monthlySummarySheet = getOrCreateSheet(getSheetName('MONTHLY_SUMMARY'));
    setupMonthlySummarySheetHeader(monthlySummarySheet);
    console.log('✓ Monthly_Summary シートを初期化しました');
    
    // 6. Request_Responses シート
    var requestResponsesSheet = getOrCreateSheet(getSheetName('REQUEST_RESPONSES'));
    setupRequestResponsesSheetHeader(requestResponsesSheet);
    console.log('✓ Request_Responses シートを初期化しました');
    
    console.log('全シートの初期化が完了しました');
    return true;
  } catch (error) {
    console.log('シート初期化エラー: ' + error.message);
    throw error;
  }
}

/**
 * シートの行数を取得
 * @param {string} sheetName - シート名
 * @return {number} 行数
 */
function getSheetRowCount(sheetName) {
  try {
    var sheet = getSheet(sheetName);
    return sheet.getLastRow();
  } catch (error) {
    throw new Error('行数の取得に失敗しました: ' + error.message);
  }
}

/**
 * シートにデータを追加
 * @param {string} sheetName - シート名
 * @param {Array} rowData - 追加するデータ配列
 */
function appendDataToSheet(sheetName, rowData) {
  if (!sheetName || !rowData || !Array.isArray(rowData)) {
    throw new Error('Invalid parameters for appendDataToSheet');
  }
  
  try {
    var sheet = getSheet(sheetName);
    sheet.appendRow(rowData);
  } catch (error) {
    throw new Error('データの追加に失敗しました: ' + error.message);
  }
} 