/**
 * 打刻WebAppの実装
 * 要件1.2, 1.3, 1.4, 1.5に対応
 */

/**
 * HTMLエスケープ関数
 * @param {string} text - エスケープするテキスト
 * @return {string} エスケープされたテキスト
 */
function escapeHtml(text) {
  if (typeof text !== 'string') {
    return String(text);
  }
  
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * CSRFトークンを生成する
 * @param {string} userEmail - ユーザーメールアドレス
 * @return {string} 生成されたCSRFトークン
 */
function generateCsrfToken(userEmail) {
  const timestamp = Date.now();
  const random = Utilities.getUuid();
  const base = `${userEmail}:${timestamp}:${random}`;
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, base, Utilities.Charset.UTF_8));
}

/**
 * CSRFトークンを保存する
 * @param {string} userEmail - ユーザーメールアドレス
 * @param {string} token - CSRFトークン
 * @param {number} expiryMinutes - 有効期限（分）
 */
function saveCsrfToken(userEmail, token, expiryMinutes = 30) {
  try {
    const tokenSheet = getSheet('CSRF_Tokens');
    const expiryTime = new Date(Date.now() + (expiryMinutes * 60 * 1000));
    
    // 最適化: 既存のトークンを効率的に削除
    removeExistingCsrfToken(tokenSheet, userEmail);
    
    // 新しいトークンを追加
    tokenSheet.appendRow([userEmail, token, expiryTime]);
    
    Logger.log(`CSRFトークンを保存しました: ${userEmail}`);
  } catch (error) {
    Logger.log(`CSRFトークン保存エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * CSRFトークンを検証する
 * @param {string} userEmail - ユーザーメールアドレス
 * @param {string} token - 検証するCSRFトークン
 * @return {boolean} トークンが有効かどうか
 */
function validateCsrfToken(userEmail, token) {
  try {
    const tokenSheet = getSheet('CSRF_Tokens');
    
    // 最適化: 特定のユーザーのトークンのみを検索
    const userToken = findCsrfTokenByUser(tokenSheet, userEmail);
    if (!userToken) {
      Logger.log(`CSRFトークンが見つかりません: ${userEmail}`);
      return false;
    }
    
    const storedToken = userToken[1];
    const expiryTime = new Date(userToken[2]);
    const now = new Date();
    
    // 有効期限チェック
    if (now > expiryTime) {
      Logger.log(`CSRFトークンの有効期限が切れています: ${userEmail}`);
      // 期限切れトークンを削除
      removeCsrfTokenByRow(tokenSheet, userToken[3]); // rowIndexを保存
      return false;
    }
    
    // トークン比較
    const isValid = storedToken === token;
    if (isValid) {
      Logger.log(`CSRFトークン検証成功: ${userEmail}`);
    } else {
      Logger.log(`CSRFトークン検証失敗: ${userEmail}`);
    }
    
    return isValid;
  } catch (error) {
    Logger.log(`CSRFトークン検証エラー: ${error.toString()}`);
    return false;
  }
}

/**
 * 期限切れのCSRFトークンをクリーンアップする
 */
function cleanupExpiredCsrfTokens() {
  try {
    const tokenSheet = getSheet('CSRF_Tokens');
    
    // 最適化: 期限切れトークンの効率的な削除
    removeExpiredCsrfTokens(tokenSheet);
    
  } catch (error) {
    Logger.log(`CSRFトークンクリーンアップエラー: ${error.toString()}`);
  }
}

/**
 * 既存のCSRFトークンを効率的に削除する
 * @param {Sheet} tokenSheet - CSRF_Tokensシート
 * @param {string} userEmail - ユーザーメールアドレス
 */
function removeExistingCsrfToken(tokenSheet, userEmail) {
  try {
    const data = tokenSheet.getDataRange().getValues();
    const userRowIndex = data.findIndex(row => row[0] === userEmail);
    if (userRowIndex > 0) { // ヘッダー行をスキップ
      tokenSheet.deleteRow(userRowIndex + 1);
    }
  } catch (error) {
    Logger.log(`removeExistingCsrfToken error: ${error.toString()}`);
  }
}

/**
 * 特定のユーザーのCSRFトークンを検索する
 * @param {Sheet} tokenSheet - CSRF_Tokensシート
 * @param {string} userEmail - ユーザーメールアドレス
 * @return {Array|null} トークン情報 [email, token, expiryTime, rowIndex] または null
 */
function findCsrfTokenByUser(tokenSheet, userEmail) {
  try {
    const data = tokenSheet.getDataRange().getValues();
    const tokens = data.slice(1); // ヘッダー行をスキップ
    
    for (let i = 0; i < tokens.length; i++) {
      if (tokens[i][0] === userEmail) {
        return [...tokens[i], i + 2]; // rowIndexを追加
      }
    }
    
    return null;
  } catch (error) {
    Logger.log(`findCsrfTokenByUser error: ${error.toString()}`);
    return null;
  }
}

/**
 * 指定された行のCSRFトークンを削除する
 * @param {Sheet} tokenSheet - CSRF_Tokensシート
 * @param {number} rowIndex - 削除する行番号
 */
function removeCsrfTokenByRow(tokenSheet, rowIndex) {
  try {
    tokenSheet.deleteRow(rowIndex);
  } catch (error) {
    Logger.log(`removeCsrfTokenByRow error: ${error.toString()}`);
  }
}

/**
 * 期限切れのCSRFトークンを効率的に削除する
 * @param {Sheet} tokenSheet - CSRF_Tokensシート
 */
function removeExpiredCsrfTokens(tokenSheet) {
  try {
    const data = tokenSheet.getDataRange().getValues();
    const tokens = data.slice(1); // ヘッダー行をスキップ
    const now = new Date();
    
    // 期限切れのトークンを削除（後ろから削除してインデックスを維持）
    for (let i = tokens.length - 1; i >= 0; i--) {
      const expiryTime = new Date(tokens[i][2]);
      if (now > expiryTime) {
        tokenSheet.deleteRow(i + 2); // ヘッダー行 + 1ベース
        Logger.log(`期限切れCSRFトークンを削除しました: ${tokens[i][0]}`);
      }
    }
  } catch (error) {
    Logger.log(`removeExpiredCsrfTokens error: ${error.toString()}`);
  }
}

/**
 * CSRF_Tokensシートを作成する
 */
function createCsrfTokensSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(getConfig('SHEET', 'SPREADSHEET_ID'));
    
    // 既存のシートをチェック
    let tokenSheet = spreadsheet.getSheetByName('CSRF_Tokens');
    if (tokenSheet) {
      Logger.log('CSRF_Tokensシートは既に存在します');
      return tokenSheet;
    }
    
    // 新しいシートを作成
    tokenSheet = spreadsheet.insertSheet('CSRF_Tokens');
    
    // ヘッダーを設定
    tokenSheet.getRange(1, 1, 1, 3).setValues([['UserEmail', 'Token', 'ExpiryTime']]);
    
    // ヘッダーのスタイルを設定
    const headerRange = tokenSheet.getRange(1, 1, 1, 3);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f0f0f0');
    
    // 列幅を調整
    tokenSheet.setColumnWidth(1, 200); // UserEmail
    tokenSheet.setColumnWidth(2, 300); // Token
    tokenSheet.setColumnWidth(3, 150); // ExpiryTime
    
    Logger.log('CSRF_Tokensシートを作成しました');
    return tokenSheet;
    
  } catch (error) {
    Logger.log(`CSRF_Tokensシート作成エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 指定された従業員の指定日付の打刻記録を効率的に取得する
 * @param {string} employeeId - 従業員ID
 * @param {Date} date - 対象日付
 * @return {Array} 打刻記録の配列
 */
function getEmployeeEntriesForDate(employeeId, date) {
  try {
    const logSheet = getSheet('Log_Raw');
    
    // 日付文字列を作成（YYYY-MM-DD形式）
    const dateString = Utilities.formatDate(date, 'JST', 'yyyy-MM-dd');
    
    // インデックス列がある場合は高速検索を使用
    if (hasIndexColumn(logSheet)) {
      return getEmployeeEntriesWithIndex(logSheet, employeeId, dateString);
    }
    
    // 従来の方法: 全データ取得後にフィルタリング
    const data = logSheet.getDataRange().getValues();
    const logEntries = data.slice(1); // ヘッダー行をスキップ
    
    return logEntries.filter(row => {
      const entryDate = new Date(row[0]); // タイムスタンプ
      const entryEmployeeId = row[1];     // 社員ID
      
      return entryEmployeeId === employeeId && 
             entryDate.toDateString() === date.toDateString();
    });
    
  } catch (error) {
    Logger.log(`getEmployeeEntriesForDate error: ${error.toString()}`);
    return [];
  }
}

/**
 * インデックス列の存在を確認する
 * @param {Sheet} sheet - 対象シート
 * @return {boolean} インデックス列が存在するかどうか
 */
function hasIndexColumn(sheet) {
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers.includes('Index');
  } catch (error) {
    return false;
  }
}

/**
 * インデックスを使用して従業員の打刻記録を高速取得する
 * @param {Sheet} sheet - Log_Rawシート
 * @param {string} employeeId - 従業員ID
 * @param {string} dateString - 日付文字列（YYYY-MM-DD）
 * @return {Array} 打刻記録の配列
 */
function getEmployeeEntriesWithIndex(sheet, employeeId, dateString) {
  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // インデックス列の位置を取得
    const indexCol = headers.indexOf('Index');
    const employeeIdCol = headers.indexOf('社員ID');
    const timestampCol = headers.indexOf('タイムスタンプ');
    
    if (indexCol === -1 || employeeIdCol === -1 || timestampCol === -1) {
      Logger.log('インデックス列または必要な列が見つかりません');
      return [];
    }
    
    // インデックスを使用して該当する行のみを取得
    const targetIndex = `${employeeId}_${dateString}`;
    const filteredRows = data.slice(1).filter(row => {
      const index = row[indexCol];
      return index === targetIndex;
    });
    
    return filteredRows;
    
  } catch (error) {
    Logger.log(`getEmployeeEntriesWithIndex error: ${error.toString()}`);
    return [];
  }
}

/**
 * Log_Rawシートにインデックス列を追加する
 * @param {Sheet} sheet - Log_Rawシート
 */
function addIndexColumnToLogRaw(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    
    // インデックス列を追加
    sheet.getRange(1, lastCol + 1).setValue('Index');
    
    // 既存データにインデックスを設定
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const employeeIdCol = headers.indexOf('社員ID');
    const timestampCol = headers.indexOf('タイムスタンプ');
    
    if (employeeIdCol === -1 || timestampCol === -1) {
      Logger.log('必要な列が見つかりません');
      return;
    }
    
    const indexValues = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const employeeId = row[employeeIdCol];
      const timestamp = new Date(row[timestampCol]);
      const dateString = Utilities.formatDate(timestamp, 'JST', 'yyyy-MM-dd');
      const index = `${employeeId}_${dateString}`;
      indexValues.push([index]);
    }
    
    if (indexValues.length > 0) {
      sheet.getRange(2, lastCol + 1, indexValues.length, 1).setValues(indexValues);
    }
    
    Logger.log('インデックス列を追加しました');
    
  } catch (error) {
    Logger.log(`addIndexColumnToLogRaw error: ${error.toString()}`);
  }
}

/**
 * 日次サマリーシートを作成・更新する
 * @param {Date} date - 対象日付
 */
function updateDailySummary(date) {
  try {
    const dateString = Utilities.formatDate(date, 'JST', 'yyyy-MM-dd');
    const summarySheetName = `Daily_Summary_${dateString}`;
    
    // サマリーシートを取得または作成
    let summarySheet = getSheet(summarySheetName);
    if (!summarySheet) {
      summarySheet = createDailySummarySheet(summarySheetName);
    }
    
    // 当日の全打刻データを取得
    const logSheet = getSheet('Log_Raw');
    const data = logSheet.getDataRange().getValues();
    const logEntries = data.slice(1);
    
    // 対象日のデータのみをフィルタリング
    const dailyEntries = logEntries.filter(row => {
      const entryDate = new Date(row[0]);
      return entryDate.toDateString() === date.toDateString();
    });
    
    // サマリーシートをクリア
    summarySheet.clear();
    
    // ヘッダーを設定
    const headers = ['従業員ID', '氏名', '出勤時刻', '退勤時刻', '休憩開始時刻', '休憩終了時刻', '勤務時間'];
    summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 従業員別にデータを集計
    const employeeSummary = {};
    dailyEntries.forEach(row => {
      const employeeId = row[1];
      const employeeName = row[2];
      const action = row[3];
      const timestamp = new Date(row[0]);
      
      if (!employeeSummary[employeeId]) {
        employeeSummary[employeeId] = {
          name: employeeName,
          IN: null,
          OUT: null,
          BRK_IN: null,
          BRK_OUT: null
        };
      }
      
      employeeSummary[employeeId][action] = timestamp;
    });
    
    // サマリーシートにデータを書き込み
    const summaryData = [];
    Object.keys(employeeSummary).forEach(employeeId => {
      const summary = employeeSummary[employeeId];
      const workHours = calculateWorkHours(summary.IN, summary.OUT, summary.BRK_IN, summary.BRK_OUT);
      
      summaryData.push([
        employeeId,
        summary.name,
        summary.IN ? Utilities.formatDate(summary.IN, 'JST', 'HH:mm:ss') : '',
        summary.OUT ? Utilities.formatDate(summary.OUT, 'JST', 'HH:mm:ss') : '',
        summary.BRK_IN ? Utilities.formatDate(summary.BRK_IN, 'JST', 'HH:mm:ss') : '',
        summary.BRK_OUT ? Utilities.formatDate(summary.BRK_OUT, 'JST', 'HH:mm:ss') : '',
        workHours
      ]);
    });
    
    if (summaryData.length > 0) {
      summarySheet.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);
    }
    
    Logger.log(`日次サマリーを更新しました: ${dateString}`);
    
  } catch (error) {
    Logger.log(`updateDailySummary error: ${error.toString()}`);
  }
}

/**
 * 日次サマリーシートを作成する
 * @param {string} sheetName - シート名
 * @return {Sheet} 作成されたシート
 */
function createDailySummarySheet(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // ヘッダーを設定
    const headers = ['従業員ID', '氏名', '出勤時刻', '退勤時刻', '休憩開始時刻', '休憩終了時刻', '勤務時間'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のスタイルを設定
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    Logger.log(`日次サマリーシートを作成しました: ${sheetName}`);
    return sheet;
    
  } catch (error) {
    Logger.log(`createDailySummarySheet error: ${error.toString()}`);
    throw error;
  }
}

/**
 * 勤務時間を計算する
 * @param {Date} inTime - 出勤時刻
 * @param {Date} outTime - 退勤時刻
 * @param {Date} brkInTime - 休憩開始時刻
 * @param {Date} brkOutTime - 休憩終了時刻
 * @return {string} 勤務時間（HH:mm形式）
 */
function calculateWorkHours(inTime, outTime, brkInTime, brkOutTime) {
  if (!inTime || !outTime) {
    return '';
  }
  
  let totalMinutes = (outTime - inTime) / (1000 * 60);
  
  // 休憩時間を差し引く
  if (brkInTime && brkOutTime) {
    const breakMinutes = (brkOutTime - brkInTime) / (1000 * 60);
    totalMinutes -= breakMinutes;
  }
  
  const hours = Math.floor(totalMinutes / 60);
  const minutes = Math.floor(totalMinutes % 60);
  
  return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
}

/**
 * X-Frame-Options設定を取得する
 * @return {HtmlService.XFrameOptionsMode} XFrameOptionsMode
 */
function getXFrameOptionsMode() {
  try {
    const mode = getConfig('SECURITY', 'X_FRAME_OPTIONS_MODE');
    
    switch (mode) {
      case 'DENY':
        return HtmlService.XFrameOptionsMode.DENY;
      case 'SAMEORIGIN':
        return HtmlService.XFrameOptionsMode.SAMEORIGIN;
      case 'ALLOWALL':
        return HtmlService.XFrameOptionsMode.ALLOWALL;
      default:
        Logger.log(`無効なX_FRAME_OPTIONS_MODE設定: ${mode}。デフォルトでDENYを使用します。`);
        return HtmlService.XFrameOptionsMode.DENY;
    }
  } catch (error) {
    Logger.log(`XFrameOptionsMode取得エラー: ${error.toString()}。デフォルトでDENYを使用します。`);
    return HtmlService.XFrameOptionsMode.DENY;
  }
}

/**
 * ユーザー認証とアクション検証を実行する共有関数
 * @param {string} action - 打刻アクション (IN/OUT/BRK_IN/BRK_OUT)
 * @param {string} csrfToken - CSRFトークン
 * @param {Object} e - イベントオブジェクト（オプション）
 * @return {Object} 検証結果 {success: boolean, message: string, user: Object|null}
 */
function validateUserAndAction(action, csrfToken, e = null) {
  try {
    Logger.log('validateUserAndAction開始: ' + action);
    
    // 認証チェック（イベントベースの認証を優先）
    let user;
    try {
      if (e) {
        user = authenticateUserFromEvent(e);
      } else {
        user = authenticateUser();
      }
    } catch (error) {
      Logger.log(`認証失敗: ${error.message}`);
      return { 
        success: false, 
        message: '認証に失敗しました',
        user: null 
      };
    }
    
    if (!user) {
      Logger.log('認証失敗');
      return { 
        success: false, 
        message: '認証に失敗しました',
        user: null 
      };
    }
    Logger.log('認証成功: ' + user.name);
    
    // CSRFトークンの検証
    if (!csrfToken || !validateCsrfToken(user.email, csrfToken)) {
      Logger.log('CSRFトークン検証失敗');
      return { 
        success: false, 
        message: 'CSRFトークンが無効です。',
        user: null 
      };
    }
    Logger.log('CSRFトークン検証成功');
    
    // アクション検証
    if (!action || !['IN', 'OUT', 'BRK_IN', 'BRK_OUT'].includes(action)) {
      Logger.log('無効なアクション: ' + action);
      return { 
        success: false, 
        message: '無効なアクションです',
        user: null 
      };
    }
    
    Logger.log('validateUserAndAction成功');
    return { 
      success: true, 
      message: 'OK',
      user: user 
    };
    
  } catch (error) {
    Logger.log('validateUserAndAction error: ' + error.toString());
    return { 
      success: false, 
      message: 'システムエラーが発生しました: ' + error.message,
      user: null 
    };
  }
}

/**
 * WebAppのGETリクエストを処理し、打刻画面のHTMLを提供
 * @param {Object} e - イベントオブジェクト
 * @return {HtmlOutput} 打刻画面のHTML
 */
function doGet(e) {
  return withErrorHandling(() => {
    // 認証チェック（イベントベースの認証を使用）
    let user;
    try {
      user = authenticateUserFromEvent(e);
    } catch (error) {
      Logger.log(`認証失敗: ${error.message}`);
      return createErrorPage('認証に失敗しました。従業員マスタに登録されていません。');
    }
    
    if (!user) {
      return createErrorPage('認証に失敗しました。従業員マスタに登録されていません。');
    }
    
    // CSRFトークンの生成と保存
    const csrfToken = generateCsrfToken(user.email);
    saveCsrfToken(user.email, csrfToken, 30); // 30分間有効
    
    // 期限切れトークンのクリーンアップ
    cleanupExpiredCsrfTokens();
    
    // 打刻画面のHTMLを生成
    const template = HtmlService.createTemplate(getTimeEntryHtml());
    
    template.user = user;
    template.csrfToken = csrfToken;
    // X-Frame-Options設定を取得
    const xFrameOptionsMode = getXFrameOptionsMode();
    
    return template.evaluate()
      .setTitle('出勤管理システム - 打刻')
      .setXFrameOptionsMode(xFrameOptionsMode)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  }, 'WebApp.doGet', 'HIGH', {
    suppressError: true,
    defaultValue: createErrorPage('システムエラーが発生しました。管理者にお問い合わせください。')
  });
}

/**
 * WebAppのPOSTリクエストを処理し、打刻データを受信・処理
 * @param {Object} e - イベントオブジェクト
 * @return {ContentService.TextOutput} JSON形式のレスポンス
 */
function doPost(e) {
  return withErrorHandling(() => {
    Logger.log('doPost開始');
    
    // パラメータの取得
    const action = e.parameter.action;
    const csrfToken = e.parameter.csrf_token;
    
    // ユーザー認証とアクション検証を実行（イベントオブジェクトを渡す）
    const validationResult = validateUserAndAction(action, csrfToken, e);
    
    if (!validationResult.success) {
      Logger.log('バリデーション失敗: ' + validationResult.message);
      return createJsonResponse(false, validationResult.message);
    }
    
    // 打刻データの記録
    Logger.log('打刻記録開始');
    const result = recordTimeEntry(validationResult.user.id, action, validationResult.user);
    Logger.log('打刻記録結果: ' + getAnonymizedTimeEntryResult(result));
    
    return createJsonResponse(result.success, result.message);
    
  }, 'WebApp.doPost', 'HIGH', {
    suppressError: true,
    defaultValue: createJsonResponse(false, 'システムエラーが発生しました。管理者にお問い合わせください。')
  });
}

/**
 * google.script.runから呼び出される打刻処理関数
 * @param {string} action - 打刻アクション (IN/OUT/BRK_IN/BRK_OUT)
 * @return {Object} 処理結果 {success: boolean, message: string}
 */
function processTimeEntry(action, csrfToken) {
  try {
    Logger.log('processTimeEntry開始: ' + action);
    
    // ユーザー認証とアクション検証を実行
    const validationResult = validateUserAndAction(action, csrfToken);
    
    if (!validationResult.success) {
      Logger.log('processTimeEntryバリデーション失敗: ' + validationResult.message);
      return { success: false, message: validationResult.message };
    }
    
    // 打刻データの記録
    const result = recordTimeEntry(validationResult.user.id, action, validationResult.user);
    Logger.log('processTimeEntry結果: ' + getAnonymizedTimeEntryResult(result));
    
    return result;
    
  } catch (error) {
    Logger.log('processTimeEntry error: ' + error.toString());
    return { success: false, message: 'システムエラーが発生しました: ' + error.message };
  }
}

/**
 * 打刻データを記録する
 * @param {string} employeeId - 従業員ID
 * @param {string} action - 打刻アクション (IN/OUT/BRK_IN/BRK_OUT)
 * @param {Object} user - 認証済みユーザー情報
 * @return {Object} 処理結果 {success: boolean, message: string}
 */
function recordTimeEntry(employeeId, action, user) {
  try {
    Logger.log('recordTimeEntry開始: ' + employeeId + ', ' + action);
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    
    // 打刻データの検証
    const validation = validateTimeEntry(employeeId, action, today);
    if (!validation.valid) {
      Logger.log('検証失敗: ' + validation.message);
      return { success: false, message: validation.message };
    }
    
    // Log_Rawシートへの記録
    const logSheet = getSheet('Log_Raw');
    const userAgent = 'Google Apps Script WebApp';
    const clientIP = getClientIP();
    
    const logEntry = [
      now,                    // タイムスタンプ
      employeeId,            // 社員ID
      user.name,             // 氏名
      action,                // Action
      clientIP,              // 端末IP
      `WebApp経由: ${userAgent} (${Session.getActiveUser().getEmail()})` // 備考
    ];
    
    logSheet.appendRow(logEntry);
    
    // インデックス列の自動設定（パフォーマンス最適化）
    try {
      const lastRow = logSheet.getLastRow();
      const dateString = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd');
      const indexValue = `${employeeId}_${dateString}`;
      
      // インデックス列が存在するかチェック
      if (hasIndexColumn(logSheet)) {
        const lastCol = logSheet.getLastColumn();
        logSheet.getRange(lastRow, lastCol).setValue(indexValue);
      }
    } catch (error) {
      Logger.log(`インデックス列設定エラー: ${error.toString()}`);
    }
    
    // 成功メッセージの生成
    const actionNames = {
      'IN': '出勤',
      'OUT': '退勤', 
      'BRK_IN': '休憩開始',
      'BRK_OUT': '休憩終了'
    };
    
    const timeString = Utilities.formatDate(now, 'JST', 'HH:mm:ss');
    const message = `${actionNames[action]}を記録しました (${timeString})`;
    
    Logger.log(`Time entry recorded: ${employeeId} - ${action} at ${now}`);
    
    return { success: true, message: message };
    
  } catch (error) {
    Logger.log('recordTimeEntry error: ' + error.toString());
    throw error;
  }
}

/**
 * 打刻データの検証（重複チェック、整合性チェック）
 * @param {string} employeeId - 従業員ID
 * @param {string} action - 打刻アクション
 * @param {Date} date - 対象日付
 * @return {Object} 検証結果 {valid: boolean, message: string}
 */
function validateTimeEntry(employeeId, action, date) {
  try {
    // 最適化されたデータ取得: 日付基準でフィルタリング
    const todayEntries = getEmployeeEntriesForDate(employeeId, date);
    
    // 重複チェック
    const duplicateEntry = todayEntries.find(row => row[3] === action); // Action列
    if (duplicateEntry) {
      const existingTime = Utilities.formatDate(new Date(duplicateEntry[0]), 'JST', 'HH:mm:ss');
      return {
        valid: false,
        message: `本日は既に${getActionName(action)}が記録されています (${existingTime})`
      };
    }
    
    // 整合性チェック
    const integrityCheck = checkTimeEntryIntegrity(todayEntries, action);
    if (!integrityCheck.valid) {
      return integrityCheck;
    }
    
    return { valid: true, message: 'OK' };
    
  } catch (error) {
    Logger.log('validateTimeEntry error: ' + error.toString());
    return { valid: false, message: 'データ検証中にエラーが発生しました' };
  }
}

/**
 * 打刻の整合性をチェック
 * @param {Array} todayEntries - 当日の打刻記録
 * @param {string} newAction - 新しい打刻アクション
 * @return {Object} チェック結果 {valid: boolean, message: string}
 */
function checkTimeEntryIntegrity(todayEntries, newAction) {
  const actions = todayEntries.map(row => row[3]).sort(); // Action列を取得してソート
  
  switch (newAction) {
    case 'IN':
      // 出勤は最初の打刻である必要がある
      if (actions.length > 0) {
        return { valid: false, message: '既に他の打刻が記録されています。出勤は最初に行ってください。' };
      }
      break;
      
    case 'OUT':
      // 退勤前に出勤が必要
      if (!actions.includes('IN')) {
        return { valid: false, message: '出勤打刻が記録されていません。先に出勤してください。' };
      }
      break;
      
    case 'BRK_IN':
      // 休憩開始前に出勤が必要
      if (!actions.includes('IN')) {
        return { valid: false, message: '出勤打刻が記録されていません。先に出勤してください。' };
      }
      // 既に休憩中でないことを確認
      const brkInCount = actions.filter(a => a === 'BRK_IN').length;
      const brkOutCount = actions.filter(a => a === 'BRK_OUT').length;
      if (brkInCount > brkOutCount) {
        return { valid: false, message: '既に休憩中です。' };
      }
      break;
      
    case 'BRK_OUT':
      // 休憩終了前に休憩開始が必要
      if (!actions.includes('BRK_IN')) {
        return { valid: false, message: '休憩開始が記録されていません。先に休憩開始してください。' };
      }
      // 休憩中であることを確認
      const brkInCount2 = actions.filter(a => a === 'BRK_IN').length;
      const brkOutCount2 = actions.filter(a => a === 'BRK_OUT').length;
      if (brkInCount2 <= brkOutCount2) {
        return { valid: false, message: '休憩中ではありません。先に休憩開始してください。' };
      }
      break;
  }
  
  return { valid: true, message: 'OK' };
}

/**
 * アクション名を日本語で取得
 * @param {string} action - アクションコード
 * @return {string} 日本語のアクション名
 */
function getActionName(action) {
  const actionNames = {
    'IN': '出勤',
    'OUT': '退勤',
    'BRK_IN': '休憩開始', 
    'BRK_OUT': '休憩終了'
  };
  return actionNames[action] || action;
}

/**
 * クライアントIPアドレスを取得（可能な場合）
 * @return {string} IPアドレスまたは'Unknown'
 */
function getClientIP() {
  try {
    // Google Apps ScriptではクライアントIPの直接取得は制限されている
    // 代替として実行環境の情報を記録
    const sessionId = Utilities.getUuid().substring(0, 8);
    const timestamp = Utilities.formatDate(new Date(), 'JST', 'HHmmss');
    return `WebApp-${timestamp}-${sessionId}`;
  } catch (error) {
    return 'Unknown';
  }
}

/**
 * エラーページのHTMLを生成
 * @param {string} errorMessage - エラーメッセージ
 * @return {HtmlOutput} エラーページのHTML
 */
function createErrorPage(errorMessage) {
  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>エラー - 出勤管理システム</title>
  <style>
    body {
      font-family: 'Helvetica Neue', Arial, sans-serif;
      max-width: 600px;
      margin: 50px auto;
      padding: 20px;
      background-color: #f5f5f5;
    }
    .error-container {
      background: white;
      padding: 30px;
      border-radius: 10px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      text-align: center;
    }
    .error-icon {
      font-size: 48px;
      color: #f44336;
      margin-bottom: 20px;
    }
    .error-message {
      color: #721c24;
      background: #f8d7da;
      padding: 15px;
      border-radius: 5px;
      border: 1px solid #f5c6cb;
      margin-bottom: 20px;
    }
    .retry-button {
      background: #2196F3;
      color: white;
      padding: 10px 20px;
      border: none;
      border-radius: 5px;
      cursor: pointer;
      font-size: 16px;
    }
    .retry-button:hover {
      background: #1976D2;
    }
  </style>
</head>
<body>
  <div class="error-container">
    <div class="error-icon">⚠️</div>
    <h2>エラーが発生しました</h2>
    <div class="error-message">${escapeHtml(errorMessage)}</div>
    <button class="retry-button" onclick="window.location.reload()">再試行</button>
  </div>
</body>
</html>
  `;
  
  return HtmlService.createHtmlOutput(html)
    .setTitle('エラー - 出勤管理システム');
}

/**
 * JSON形式のレスポンスを生成
 * @param {boolean} success - 成功フラグ
 * @param {string} message - メッセージ
 * @return {ContentService.TextOutput} JSON形式のレスポンス
 */
function createJsonResponse(success, message) {
  const response = {
    success: success,
    message: message,
    timestamp: new Date().toISOString()
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 打刻画面のHTMLテンプレートを取得
 * @return {string} HTMLテンプレート
 */
function getTimeEntryHtml() {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>出勤管理システム - 打刻</title>
  <style>
    body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f5f5f5; }
    .container { background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
    .header { text-align: center; margin-bottom: 30px; color: #333; }
    .user-info { background: #e8f4fd; padding: 15px; border-radius: 5px; margin-bottom: 30px; text-align: center; }
    .button-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 30px; }
    .time-button { padding: 20px; font-size: 16px; font-weight: bold; border: none; border-radius: 8px; cursor: pointer; }
    .btn-in { background: #4CAF50; color: white; }
    .btn-out { background: #f44336; color: white; }
    .btn-brk-in { background: #FF9800; color: white; }
    .btn-brk-out { background: #2196F3; color: white; }
    .current-time { text-align: center; font-size: 24px; font-weight: bold; color: #333; margin-bottom: 20px; }
    .loading { display: none; text-align: center; margin-top: 20px; }
    .message { margin-top: 20px; padding: 15px; border-radius: 5px; text-align: center; display: none; }
    .success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
    .error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>出勤管理システム</h1>
      <h2>打刻画面</h2>
    </div>
    <div class="user-info">
      <strong><?= escapeHtml(user.name) ?></strong> さん<br>
      <small><?= escapeHtml(user.email) ?></small>
    </div>
    <div class="current-time" id="currentTime"></div>
    <div class="button-grid">
      <button class="time-button btn-in" onclick="submitTimeEntry('IN')">出勤</button>
      <button class="time-button btn-out" onclick="submitTimeEntry('OUT')">退勤</button>
      <button class="time-button btn-brk-in" onclick="submitTimeEntry('BRK_IN')">休憩開始</button>
      <button class="time-button btn-brk-out" onclick="submitTimeEntry('BRK_OUT')">休憩終了</button>
    </div>
    <div class="loading" id="loading">処理中...</div>
    <div class="message" id="message"></div>
    <input type="hidden" id="csrfToken" value="<?= escapeHtml(csrfToken) ?>">
  </div>
  <script>
    function updateCurrentTime() {
      const now = new Date();
      const timeString = now.toLocaleString('ja-JP', {
        year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', second: '2-digit'
      });
      document.getElementById('currentTime').textContent = timeString;
    }
    setInterval(updateCurrentTime, 1000);
    updateCurrentTime();
    
    function submitTimeEntry(action) {
      const loading = document.getElementById('loading');
      const message = document.getElementById('message');
      const buttons = document.querySelectorAll('.time-button');
      
      buttons.forEach(btn => btn.disabled = true);
      loading.style.display = 'block';
      message.style.display = 'none';
      
      // Google Apps Script特有の方法: google.script.runを使用
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler(function(result) {
            loading.style.display = 'none';
            showMessage(result.message, result.success ? 'success' : 'error');
            buttons.forEach(btn => btn.disabled = false);
          })
          .withFailureHandler(function(error) {
            loading.style.display = 'none';
            showMessage('エラーが発生しました: ' + error.message, 'error');
            buttons.forEach(btn => btn.disabled = false);
          })
          .processTimeEntry(action, document.getElementById('csrfToken').value);
      } else {
        // フォールバック: 従来のPOST方式
        const params = new URLSearchParams();
        params.append('action', action);
        params.append('csrf_token', document.getElementById('csrfToken').value);
        
        fetch(window.location.href, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
          body: params.toString()
        })
        .then(response => response.text())
        .then(data => {
          loading.style.display = 'none';
          
          if (data.trim().startsWith('<!doctype') || data.trim().startsWith('<html')) {
            showMessage('サーバーエラー: リクエストが正しく処理されませんでした', 'error');
            return;
          }
          
          try {
            const result = JSON.parse(data);
            showMessage(result.message, result.success ? 'success' : 'error');
          } catch (e) {
            showMessage('レスポンス解析エラー', 'error');
          }
        })
        .catch(error => {
          loading.style.display = 'none';
          showMessage('通信エラー: ' + error.message, 'error');
        })
        .finally(() => {
          buttons.forEach(btn => btn.disabled = false);
        });
      }
    }
    
    function showMessage(text, type) {
      const message = document.getElementById('message');
      message.textContent = text;
      message.className = 'message ' + type;
      message.style.display = 'block';
      setTimeout(() => { message.style.display = 'none'; }, 3000);
    }
  </script>
</body>
</html>`;
}

// ========== テスト用関数 ==========

/**
 * WebApp機能のテスト
 */
function testWebAppFunctions() {
  try {
    Logger.log('=== WebApp機能テスト開始 ===');
    
    // 1. 認証テスト
    Logger.log('1. 認証テスト');
    try {
      const user = authenticateUser();
      Logger.log('✓ 認証成功: ' + user.name + ' (' + user.id + ')');
    } catch (error) {
      Logger.log('✗ 認証失敗: ' + error.message);
    }
    
    // 2. シートアクセステスト
    Logger.log('2. シートアクセステスト');
    try {
      const logSheet = getSheet('Log_Raw');
      Logger.log('✓ Log_Rawシートアクセス成功');
    } catch (error) {
      Logger.log('✗ Log_Rawシートアクセス失敗: ' + error.message);
    }
    
    // 3. 従業員情報取得テスト
    Logger.log('3. 従業員情報取得テスト');
    try {
      const user = authenticateUser();
      if (user) {
        Logger.log('✓ 従業員情報取得成功: ' + user.name);
      } else {
        Logger.log('✗ 従業員情報が見つかりません');
      }
    } catch (error) {
      Logger.log('✗ 従業員情報取得エラー: ' + error.message);
    }
    
    // 4. JSON作成テスト
    Logger.log('4. JSON作成テスト');
    try {
      const jsonResponse = createJsonResponse(true, 'テストメッセージ');
      Logger.log('✓ JSON作成成功');
    } catch (error) {
      Logger.log('✗ JSON作成失敗: ' + error.message);
    }
    
    Logger.log('=== WebApp機能テスト完了 ===');
    
  } catch (error) {
    Logger.log('テスト実行エラー: ' + error.toString());
  }
}

/**
 * 手動で打刻テストを実行
 */
function testTimeEntryManually() {
  try {
    Logger.log('=== 手動打刻テスト開始 ===');
    
    const user = authenticateUser();
    const result = recordTimeEntry(user.id, 'IN', user);
    
    Logger.log('打刻テスト結果: ' + getAnonymizedTimeEntryResult(result));
    Logger.log('=== 手動打刻テスト完了 ===');
    
    return result;
    
  } catch (error) {
    Logger.log('手動打刻テストエラー: ' + error.toString());
    return { success: false, message: error.message };
  }
}

/**
 * CSRF保護機能のテスト
 */
function testCsrfProtection() {
  try {
    Logger.log('=== CSRF保護機能テスト開始 ===');
    
    // 1. CSRF_Tokensシートの作成テスト
    Logger.log('1. CSRF_Tokensシート作成テスト');
    try {
      const tokenSheet = createCsrfTokensSheet();
      Logger.log('✓ CSRF_Tokensシート作成成功');
    } catch (error) {
      Logger.log('✗ CSRF_Tokensシート作成失敗: ' + error.message);
    }
    
    // 2. ユーザー認証テスト
    Logger.log('2. ユーザー認証テスト');
    const user = authenticateUser();
    if (!user) {
      Logger.log('✗ ユーザー認証失敗');
      return;
    }
    Logger.log('✓ ユーザー認証成功: ' + user.email);
    
    // 3. CSRFトークン生成テスト
    Logger.log('3. CSRFトークン生成テスト');
    try {
      const token = generateCsrfToken(user.email);
      Logger.log('✓ CSRFトークン生成成功: ' + token.substring(0, 20) + '...');
    } catch (error) {
      Logger.log('✗ CSRFトークン生成失敗: ' + error.message);
    }
    
    // 4. CSRFトークン保存テスト
    Logger.log('4. CSRFトークン保存テスト');
    try {
      const token = generateCsrfToken(user.email);
      saveCsrfToken(user.email, token, 5); // 5分間有効
      Logger.log('✓ CSRFトークン保存成功');
    } catch (error) {
      Logger.log('✗ CSRFトークン保存失敗: ' + error.message);
    }
    
    // 5. CSRFトークン検証テスト
    Logger.log('5. CSRFトークン検証テスト');
    try {
      const token = generateCsrfToken(user.email);
      saveCsrfToken(user.email, token, 5);
      
      // 有効なトークンの検証
      const isValid = validateCsrfToken(user.email, token);
      Logger.log('✓ 有効なトークン検証: ' + (isValid ? '成功' : '失敗'));
      
      // 無効なトークンの検証
      const isInvalid = validateCsrfToken(user.email, 'invalid_token');
      Logger.log('✓ 無効なトークン検証: ' + (!isInvalid ? '成功' : '失敗'));
      
    } catch (error) {
      Logger.log('✗ CSRFトークン検証失敗: ' + error.message);
    }
    
    // 6. クリーンアップテスト
    Logger.log('6. クリーンアップテスト');
    try {
      cleanupExpiredCsrfTokens();
      Logger.log('✓ クリーンアップ実行成功');
    } catch (error) {
      Logger.log('✗ クリーンアップ失敗: ' + error.message);
    }
    
    Logger.log('=== CSRF保護機能テスト完了 ===');
    
  } catch (error) {
    Logger.log('CSRF保護機能テストエラー: ' + error.toString());
  }
}

/**
 * X-Frame-Options設定のテスト
 */
function testXFrameOptions() {
  try {
    Logger.log('=== X-Frame-Options設定テスト開始 ===');
    
    // 1. 設定取得テスト
    Logger.log('1. X-Frame-Options設定取得テスト');
    try {
      const mode = getXFrameOptionsMode();
      Logger.log('✓ X-Frame-Options設定取得成功: ' + mode);
    } catch (error) {
      Logger.log('✗ X-Frame-Options設定取得失敗: ' + error.message);
    }
    
    // 2. 設定値確認テスト
    Logger.log('2. 設定値確認テスト');
    try {
      const configMode = getConfig('SECURITY', 'X_FRAME_OPTIONS_MODE');
      Logger.log('✓ 設定値確認成功: ' + configMode);
    } catch (error) {
      Logger.log('✗ 設定値確認失敗: ' + error.message);
    }
    
    // 3. 無効な設定値テスト
    Logger.log('3. 無効な設定値テスト');
    try {
      // 一時的に無効な値を設定してテスト
      const originalMode = getConfig('SECURITY', 'X_FRAME_OPTIONS_MODE');
      SECURITY_CONFIG.X_FRAME_OPTIONS_MODE = 'INVALID_MODE';
      
      const mode = getXFrameOptionsMode();
      Logger.log('✓ 無効な設定値処理成功: ' + mode);
      
      // 元の値に戻す
      SECURITY_CONFIG.X_FRAME_OPTIONS_MODE = originalMode;
    } catch (error) {
      Logger.log('✗ 無効な設定値処理失敗: ' + error.message);
    }
    
    Logger.log('=== X-Frame-Options設定テスト完了 ===');
    
  } catch (error) {
    Logger.log('X-Frame-Options設定テストエラー: ' + error.toString());
  }
}

/**
 * validateUserAndAction関数のテスト
 */
function testValidateUserAndAction() {
  try {
    Logger.log('=== validateUserAndAction関数テスト開始 ===');
    
    // 1. 正常なケースのテスト
    Logger.log('1. 正常なケースのテスト');
    try {
      const user = authenticateUser();
      if (!user) {
        Logger.log('✗ ユーザー認証失敗のためテストをスキップ');
        return;
      }
      
      // CSRFトークンを生成
      const csrfToken = generateCsrfToken(user.email);
      saveCsrfToken(user.email, csrfToken, 5); // 5分間有効
      
      const result = validateUserAndAction('IN', csrfToken);
      if (result.success) {
        Logger.log('✓ 正常なケーステスト成功: ' + result.message);
        Logger.log('✓ ユーザー情報取得成功: ' + result.user.name);
      } else {
        Logger.log('✗ 正常なケーステスト失敗: ' + result.message);
      }
    } catch (error) {
      Logger.log('✗ 正常なケーステストエラー: ' + error.message);
    }
    
    // 2. 無効なアクションのテスト
    Logger.log('2. 無効なアクションのテスト');
    try {
      const user = authenticateUser();
      if (!user) {
        Logger.log('✗ ユーザー認証失敗のためテストをスキップ');
        return;
      }
      
      const csrfToken = generateCsrfToken(user.email);
      saveCsrfToken(user.email, csrfToken, 5);
      
      const result = validateUserAndAction('INVALID_ACTION', csrfToken);
      if (!result.success && result.message === '無効なアクションです') {
        Logger.log('✓ 無効なアクションテスト成功: ' + result.message);
      } else {
        Logger.log('✗ 無効なアクションテスト失敗: ' + result.message);
      }
    } catch (error) {
      Logger.log('✗ 無効なアクションテストエラー: ' + error.message);
    }
    
    // 3. 無効なCSRFトークンのテスト
    Logger.log('3. 無効なCSRFトークンのテスト');
    try {
      const result = validateUserAndAction('IN', 'invalid_token');
      if (!result.success && result.message === 'CSRFトークンが無効です。') {
        Logger.log('✓ 無効なCSRFトークンテスト成功: ' + result.message);
      } else {
        Logger.log('✗ 無効なCSRFトークンテスト失敗: ' + result.message);
      }
    } catch (error) {
      Logger.log('✗ 無効なCSRFトークンテストエラー: ' + error.message);
    }
    
    // 4. 空のアクションのテスト
    Logger.log('4. 空のアクションのテスト');
    try {
      const user = authenticateUser();
      if (!user) {
        Logger.log('✗ ユーザー認証失敗のためテストをスキップ');
        return;
      }
      
      const csrfToken = generateCsrfToken(user.email);
      saveCsrfToken(user.email, csrfToken, 5);
      
      const result = validateUserAndAction('', csrfToken);
      if (!result.success && result.message === '無効なアクションです') {
        Logger.log('✓ 空のアクションテスト成功: ' + result.message);
      } else {
        Logger.log('✗ 空のアクションテスト失敗: ' + result.message);
      }
    } catch (error) {
      Logger.log('✗ 空のアクションテストエラー: ' + error.message);
    }
    
    Logger.log('=== validateUserAndAction関数テスト完了 ===');
    
  } catch (error) {
    Logger.log('validateUserAndAction関数テストエラー: ' + error.toString());
  }
}

/**
 * パフォーマンス最適化機能をテストする
 */
function testPerformanceOptimization() {
  try {
    Logger.log('=== パフォーマンス最適化テスト開始 ===');
    
    // 1. インデックス列の追加テスト
    Logger.log('1. インデックス列の追加テスト');
    const logSheet = getSheet('Log_Raw');
    if (logSheet) {
      if (!hasIndexColumn(logSheet)) {
        addIndexColumnToLogRaw(logSheet);
        Logger.log('✓ インデックス列を追加しました');
      } else {
        Logger.log('✓ インデックス列は既に存在します');
      }
    } else {
      Logger.log('✗ Log_Rawシートが見つかりません');
    }
    
    // 2. 最適化されたデータ取得テスト
    Logger.log('2. 最適化されたデータ取得テスト');
    const testUser = authenticateUser();
    if (testUser) {
      const today = new Date();
      const startTime = new Date();
      const entries = getEmployeeEntriesForDate(testUser.id, today);
      const endTime = new Date();
      const executionTime = endTime - startTime;
      
      Logger.log(`✓ データ取得完了: ${entries.length}件, 実行時間: ${executionTime}ms`);
    } else {
      Logger.log('✗ テストユーザーの認証に失敗しました');
    }
    
    // 3. 日次サマリー機能テスト
    Logger.log('3. 日次サマリー機能テスト');
    const today = new Date();
    const startTime2 = new Date();
    updateDailySummary(today);
    const endTime2 = new Date();
    const executionTime2 = endTime2 - startTime2;
    
    Logger.log(`✓ 日次サマリー更新完了, 実行時間: ${executionTime2}ms`);
    
    // 4. CSRFトークン最適化テスト
    Logger.log('4. CSRFトークン最適化テスト');
    const testEmail = 'test@example.com';
    const testToken = 'test-token-' + Date.now();
    
    const startTime3 = new Date();
    saveCsrfToken(testEmail, testToken, 1); // 1分間有効
    const endTime3 = new Date();
    const executionTime3 = endTime3 - startTime3;
    
    Logger.log(`✓ CSRFトークン保存完了, 実行時間: ${executionTime3}ms`);
    
    // 5. クリーンアップテスト
    Logger.log('5. クリーンアップテスト');
    const startTime4 = new Date();
    cleanupExpiredCsrfTokens();
    const endTime4 = new Date();
    const executionTime4 = endTime4 - startTime4;
    
    Logger.log(`✓ クリーンアップ完了, 実行時間: ${executionTime4}ms`);
    
    Logger.log('=== パフォーマンス最適化テスト完了 ===');
    
  } catch (error) {
    Logger.log(`testPerformanceOptimization error: ${error.toString()}`);
  }
}

/**
 * データベース移行計画の概要を表示する
 */
function showDatabaseMigrationPlan() {
  Logger.log('=== データベース移行計画 ===');
  Logger.log('');
  Logger.log('1. 短期計画（3-6ヶ月）:');
  Logger.log('   - BigQueryへの段階的移行');
  Logger.log('   - 日次データの自動同期');
  Logger.log('   - レポート機能のBigQuery活用');
  Logger.log('');
  Logger.log('2. 中期計画（6-12ヶ月）:');
  Logger.log('   - リアルタイムデータ同期');
  Logger.log('   - 高度な分析機能の実装');
  Logger.log('   - パフォーマンス監視の強化');
  Logger.log('');
  Logger.log('3. 長期計画（1年以上）:');
  Logger.log('   - 完全なデータベース移行');
  Logger.log('   - マイクロサービス化');
  Logger.log('   - クラウドネイティブ化');
  Logger.log('');
  Logger.log('現在の最適化により、移行までの間もパフォーマンスを維持できます。');
}

/**
 * 機密情報を除外した打刻結果データを取得
 * ログ出力用に使用
 * @param {Object} result - 打刻処理結果
 * @returns {string} 匿名化された結果データ
 */
function getAnonymizedTimeEntryResult(result) {
  try {
    if (!result || typeof result !== 'object') {
      return 'Invalid result data';
    }
    
    // 機密情報を除外した構造体を作成
    const anonymizedResult = {
      success: result.success || false,
      message: result.message || 'No message',
      hasUserData: result.user ? true : false,
      hasEmployeeId: result.employeeId ? true : false,
      timestamp: result.timestamp ? '***' : 'No timestamp',
      action: result.action || 'No action'
    };
    
    return JSON.stringify(anonymizedResult);
    
  } catch (error) {
    Logger.log('打刻結果匿名化処理エラー: ' + error.toString());
    return 'Error in anonymization';
  }
}