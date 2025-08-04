/**
 * 打刻WebAppの実装
 * 要件1.2, 1.3, 1.4, 1.5に対応
 */

/**
 * WebAppのGETリクエストを処理し、打刻画面のHTMLを提供
 * @param {Object} e - イベントオブジェクト
 * @return {HtmlOutput} 打刻画面のHTML
 */
function doGet(e) {
  return withErrorHandling(() => {
    // 認証チェック
    const user = authenticateUser();
    if (!user) {
      return createErrorPage('認証に失敗しました。従業員マスタに登録されていません。');
    }
    
    // 打刻画面のHTMLを生成
    const template = HtmlService.createTemplate(getTimeEntryHtml());
    
    template.user = user;
    return template.evaluate()
      .setTitle('出勤管理システム - 打刻')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
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
    
    // 認証チェック
    const user = authenticateUser();
    if (!user) {
      Logger.log('認証失敗');
      return createJsonResponse(false, '認証に失敗しました');
    }
    Logger.log('認証成功: ' + user.name);
    
    // パラメータの取得
    const action = e.parameter.action;
    Logger.log('アクション: ' + action);
    if (!action || !['IN', 'OUT', 'BRK_IN', 'BRK_OUT'].includes(action)) {
      Logger.log('無効なアクション');
      return createJsonResponse(false, '無効なアクションです');
    }
    
    // 打刻データの記録
    Logger.log('打刻記録開始');
    const result = recordTimeEntry(user.id, action);
    Logger.log('打刻記録結果: ' + JSON.stringify(result));
    
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
function processTimeEntry(action) {
  try {
    Logger.log('processTimeEntry開始: ' + action);
    
    // 認証チェック
    const user = authenticateUser();
    if (!user) {
      Logger.log('認証失敗');
      return { success: false, message: '認証に失敗しました' };
    }
    
    // アクション検証
    if (!action || !['IN', 'OUT', 'BRK_IN', 'BRK_OUT'].includes(action)) {
      Logger.log('無効なアクション: ' + action);
      return { success: false, message: '無効なアクションです' };
    }
    
    // 打刻データの記録
    const result = recordTimeEntry(user.id, action);
    Logger.log('processTimeEntry結果: ' + JSON.stringify(result));
    
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
 * @return {Object} 処理結果 {success: boolean, message: string}
 */
function recordTimeEntry(employeeId, action) {
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
    
    // 従業員情報の取得
    const employee = getEmployeeInfo(Session.getActiveUser().getEmail());
    if (!employee) {
      Logger.log('従業員情報取得失敗');
      return { success: false, message: '従業員情報が見つかりません' };
    }
    
    // Log_Rawシートへの記録
    const logSheet = getSheet('Log_Raw');
    const userAgent = 'Google Apps Script WebApp';
    const clientIP = getClientIP();
    
    const logEntry = [
      now,                    // タイムスタンプ
      employeeId,            // 社員ID
      employee.name,         // 氏名
      action,                // Action
      clientIP,              // 端末IP
      `WebApp経由: ${userAgent} (${Session.getActiveUser().getEmail()})` // 備考
    ];
    
    logSheet.appendRow(logEntry);
    
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
    const logSheet = getSheet('Log_Raw');
    const data = logSheet.getDataRange().getValues();
    
    // ヘッダー行をスキップ
    const logEntries = data.slice(1);
    
    // 当日の該当従業員の打刻記録を取得
    const todayEntries = logEntries.filter(row => {
      const entryDate = new Date(row[0]); // タイムスタンプ
      const entryEmployeeId = row[1];     // 社員ID
      
      return entryEmployeeId === employeeId && 
             entryDate.toDateString() === date.toDateString();
    });
    
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
    <div class="error-message">${errorMessage}</div>
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
      <strong><?= user.name ?></strong> さん<br>
      <small><?= user.email ?></small>
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
          .processTimeEntry(action);
      } else {
        // フォールバック: 従来のPOST方式
        const params = new URLSearchParams();
        params.append('action', action);
        
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
      const userEmail = Session.getActiveUser().getEmail();
      const employee = getEmployeeInfo(userEmail);
      if (employee) {
        Logger.log('✓ 従業員情報取得成功: ' + employee.name);
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
    const result = recordTimeEntry(user.id, 'IN');
    
    Logger.log('打刻テスト結果: ' + JSON.stringify(result));
    Logger.log('=== 手動打刻テスト完了 ===');
    
    return result;
    
  } catch (error) {
    Logger.log('手動打刻テストエラー: ' + error.toString());
    return { success: false, message: error.message };
  }
}