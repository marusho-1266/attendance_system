/**
 * WebApp.gs - 出勤管理システムのWebAPIエンドポイント
 * TDD実装 - Refactorフェーズ（品質向上版）
 * 
 * 主要機能:
 * - doGet(e): 打刻画面の表示
 * - doPost(e): 打刻データの受信・処理
 * - generateClockHTML(userInfo): HTML生成
 * - processClock(action, userInfo): 打刻処理
 * 
 * Refactor改善点:
 * - エラーハンドリングの強化
 * - 定数の活用
 * - 関数の分離と可読性向上
 * - セキュリティの向上
 */

// WebApp定数定義
const WEBAPP_CONFIG = {
  DUPLICATE_CHECK_MINUTES: 5,
  HTML_TITLE: '出勤管理システム',
  ERROR_MESSAGES: {
    UNAUTHORIZED: '認証が必要です',
    INVALID_ACTION: '無効なアクションです',
    DUPLICATE_ACTION: '重複した操作です。5分以内に同じ操作が記録されています。',
    SAVE_FAILED: 'データの保存に失敗しました',
    SYSTEM_ERROR: 'システムエラーが発生しました'
  }
};

/**
 * Webアプリケーションの GET リクエスト処理
 * @param {Object} e - イベントオブジェクト
 * @return {HtmlOutput} 打刻画面のHTML
 */
function doGet(e) {
  try {
    console.log('doGet: リクエスト開始', { pathInfo: e?.pathInfo, parameters: e?.parameter });
    
    // 認証チェック
    const userInfo = authenticateWebAppUser();
    // 追加ログ出力
    if (userInfo) {
      console.log('doGet: userInfo', JSON.stringify(userInfo));
    } else {
      console.log('doGet: userInfo is null');
    }
    // 管理者リストのログ出力
    try {
      var managers = (typeof getManagerEmails === 'function') ? getManagerEmails() : [];
      console.log('doGet: 管理者リスト', JSON.stringify(managers));
    } catch (err) {
      console.log('doGet: 管理者リスト取得エラー', err.message);
    }
    
    if (!userInfo || !userInfo.isAuthenticated) {
      console.log('doGet: 認証失敗');
      return createUnauthorizedResponse();
    }
    
    console.log('doGet: 認証成功', { email: userInfo.email, name: userInfo.name });
    
    // 認証済みの場合、打刻画面を生成
    const html = generateClockHTML(userInfo);
    return HtmlService.createHtmlOutput(html).setTitle(WEBAPP_CONFIG.HTML_TITLE);
    
  } catch (error) {
    console.error('doGet: エラー発生', error);
    return createErrorResponse(error);
  }
}

/**
 * Webアプリケーションの POST リクエスト処理
 * @param {Object} e - イベントオブジェクト
 * @return {ContentOutput} JSON レスポンス
 */
function doPost(e) {
  try {
    console.log('doPost: リクエスト開始');
    
    // 認証チェック
    const userInfo = authenticateWebAppUser();
    
    if (!userInfo || !userInfo.isAuthenticated) {
      console.log('doPost: 認証失敗');
      return createJsonResponse({
        success: false,
        message: WEBAPP_CONFIG.ERROR_MESSAGES.UNAUTHORIZED
      });
    }
    
    // リクエストデータの解析
    const requestData = parseRequestData(e);
    console.log('doPost: リクエストデータ解析完了', { action: requestData.action });
    
    // アクションのバリデーション
    const validationResult = validateAction(requestData.action);
    if (!validationResult.isValid) {
      console.log('doPost: バリデーションエラー', validationResult.error);
      return createJsonResponse({
        success: false,
        message: validationResult.error
      });
    }
    
    // 打刻処理の実行
    const result = processClock(requestData.action, userInfo);
    console.log('doPost: 打刻処理完了', { success: result.success, action: requestData.action });
    
    return createJsonResponse(result);
    
  } catch (error) {
    console.error('doPost: エラー発生', error);
    return createJsonResponse({
      success: false,
      message: WEBAPP_CONFIG.ERROR_MESSAGES.SYSTEM_ERROR,
      error: error.message
    });
  }
}

/**
 * WebApp用認証処理（認証チェックの分離）
 * @return {Object|null} ユーザー情報または null
 */
function authenticateWebAppUser() {
  try {
    var activeUser = Session.getActiveUser();
    var email = (activeUser && activeUser.getEmail) ? activeUser.getEmail() : null;
    if (!email) return { isAuthenticated: false, email: null };
    var isAuthenticated = authenticateUser(email);
    var employeeInfo = isAuthenticated ? getCachedEmployee(email) : null;
    return {
      isAuthenticated: isAuthenticated,
      email: email,
      name: employeeInfo ? employeeInfo.name : undefined,
      employeeId: employeeInfo ? employeeInfo.employeeId : undefined
    };
  } catch (error) {
    console.error('WebApp認証エラー:', error);
    return { isAuthenticated: false, email: null };
  }
}

/**
 * リクエストデータの解析
 * @param {Object} e - イベントオブジェクト
 * @return {Object} 解析されたリクエストデータ
 */
function parseRequestData(e) {
  let requestData = {};
  
  if (e.postData && e.postData.contents) {
    // JSON形式のデータ
    try {
      requestData = JSON.parse(e.postData.contents);
    } catch (parseError) {
      console.warn('JSON解析失敗、フォーム形式で再試行');
      requestData = e.parameter || {};
    }
  } else if (e.parameter) {
    // フォーム形式のデータ
    requestData = e.parameter;
  }
  
  return requestData;
}

/**
 * アクションのバリデーション
 * @param {string} action - バリデーション対象のアクション
 * @return {Object} バリデーション結果
 */
function validateAction(action) {
  const validActions = ['IN', 'OUT', 'BRK_IN', 'BRK_OUT'];
  
  if (!action) {
    return {
      isValid: false,
      error: 'アクションが指定されていません'
    };
  }
  
  if (!validActions.includes(action)) {
    return {
      isValid: false,
      error: WEBAPP_CONFIG.ERROR_MESSAGES.INVALID_ACTION
    };
  }
  
  return {
    isValid: true,
    error: null
  };
}

/**
 * 未認証レスポンスの生成
 * @return {HtmlOutput} 未認証HTML
 */
function createUnauthorizedResponse() {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>認証エラー</title>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        body { 
          font-family: -apple-system, BlinkMacSystemFont, sans-serif; 
          text-align: center; 
          padding: 50px; 
          background: #f5f5f5; 
        }
        .error-container { 
          background: white; 
          padding: 40px; 
          border-radius: 8px; 
          box-shadow: 0 2px 10px rgba(0,0,0,0.1); 
          max-width: 400px; 
          margin: 0 auto; 
        }
        h1 { color: #d32f2f; margin-bottom: 20px; }
        p { color: #666; line-height: 1.5; }
      </style>
    </head>
    <body>
      <div class="error-container">
        <h1>🔐 認証が必要です</h1>
        <p>システムにアクセスするには、認証が必要です。</p>
        <p>管理者にお問い合わせください。</p>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html).setTitle('認証エラー');
}

/**
 * エラーレスポンスの生成
 * @param {Error} error - エラーオブジェクト
 * @return {HtmlOutput} エラーHTML
 */
function createErrorResponse(error) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>システムエラー</title>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        body { 
          font-family: -apple-system, BlinkMacSystemFont, sans-serif; 
          text-align: center; 
          padding: 50px; 
          background: #f5f5f5; 
        }
        .error-container { 
          background: white; 
          padding: 40px; 
          border-radius: 8px; 
          box-shadow: 0 2px 10px rgba(0,0,0,0.1); 
          max-width: 500px; 
          margin: 0 auto; 
        }
        h1 { color: #d32f2f; margin-bottom: 20px; }
        p { color: #666; line-height: 1.5; }
        .error-detail { 
          background: #fafafa; 
          padding: 15px; 
          border-radius: 4px; 
          margin-top: 20px; 
          font-family: monospace; 
          text-align: left; 
        }
      </style>
    </head>
    <body>
      <div class="error-container">
        <h1>⚠️ システムエラー</h1>
        <p>システムエラーが発生しました。</p>
        <p>しばらく時間をおいてから再度お試しください。</p>
        <div class="error-detail">
          エラー詳細: ${error.message}
        </div>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html).setTitle('システムエラー');
}

/**
 * JSONレスポンスの生成
 * @param {Object} data - レスポンスデータ
 * @return {ContentOutput} JSON形式のレスポンス
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 打刻画面のHTML生成（リファクタリング版）
 * @param {Object} userInfo - ユーザー情報
 * @return {string} HTML文字列
 */
function generateClockHTML(userInfo) {
  // 現在時刻の取得
  const now = new Date();
  const currentTime = formatTime(now);
  const currentDate = formatDate(now);
  
  // HTMLテンプレートの生成
  return createHTMLTemplate({
    userInfo,
    currentTime,
    currentDate,
    title: WEBAPP_CONFIG.HTML_TITLE
  });
}

/**
 * HTMLテンプレートの生成（分離されたHTML生成ロジック）
 * @param {Object} data - テンプレート用データ
 * @return {string} HTML文字列
 */
function createHTMLTemplate(data) {
  const { userInfo, currentTime, currentDate, title } = data;
  
  return `
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title}</title>
  ${getHTMLStyles()}
</head>
<body>
  <div class="container">
    ${getHeaderHTML(userInfo)}
    ${getCurrentTimeHTML(currentTime, currentDate)}
    ${getButtonsHTML()}
    ${getStatusHTML()}
  </div>
  ${getJavaScriptCode()}
</body>
</html>
  `;
}

/**
 * HTMLスタイルの取得
 * @return {string} CSSスタイル
 */
function getHTMLStyles() {
  return `
  <style>
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    
    .container {
      background: white;
      border-radius: 12px;
      box-shadow: 0 15px 35px rgba(0, 0, 0, 0.1);
      padding: 40px;
      max-width: 500px;
      width: 100%;
      text-align: center;
    }
    
    .header {
      margin-bottom: 30px;
    }
    
    .title {
      color: #333;
      font-size: 28px;
      font-weight: 600;
      margin-bottom: 10px;
    }
    
    .user-info {
      color: #666;
      font-size: 16px;
      margin-bottom: 5px;
    }
    
    .current-time {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 20px;
      margin-bottom: 30px;
    }
    
    .time-display {
      font-size: 24px;
      font-weight: 700;
      color: #333;
      margin-bottom: 5px;
    }
    
    .date-display {
      font-size: 16px;
      color: #666;
    }
    
    .buttons {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 15px;
      margin-bottom: 20px;
    }
    
    .btn {
      background: #4CAF50;
      color: white;
      border: none;
      border-radius: 8px;
      padding: 15px 20px;
      font-size: 16px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.3s ease;
    }
    
    .btn:hover {
      transform: translateY(-2px);
      box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    }
    
    .btn:disabled {
      opacity: 0.6;
      cursor: not-allowed;
      transform: none;
    }
    
    .btn.out {
      background: #f44336;
    }
    
    .btn.break-in {
      background: #ff9800;
    }
    
    .btn.break-out {
      background: #2196F3;
    }
    
    .status {
      margin-top: 20px;
      padding: 15px;
      border-radius: 8px;
      display: none;
    }
    
    .status.success {
      background: #d4edda;
      color: #155724;
      border: 1px solid #c3e6cb;
    }
    
    .status.error {
      background: #f8d7da;
      color: #721c24;
      border: 1px solid #f5c6cb;
    }
  </style>
  `;
}

/**
 * ヘッダーHTMLの取得
 * @param {Object} userInfo - ユーザー情報
 * @return {string} ヘッダーHTML
 */
function getHeaderHTML(userInfo) {
  return `
    <div class="header">
      <h1 class="title">${WEBAPP_CONFIG.HTML_TITLE}</h1>
      <div class="user-info">ようこそ、${userInfo.name} さん</div>
      <div class="user-info">社員ID: ${userInfo.employeeId || 'N/A'}</div>
    </div>
  `;
}

/**
 * 現在時刻表示HTMLの取得
 * @param {string} currentTime - 現在時刻
 * @param {string} currentDate - 現在日付
 * @return {string} 時刻表示HTML
 */
function getCurrentTimeHTML(currentTime, currentDate) {
  return `
    <div class="current-time">
      <div class="time-display" id="currentTime">${currentTime}</div>
      <div class="date-display">${currentDate}</div>
    </div>
  `;
}

/**
 * ボタンHTMLの取得
 * @return {string} ボタンHTML
 */
function getButtonsHTML() {
  return `
    <div class="buttons">
      <button class="btn" onclick="clockAction('IN')">出勤</button>
      <button class="btn out" onclick="clockAction('OUT')">退勤</button>
      <button class="btn break-in" onclick="clockAction('BRK_IN')">休憩開始</button>
      <button class="btn break-out" onclick="clockAction('BRK_OUT')">休憩終了</button>
    </div>
  `;
}

/**
 * ステータス表示HTMLの取得
 * @return {string} ステータスHTML
 */
function getStatusHTML() {
  return `<div id="status" class="status"></div>`;
}

/**
 * JavaScriptコードの取得
 * @return {string} JavaScript
 */
function getJavaScriptCode() {
  return `
  <script>
    // 現在時刻の更新
    function updateTime() {
      const now = new Date();
      const timeString = now.toLocaleTimeString('ja-JP', {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      });
      document.getElementById('currentTime').textContent = timeString;
    }
    
    // 1秒ごとに時刻を更新
    setInterval(updateTime, 1000);
    
    // 打刻アクション（改善版）
    function clockAction(action) {
      const statusDiv = document.getElementById('status');
      const buttons = document.querySelectorAll('.btn');
      
      // ボタンを無効化（重複送信防止）
      buttons.forEach(btn => btn.disabled = true);
      statusDiv.style.display = 'none';
      
      // データの準備
      const data = {
        action: action,
        timestamp: new Date().toISOString()
      };
      
      // POSTリクエストの送信
      fetch(window.location.href, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(data)
      })
      .then(response => response.json())
      .then(result => {
        statusDiv.style.display = 'block';
        if (result.success) {
          statusDiv.className = 'status success';
          statusDiv.textContent = getActionMessage(action) + 'を記録しました。';
        } else {
          statusDiv.className = 'status error';
          statusDiv.textContent = 'エラー: ' + result.message;
        }
      })
      .catch(error => {
        statusDiv.style.display = 'block';
        statusDiv.className = 'status error';
        statusDiv.textContent = 'エラーが発生しました: ' + error.message;
      })
      .finally(() => {
        // ボタンを再有効化
        setTimeout(() => {
          buttons.forEach(btn => btn.disabled = false);
        }, 2000);
      });
    }
    
    function getActionMessage(action) {
      const messages = {
        'IN': '出勤',
        'OUT': '退勤',
        'BRK_IN': '休憩開始',
        'BRK_OUT': '休憩終了'
      };
      return messages[action] || 'アクション';
    }
  </script>
  `;
}

/**
 * 打刻処理の実行（リファクタリング版）
 * @param {string} action - アクション（IN/OUT/BRK_IN/BRK_OUT）
 * @param {Object} userInfo - ユーザー情報
 * @return {Object} 処理結果
 */
function processClock(action, userInfo) {
  try {
    console.log('processClock: 処理開始', { action, email: userInfo.email });
    
    // Log_Rawシートの取得
    const logRawSheet = getOrCreateSheet(getSheetName('LOG_RAW'));
    
    // 現在時刻の取得
    const now = new Date();
    const timestamp = now.toISOString();
    
    // 重複チェック
    const duplicateCheck = checkDuplicateAction(logRawSheet, action, userInfo.email, now);
    if (!duplicateCheck.isValid) {
      console.log('processClock: 重複チェックエラー', duplicateCheck.error);
      return {
        success: false,
        message: duplicateCheck.error
      };
    }
    
    // データの準備と保存
    const saveResult = saveClockData(logRawSheet, action, userInfo, now, timestamp);
    
    console.log('processClock: 処理完了', { success: saveResult.success, action });
    return saveResult;
    
  } catch (error) {
    console.error('processClock: エラー発生', error);
    return {
      success: false,
      message: WEBAPP_CONFIG.ERROR_MESSAGES.SAVE_FAILED,
      error: error.message
    };
  }
}

/**
 * 重複アクションのチェック（全行走査版）
 * @param {Sheet} logRawSheet - Log_Rawシート
 * @param {string} action - アクション
 * @param {string} email - メールアドレス
 * @param {Date} now - 現在時刻
 * @return {Object} チェック結果
 */
function checkDuplicateAction(logRawSheet, action, email, now) {
  try {
    const data = logRawSheet.getDataRange().getValues();
    if (data.length <= 1) {
      // ヘッダー行のみの場合は重複なし
      return { isValid: true };
    }
    // 逆順で全行を走査（新しい順）
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const [timestamp, rowEmail, , , rowAction] = row;
      if (rowEmail === email && rowAction === action) {
        const timeDiff = (now.getTime() - new Date(timestamp).getTime()) / (1000 * 60);
        // 同一日かつ5分以内
        if (formatDate(now) === formatDate(new Date(timestamp)) && timeDiff < WEBAPP_CONFIG.DUPLICATE_CHECK_MINUTES) {
          return {
            isValid: false,
            error: WEBAPP_CONFIG.ERROR_MESSAGES.DUPLICATE_ACTION
          };
        }
      }
    }
    return { isValid: true };
  } catch (error) {
    console.warn('重複チェックエラー（処理継続）:', error);
    return { isValid: true }; // エラーが発生した場合は処理を継続
  }
}

/**
 * 打刻データの保存
 * @param {Sheet} logRawSheet - Log_Rawシート
 * @param {string} action - アクション
 * @param {Object} userInfo - ユーザー情報
 * @param {Date} now - 現在時刻
 * @param {string} timestamp - タイムスタンプ
 * @return {Object} 保存結果
 */
function saveClockData(logRawSheet, action, userInfo, now, timestamp) {
  try {
    // データの準備
    const rowData = [
      timestamp,                     // A: Timestamp
      userInfo.email,               // B: Email
      userInfo.name,                // C: Name
      userInfo.employeeId || 'N/A', // D: Employee_ID
      action,                       // E: Action
      formatDate(now),              // F: Date
      formatTime(now),              // G: Time
      '', '', '', '', '', ''        // H-M: 予約カラム
    ];
    
    // Log_Rawシートにデータを追加
    logRawSheet.appendRow(rowData);
    
    // 成功レスポンス
    return {
      success: true,
      message: '記録しました',
      data: {
        timestamp: timestamp,
        action: action,
        user: userInfo.name
      }
    };
    
  } catch (error) {
    throw new Error(`データ保存エラー: ${error.message}`);
  }
} 