/**
 * 認証・セキュリティモジュール
 * 従業員の認証、アクセス権限の検証、セキュリティ機能を提供
 */

/**
 * メールアドレスの形式を検証
 * @param {string} email - 検証するメールアドレス
 * @returns {boolean} 有効なメールアドレス形式かどうか
 */
function validateEmailFormat(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }
  
  // 基本的なメールアドレス形式の正規表現
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  
  // 基本的な形式チェック
  if (!emailRegex.test(email)) {
    return false;
  }
  
  // 追加のセキュリティチェック
  const emailLower = email.toLowerCase();
  
  // 危険な文字やパターンのチェック
  const dangerousPatterns = [
    /<script/i,
    /javascript:/i,
    /data:/i,
    /vbscript:/i,
    /on\w+\s*=/i,
    /<iframe/i,
    /<object/i,
    /<embed/i
  ];
  
  for (const pattern of dangerousPatterns) {
    if (pattern.test(emailLower)) {
      return false;
    }
  }
  
  // 長さチェック（RFC 5321準拠）
  if (email.length > 254) {
    return false;
  }
  
  // ローカル部分とドメイン部分の長さチェック
  const parts = email.split('@');
  if (parts.length !== 2) {
    return false;
  }
  
  const localPart = parts[0];
  const domainPart = parts[1];
  
  if (localPart.length > 64 || domainPart.length > 253) {
    return false;
  }
  
  return true;
}

/**
 * イベントの信頼性を検証
 * @param {Object} e - WebAppイベントオブジェクト
 * @returns {boolean} 信頼できるイベントかどうか
 */
function validateEventTrust(e) {
  if (!e) {
    return false;
  }
  
  // セッションからの認証確認
  try {
    const activeUser = Session.getActiveUser();
    if (activeUser && activeUser.getEmail()) {
      // セッションが有効な場合は信頼できる
      return true;
    }
  } catch (error) {
    Logger.log(`セッション認証チェックエラー: ${error.message}`);
  }
  
  // イベントの発生元チェック（必要に応じて拡張）
  // 例: ContentService.getMimeType()、特定のヘッダー値など
  
  // 現在は基本的なチェックのみ
  // 将来的にはCSRFトークン、署名検証などを追加可能
  
  return false;
}

/**
 * 現在のユーザーを認証し、従業員情報を取得
 * @returns {Object} 認証された従業員情報
 * @throws {Error} 認証失敗時
 */
function authenticateUser() {
  return withErrorHandling(() => {
    // Google Apps Scriptのセッションから現在のユーザーのメールアドレスを取得
    const userEmail = Session.getActiveUser().getEmail();
    
    if (!userEmail) {
      Logger.log('Session.getActiveUser()がnullを返しました。代替認証方法を試行します。');
      throw new Error('ユーザーのメールアドレスを取得できませんでした。WebAppのデプロイメント設定を確認してください。');
    }
    
    Logger.log(`認証試行: ${userEmail}`);
    
    // 従業員マスターとの照合
    const employeeInfo = getEmployeeInfo(userEmail);
    
    if (!employeeInfo) {
      Logger.log(`認証失敗: 従業員マスタに存在しません - ${userEmail}`);
      throw new Error('認証失敗: 従業員マスタに登録されていません');
    }
    
    Logger.log(`認証成功: ${employeeInfo.name} (${employeeInfo.id})`);
    
    // アクセスログを記録
    logAccess('LOGIN', employeeInfo.id, `認証成功: ${userEmail}`);
    
    return employeeInfo;
    
  }, 'Authentication.authenticateUser', 'HIGH', {
    userEmail: Session.getActiveUser().getEmail()
  });
}

/**
 * WebAppイベントからユーザーを認証し、従業員情報を取得
 * @param {Object} e - WebAppイベントオブジェクト
 * @returns {Object} 認証された従業員情報
 * @throws {Error} 認証失敗時
 */
function authenticateUserFromEvent(e) {
  return withErrorHandling(() => {
    let userEmail = null;
    
    // 1. イベントパラメータからメールアドレスを取得（セキュリティ検証付き）
    if (e && e.parameter && e.parameter.email) {
      const eventEmail = e.parameter.email;
      
      // メールアドレスの形式検証
      if (!validateEmailFormat(eventEmail)) {
        Logger.log(`セキュリティ警告: 無効なメールアドレス形式 - ${eventEmail}`);
        throw new Error('セキュリティエラー: 無効なメールアドレス形式です');
      }
      
      // イベントの信頼性検証
      if (!validateEventTrust(e)) {
        Logger.log(`セキュリティ警告: 信頼できないイベントソース - ${eventEmail}`);
        throw new Error('セキュリティエラー: 信頼できないイベントソースです');
      }
      
      userEmail = eventEmail;
      Logger.log(`イベントパラメータからメールアドレスを取得（検証済み）: ${userEmail}`);
    }
    
    // 2. イベントパラメータにない場合は、セッションから取得を試行
    if (!userEmail) {
      try {
        const sessionEmail = Session.getActiveUser().getEmail();
        if (sessionEmail) {
          // セッションから取得したメールアドレスも形式検証
          if (!validateEmailFormat(sessionEmail)) {
            Logger.log(`セキュリティ警告: セッションから無効なメールアドレス形式 - ${sessionEmail}`);
            throw new Error('セキュリティエラー: セッションから無効なメールアドレス形式です');
          }
          
          userEmail = sessionEmail;
          Logger.log(`セッションからメールアドレスを取得（検証済み）: ${userEmail}`);
        }
      } catch (error) {
        Logger.log(`セッションからのメールアドレス取得に失敗: ${error.message}`);
      }
    }
    
    // 3. それでも取得できない場合は、デプロイメント設定の問題として処理
    if (!userEmail) {
      Logger.log('すべての認証方法でメールアドレスの取得に失敗しました');
      throw new Error('ユーザーのメールアドレスを取得できませんでした。デプロイメント設定を確認してください。');
    }
    
    Logger.log(`認証試行: ${userEmail}`);
    
    // 従業員マスターとの照合
    const employeeInfo = getEmployeeInfo(userEmail);
    
    if (!employeeInfo) {
      Logger.log(`認証失敗: 従業員マスタに存在しません - ${userEmail}`);
      throw new Error('認証失敗: 従業員マスタに登録されていません');
    }
    
    Logger.log(`認証成功: ${employeeInfo.name} (${employeeInfo.id})`);
    
    // アクセスログを記録
    logAccess('LOGIN', employeeInfo.id, `認証成功: ${userEmail}`);
    
    return employeeInfo;
    
  }, 'Authentication.authenticateUserFromEvent', 'HIGH', {
    eventParameter: e ? e.parameter : null,
    sessionUser: (() => {
      const activeUser = Session.getActiveUser();
      return activeUser ? activeUser.getEmail() : null;
    })()
  });
}

/**
 * 指定されたフィールドで従業員を検索する共通ヘルパー関数
 * @param {number} fieldIndex - 検索するフィールドのインデックス（0ベース）
 * @param {string} searchValue - 検索値
 * @param {boolean} caseSensitive - 大文字小文字を区別するかどうか（デフォルト: false）
 * @returns {Object|null} 従業員情報またはnull
 */
function findEmployeeByField(fieldIndex, searchValue, caseSensitive = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Employee');
  
  if (!sheet) {
    throw new Error('Master_Employeeシートが見つかりません');
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return null;
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  
  // 指定されたフィールドで検索
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const fieldValue = row[fieldIndex];
    
    if (fieldValue) {
      const fieldStr = fieldValue.toString();
      const searchStr = searchValue.toString();
      
      const isMatch = caseSensitive 
        ? fieldStr === searchStr
        : fieldStr.toLowerCase() === searchStr.toLowerCase();
      
      if (isMatch) {
        return {
          id: row[0],           // 列A: 社員ID
          name: row[1],         // 列B: 氏名
          email: row[2],        // 列C: Gmail
          department: row[3],   // 列D: 所属
          employmentType: row[4], // 列E: 雇用区分
          supervisorEmail: row[5], // 列F: 上司Gmail
          startTime: row[6],    // 列G: 基準始業
          endTime: row[7]       // 列H: 基準終業
        };
      }
    }
  }
  
  return null;
}

/**
 * メールアドレスから従業員情報を取得
 * @param {string} email - 検索するメールアドレス
 * @returns {Object|null} 従業員情報またはnull
 */
function getEmployeeInfo(email) {
  return withErrorHandling(() => {
    // メールアドレスで検索（列C: Gmail、インデックス2）
    return findEmployeeByField(2, email, false);
    
  }, 'Authentication.getEmployeeInfo', 'MEDIUM', {
    email: email
  });
}

/**
 * 従業員のアクセス権限を検証
 * @param {string} email - 検証するメールアドレス
 * @returns {boolean} アクセス許可の可否
 */
function validateEmployeeAccess(email) {
  try {
    if (!email) {
      Logger.log('アクセス権限検証: メールアドレスが指定されていません');
      return false;
    }
    
    const employeeInfo = getEmployeeInfo(email);
    
    if (!employeeInfo) {
      Logger.log(`アクセス権限検証失敗: 従業員マスタに存在しません - ${email}`);
      return false;
    }
    
    // 追加の権限チェック（必要に応じて拡張）
    // 例: 退職者フラグ、アクセス停止フラグなど
    
    Logger.log(`アクセス権限検証成功: ${employeeInfo.name} (${employeeInfo.id})`);
    return true;
    
  } catch (error) {
    Logger.log(`アクセス権限検証エラー: ${error.message}`);
    return false;
  }
}

/**
 *
 承認者権限を検証
 * @param {string} email - 検証するメールアドレス
 * @param {string} targetEmployeeId - 対象従業員のID
 * @returns {boolean} 承認権限の可否
 */
function validateApprovalAccess(email, targetEmployeeId) {
  try {
    if (!email || !targetEmployeeId) {
      Logger.log('承認権限検証: 必要なパラメータが不足しています');
      return false;
    }
    
    // 対象従業員の情報を取得
    const targetEmployee = getEmployeeById(targetEmployeeId);
    if (!targetEmployee) {
      Logger.log(`承認権限検証失敗: 対象従業員が見つかりません - ${targetEmployeeId}`);
      return false;
    }
    
    // 承認者（上司）のメールアドレスと一致するかチェック
    if (targetEmployee.supervisorEmail && 
        targetEmployee.supervisorEmail.toLowerCase() === email.toLowerCase()) {
      Logger.log(`承認権限検証成功: ${email} は ${targetEmployeeId} の承認者です`);
      return true;
    }
    
    // 管理者権限のチェック（必要に応じて実装）
    if (isSystemAdmin(email)) {
      Logger.log(`承認権限検証成功: ${email} はシステム管理者です`);
      return true;
    }
    
    Logger.log(`承認権限検証失敗: ${email} は ${targetEmployeeId} の承認権限がありません`);
    return false;
    
  } catch (error) {
    Logger.log(`承認権限検証エラー: ${error.message}`);
    return false;
  }
}

/**
 * 社員IDから従業員情報を取得
 * @param {string} employeeId - 社員ID
 * @returns {Object|null} 従業員情報またはnull
 */
function getEmployeeById(employeeId) {
  try {
    // 社員IDで検索（列A: 社員ID、インデックス0）
    return findEmployeeByField(0, employeeId, true);
    
  } catch (error) {
    Logger.log(`従業員情報取得エラー (ID検索): ${error.message}`);
    throw new Error(`従業員情報の取得に失敗しました: ${error.message}`);
  }
}

/**
 * システム管理者権限をチェック
 * @param {string} email - チェックするメールアドレス
 * @returns {boolean} 管理者権限の可否
 */
function isSystemAdmin(email) {
  try {
    // 設定から管理者メールアドレスを取得
    const adminEmails = getAdminEmails();
    
    if (!adminEmails || adminEmails.length === 0) {
      return false;
    }
    
    return adminEmails.some(adminEmail => 
      adminEmail.toLowerCase() === email.toLowerCase()
    );
    
  } catch (error) {
    Logger.log(`管理者権限チェックエラー: ${error.message}`);
    return false;
  }
}

/**
 * アクセスログを記録
 * @param {string} action - アクション種別
 * @param {string} employeeId - 従業員ID
 * @param {string} details - 詳細情報
 */
function logAccess(action, employeeId, details) {
  try {
    // アクセスログシートの取得または作成
    let logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Access_Log');
    
    if (!logSheet) {
      // アクセスログシートが存在しない場合は作成
      logSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Access_Log');
      
      // ヘッダー行を設定
      logSheet.getRange(1, 1, 1, 5).setValues([[
        'タイムスタンプ', '社員ID', 'アクション', '詳細', 'ユーザーEmail'
      ]]);
      
      // ヘッダー行のフォーマット
      const headerRange = logSheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f0f0f0');
    }
    
    // ログエントリを追加
    const logEntry = [
      new Date(),
      employeeId || '',
      action,
      details || '',
      Session.getActiveUser().getEmail() || ''
    ];
    
    logSheet.appendRow(logEntry);
    
  } catch (error) {
    Logger.log(`アクセスログ記録エラー: ${error.message}`);
    // ログ記録の失敗は処理を停止させない
  }
}

/**
 * 管理者メールアドレスを取得
 * @returns {string|null} 管理者メールアドレス
 */
function getAdminEmail() {
  try {
    // Config.gsから管理者メールアドレスを取得
    if (typeof getConfig === 'function') {
      const config = getConfig();
      return config.adminEmail || null;
    }
    
    // フォールバック: スプレッドシートの所有者
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const owner = spreadsheet.getOwner();
    return owner ? owner.getEmail() : null;
    
  } catch (error) {
    Logger.log(`管理者メール取得エラー: ${error.message}`);
    return null;
  }
}

/**
 * 管理者メールアドレス一覧を取得
 * @returns {Array<string>} 管理者メールアドレス配列
 */
function getAdminEmails() {
  try {
    // Config.gsから管理者メールアドレス一覧を取得
    if (typeof getConfig === 'function') {
      const config = getConfig();
      if (config.adminEmails && Array.isArray(config.adminEmails)) {
        return config.adminEmails;
      }
      if (config.adminEmail) {
        return [config.adminEmail];
      }
    }
    
    // フォールバック: スプレッドシートの所有者
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const owner = spreadsheet.getOwner();
    return owner ? [owner.getEmail()] : [];
    
  } catch (error) {
    Logger.log(`管理者メール一覧取得エラー: ${error.message}`);
    return [];
  }
}

/**
 * 認証が必要な処理を実行するラッパー関数
 * @param {Function} callback - 実行する関数
 * @returns {*} コールバック関数の戻り値
 * @throws {Error} 認証失敗時
 */
function requireAuthentication(callback) {
  try {
    const user = authenticateUser();
    return callback(user);
  } catch (error) {
    Logger.log(`認証必須処理でエラー: ${error.message}`);
    throw error;
  }
}

/**
 * WebAppのデプロイメント設定を確認する
 * @returns {Object} デプロイメント設定情報
 */
function checkDeploymentSettings() {
  try {
    const deploymentInfo = {
      sessionUser: null,
      sessionUserEmail: null,
      hasActiveUser: false,
      deploymentMode: 'Unknown',
      recommendations: []
    };
    
    // セッション情報の確認
    try {
      const activeUser = Session.getActiveUser();
      deploymentInfo.hasActiveUser = activeUser !== null;
      
      if (activeUser) {
        deploymentInfo.sessionUser = activeUser;
        deploymentInfo.sessionUserEmail = activeUser.getEmail();
      }
    } catch (error) {
      deploymentInfo.recommendations.push('Session.getActiveUser()でエラーが発生: ' + error.message);
    }
    
    // デプロイメントモードの推定
    if (deploymentInfo.hasActiveUser && deploymentInfo.sessionUserEmail) {
      deploymentInfo.deploymentMode = 'Execute as: Me, Who has access: Anyone';
    } else if (deploymentInfo.hasActiveUser && !deploymentInfo.sessionUserEmail) {
      deploymentInfo.deploymentMode = 'Execute as: Me, Who has access: Anyone with Google Account';
    } else {
      deploymentInfo.deploymentMode = 'Execute as: User accessing the web app';
      deploymentInfo.recommendations.push('デプロイメント設定を「Execute as: Me」に変更することを推奨します');
    }
    
    // 推奨事項の追加
    if (!deploymentInfo.sessionUserEmail) {
      deploymentInfo.recommendations.push('Session.getActiveUser().getEmail()がnullを返しています');
      deploymentInfo.recommendations.push('WebAppのデプロイメント設定で「Execute as: Me」を選択してください');
      deploymentInfo.recommendations.push('または、イベントパラメータからemailを取得する方法を使用してください');
    }
    
    Logger.log('デプロイメント設定確認結果: ' + JSON.stringify(deploymentInfo, null, 2));
    
    return deploymentInfo;
    
  } catch (error) {
    Logger.log(`デプロイメント設定確認エラー: ${error.message}`);
    return {
      error: error.message,
      recommendations: ['デプロイメント設定の確認に失敗しました']
    };
  }
}

/**
 * 承認権限が必要な処理を実行するラッパー関数
 * @param {string} targetEmployeeId - 対象従業員ID
 * @param {Function} callback - 実行する関数
 * @returns {*} コールバック関数の戻り値
 * @throws {Error} 権限不足時
 */
function requireApprovalAccess(targetEmployeeId, callback) {
  try {
    const user = authenticateUser();
    
    if (!validateApprovalAccess(user.email, targetEmployeeId)) {
      throw new Error('承認権限がありません');
    }
    
    return callback(user);
  } catch (error) {
    Logger.log(`承認権限必須処理でエラー: ${error.message}`);
    throw error;
  }
}

/**
 * 新しい認証機能のテスト
 */
function testNewAuthenticationMethods() {
  try {
    Logger.log('=== 新しい認証機能テスト開始 ===');
    
    // 1. デプロイメント設定確認テスト
    Logger.log('1. デプロイメント設定確認テスト');
    const deploymentInfo = checkDeploymentSettings();
    Logger.log('デプロイメント情報: ' + JSON.stringify(deploymentInfo, null, 2));
    
    // 2. 従来の認証方法テスト
    Logger.log('2. 従来の認証方法テスト');
    try {
      const user1 = authenticateUser();
      Logger.log('✓ 従来の認証成功: ' + user1.name);
    } catch (error) {
      Logger.log('✗ 従来の認証失敗: ' + error.message);
    }
    
    // 3. イベントベース認証テスト（モックイベント）
    Logger.log('3. イベントベース認証テスト');
    try {
      const mockEvent = {
        parameter: {
          email: Session.getActiveUser().getEmail()
        }
      };
      const user2 = authenticateUserFromEvent(mockEvent);
      Logger.log('✓ イベントベース認証成功: ' + user2.name);
    } catch (error) {
      Logger.log('✗ イベントベース認証失敗: ' + error.message);
    }
    
    // 4. イベントベース認証テスト（emailパラメータなし）
    Logger.log('4. イベントベース認証テスト（emailパラメータなし）');
    try {
      const mockEvent2 = {
        parameter: {}
      };
      const user3 = authenticateUserFromEvent(mockEvent2);
      Logger.log('✓ フォールバック認証成功: ' + user3.name);
    } catch (error) {
      Logger.log('✗ フォールバック認証失敗: ' + error.message);
    }
    
    Logger.log('=== 新しい認証機能テスト完了 ===');
    
  } catch (error) {
    Logger.log('新しい認証機能テストエラー: ' + error.toString());
  }
}