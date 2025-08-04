/**
 * 認証・セキュリティモジュール
 * 従業員の認証、アクセス権限の検証、セキュリティ機能を提供
 */

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
      throw new Error('ユーザーのメールアドレスを取得できませんでした');
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
 * メールアドレスから従業員情報を取得
 * @param {string} email - 検索するメールアドレス
 * @returns {Object|null} 従業員情報またはnull
 */
function getEmployeeInfo(email) {
  return withErrorHandling(() => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Employee');
    
    if (!sheet) {
      throw new Error('Master_Employeeシートが見つかりません');
    }
    
    // データ範囲を取得（ヘッダー行を除く）
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('従業員マスタにデータが存在しません');
      return null;
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    // メールアドレスで検索（列C: Gmail）
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[2] && row[2].toString().toLowerCase() === email.toLowerCase()) {
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
    
    return null;
    
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
}/**
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
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Employee');
    
    if (!sheet) {
      throw new Error('Master_Employeeシートが見つかりません');
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return null;
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    // 社員IDで検索（列A: 社員ID）
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[0].toString() === employeeId.toString()) {
        return {
          id: row[0],
          name: row[1],
          email: row[2],
          department: row[3],
          employmentType: row[4],
          supervisorEmail: row[5],
          startTime: row[6],
          endTime: row[7]
        };
      }
    }
    
    return null;
    
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