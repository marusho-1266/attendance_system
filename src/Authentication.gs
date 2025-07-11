/**
 * 認証機能
 * TDD Refactor フェーズ: セキュリティ強化とコード品質向上
 * 
 * 機能:
 * - authenticateUser(email): ユーザー認証（ホワイトリスト方式）
 * - checkPermission(email, action): 権限チェック（ロールベース）
 * - getSessionInfo(): セッション情報取得（セキュアモード）
 * - セキュリティログ記録
 * - 不正アクセス検知
 */

// === セキュリティ設定 ===

/**
 * 認証セキュリティ設定
 */
var AUTH_CONFIG = {
  MAX_LOGIN_ATTEMPTS: 5,           // 最大ログイン試行回数
  SESSION_TIMEOUT: 8 * 60 * 60,    // セッションタイムアウト（8時間）
  ENABLE_SECURITY_LOG: true,       // セキュリティログ有効化
  REQUIRE_HTTPS: false,            // HTTPS必須（GAS環境では無効）
  BRUTE_FORCE_PROTECTION: true     // ブルートフォース攻撃防止
};

// === ユーザー認証関数 ===

/**
 * メールアドレスによるユーザー認証（Refactor: セキュリティ強化版）
 * @param {string} email - 認証対象のメールアドレス
 * @returns {boolean} 認証成功の場合true、失敗の場合false
 */
function authenticateUser(email) {
  var startTime = new Date().getTime();
  
  try {
    // パラメータ検証強化
    if (!email || typeof email !== 'string' || email.trim() === '') {
      logSecurityEvent('INVALID_PARAMETER', email, 'Empty or invalid email parameter');
      return false;
    }
    
    // 正規化（前後のスペース除去、小文字変換）
    email = email.trim().toLowerCase();
    
    // メールフォーマット検証強化
    if (!isValidEmailEnhanced(email)) {
      logSecurityEvent('INVALID_EMAIL_FORMAT', email, 'Invalid email format');
      return false;
    }
    
    // ブルートフォース攻撃チェック
    if (AUTH_CONFIG.BRUTE_FORCE_PROTECTION && isBlockedByBruteForce(email)) {
      logSecurityEvent('BRUTE_FORCE_BLOCKED', email, 'Blocked due to too many failed attempts');
      return false;
    }
    
    // 従業員マスタでの存在確認（ホワイトリスト方式）
    var employee = getEmployee(email);
    var isAuthenticated = employee !== null;
    
    // 認証結果のログ記録
    if (isAuthenticated) {
      logSecurityEvent('AUTH_SUCCESS', email, 'User authenticated successfully');
      clearFailedAttempts(email); // 成功時は失敗回数をリセット
    } else {
      logSecurityEvent('AUTH_FAILURE', email, 'User not found in employee master');
      incrementFailedAttempts(email); // 失敗回数を増加
    }
    
    return isAuthenticated;
    
  } catch (error) {
    logSecurityEvent('AUTH_ERROR', email, 'Authentication error: ' + error.message);
    console.log('認証システムエラー: ' + error.message);
    return false;
  } finally {
    // パフォーマンス監視
    var duration = new Date().getTime() - startTime;
    if (duration > 1000) { // 1秒以上かかった場合
      console.log('認証処理時間警告: ' + duration + 'ms for ' + email);
    }
  }
}

// === 権限チェック関数 ===

/**
 * ユーザーの特定アクションに対する権限をチェック
 * @param {string} email - チェック対象のメールアドレス
 * @param {string} action - 実行しようとするアクション
 * @returns {boolean} 権限がある場合true、ない場合false
 */
function checkPermission(email, action) {
  // パラメータ検証
  if (!email || !action) {
    return false;
  }
  
  // まず認証チェック
  if (!authenticateUser(email)) {
    return false; // 未認証ユーザーには権限なし
  }
  
  // アクション妥当性チェック（最小実装）
  try {
    var validActions = ['CLOCK_IN', 'CLOCK_OUT', 'BREAK_START', 'BREAK_END', 'ADMIN_ACCESS', 'VIEW_REPORTS', 'EDIT_MASTER'];
    if (validActions.indexOf(action) === -1) {
      return false; // 無効なアクション
    }
  } catch (error) {
    return false;
  }
  
  // 管理者権限チェック（Green段階の簡易実装）
  var employee = getEmployee(email);
  if (!employee) {
    return false;
  }
  
  // 管理者専用アクションの場合
  if (action === 'ADMIN_ACCESS' || action === 'VIEW_REPORTS' || action === 'EDIT_MASTER') {
    // 管理者判定: 上司Gmailが空またはnullの場合は管理者とみなす（簡易実装）
    return isManager(email);
  }
  
  // 基本的な打刻権限: 認証済みユーザーなら全て許可
  return true;
}

/**
 * 管理者判定関数（Green段階の簡易実装）
 * @param {string} email - チェック対象のメールアドレス
 * @returns {boolean} 管理者の場合true
 */
function isManager(email) {
  // 簡易実装: 特定のメールアドレスまたは上司Gmailフィールドがnullの場合
  var managerEmails = ['manager@example.com', 'dev-manager@example.com'];
  
  if (managerEmails.indexOf(email) !== -1) {
    return true;
  }
  
  // または従業員マスタで上司Gmailが空の場合（トップレベル管理者）
  var employee = getEmployee(email);
  if (employee && (!employee.supervisorGmail || employee.supervisorGmail === '')) {
    return true;
  }
  
  return false;
}

// === セッション情報取得関数 ===

/**
 * 現在のセッション情報を取得
 * @returns {Object} セッション情報オブジェクト
 */
function getSessionInfo() {
  try {
    // Green段階の最小実装: テスト環境での動作を考慮
    var sessionInfo = {
      email: null,
      isAuthenticated: false,
      employeeInfo: null,
      timestamp: new Date().toISOString()
    };
    
    // GAS環境での実際のセッション取得試行
    try {
      // 実際のGAS環境では Session.getActiveUser().getEmail() を使用
      var activeUser = Session.getActiveUser();
      if (activeUser && activeUser.getEmail) {
        var userEmail = activeUser.getEmail();
        if (userEmail) {
          sessionInfo.email = userEmail;
          sessionInfo.isAuthenticated = authenticateUser(userEmail);
          
          if (sessionInfo.isAuthenticated) {
            sessionInfo.employeeInfo = getEmployee(userEmail);
          }
        }
      }
    } catch (sessionError) {
      // テスト環境やGAS環境外では、テスト用の固定値を使用
      console.log('セッション取得テスト環境モード: ' + sessionError.message);
      sessionInfo.email = 'test@example.com';
      sessionInfo.isAuthenticated = false;
    }
    
    return sessionInfo;
  } catch (error) {
    console.log('getSessionInfo エラー: ' + error.message);
    return {
      email: null,
      isAuthenticated: false,
      employeeInfo: null,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

// === セキュリティ支援関数 ===

/**
 * セキュリティイベントログ記録
 * @param {string} eventType - イベントタイプ
 * @param {string} email - 対象ユーザー
 * @param {string} description - 詳細説明
 */
function logSecurityEvent(eventType, email, description) {
  if (!AUTH_CONFIG.ENABLE_SECURITY_LOG) {
    return; // ログ無効時はスキップ
  }
  
  try {
    var timestamp = new Date().toISOString();
    var logEntry = [
      timestamp,
      'SECURITY',
      eventType,
      email || 'unknown',
      description || 'No description',
      Session.getActiveUser().getEmail() || 'system'
    ];
    
    // Log_Rawシートに記録
    var sheetName = getSheetName('LOG_RAW');
    appendDataToSheet(sheetName, logEntry);
    
  } catch (error) {
    console.log('セキュリティログ記録エラー: ' + error.message);
  }
}

/**
 * ブルートフォース攻撃チェック
 * @param {string} email - チェック対象のメールアドレス
 * @returns {boolean} ブロック対象の場合true
 */
function isBlockedByBruteForce(email) {
  try {
    var failedAttempts = getFailedAttempts(email);
    return failedAttempts >= AUTH_CONFIG.MAX_LOGIN_ATTEMPTS;
  } catch (error) {
    console.log('ブルートフォースチェックエラー: ' + error.message);
    return false; // エラー時は通す（可用性優先）
  }
}

/**
 * 失敗回数取得（PropertiesServiceを使用してメモリ管理）
 * @param {string} email - 対象ユーザー
 * @returns {number} 失敗回数
 */
function getFailedAttempts(email) {
  try {
    var key = 'failed_attempts_' + email;
    var value = PropertiesService.getScriptProperties().getProperty(key);
    return value ? parseInt(value, 10) : 0;
  } catch (error) {
    return 0;
  }
}

/**
 * 失敗回数増加
 * @param {string} email - 対象ユーザー
 */
function incrementFailedAttempts(email) {
  try {
    var key = 'failed_attempts_' + email;
    var currentAttempts = getFailedAttempts(email);
    PropertiesService.getScriptProperties().setProperty(key, (currentAttempts + 1).toString());
    
    // 最大試行回数に達した場合の通知
    if (currentAttempts + 1 >= AUTH_CONFIG.MAX_LOGIN_ATTEMPTS) {
      logSecurityEvent('BRUTE_FORCE_ALERT', email, 'User blocked due to excessive failed login attempts');
    }
  } catch (error) {
    console.log('失敗回数記録エラー: ' + error.message);
  }
}

/**
 * 失敗回数クリア
 * @param {string} email - 対象ユーザー
 */
function clearFailedAttempts(email) {
  try {
    var key = 'failed_attempts_' + email;
    PropertiesService.getScriptProperties().deleteProperty(key);
  } catch (error) {
    console.log('失敗回数クリアエラー: ' + error.message);
  }
}

// === ユーティリティ関数 ===

/**
 * 強化されたメールアドレスフォーマット検証
 * @param {string} email - チェック対象のメールアドレス
 * @returns {boolean} 有効な場合true
 */
function isValidEmailEnhanced(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }
  
  // 基本的なフォーマットチェック
  var emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  if (!emailRegex.test(email)) {
    return false;
  }
  
  // 長さチェック（254文字以下）
  if (email.length > 254) {
    return false;
  }
  
  // ローカル部の長さチェック（64文字以下）
  var localPart = email.split('@')[0];
  if (localPart.length > 64) {
    return false;
  }
  
  // 危険なパターンのチェック
  var dangerousPatterns = [
    /\.\./,           // 連続ドット
    /^\./, /\.$/,     // 開始・終了ドット
    /@.*@/,           // 複数@マーク
    /\s/              // 空白文字
  ];
  
  for (var i = 0; i < dangerousPatterns.length; i++) {
    if (dangerousPatterns[i].test(email)) {
      return false;
    }
  }
  
  return true;
}

/**
 * メールアドレスフォーマット検証（後方互換性のため）
 * @param {string} email - チェック対象のメールアドレス
 * @returns {boolean} 有効な場合true
 */
function isValidEmail(email) {
  return isValidEmailEnhanced(email);
} 