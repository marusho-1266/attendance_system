/**
 * システム設定管理
 * System_Configシートから設定値を動的に読み取り管理
 */

// === 設定値キャッシュ（パフォーマンス向上） ===

/**
 * 設定値キャッシュ（1回読み込んだら保持）
 */
var CONFIG_CACHE = {};
var CONFIG_CACHE_TIMESTAMP = null;
var CONFIG_CACHE_EXPIRY = 5 * 60 * 1000; // 5分間キャッシュ

// === 設定値読み取り関数 ===

/**
 * System_Configシートから設定値を取得
 * @param {string} configKey - 設定キー
 * @returns {string|null} 設定値（見つからない場合はnull）
 */
function getSystemConfig(configKey) {
  if (!configKey) {
    throw new Error('Configuration key is required');
  }
  
  try {
    // キャッシュチェック
    if (isConfigCacheValid() && CONFIG_CACHE.hasOwnProperty(configKey)) {
      return CONFIG_CACHE[configKey];
    }
    
    // キャッシュが無効または存在しない場合は再読み込み
    loadConfigCache();
    
    return CONFIG_CACHE[configKey] || null;
    
  } catch (error) {
    console.log('設定値取得エラー: ' + error.message);
    return null;
  }
}

/**
 * 管理者メールアドレスリストを取得
 * @returns {Array} 管理者メールアドレスの配列
 */
function getManagerEmails() {
  try {
    var managerEmailsString = getSystemConfig('MANAGER_EMAILS');
    
    if (!managerEmailsString) {
      console.log('管理者メール設定が見つかりません。従業員マスタから判定します。');
      return []; // 空配列を返し、従業員マスタの上司Gmail判定にフォールバック
    }
    
    // カンマ区切りの文字列を配列に変換
    var emails = managerEmailsString.split(',').map(function(email) {
      return email.trim().toLowerCase();
    }).filter(function(email) {
      return email.length > 0 && isValidEmailEnhanced(email);
    });
    
    console.log('管理者メール取得: ' + emails.length + '件');
    return emails;
    
  } catch (error) {
    console.log('管理者メール取得エラー: ' + error.message);
    return []; // エラー時は空配列（フォールバック動作）
  }
}

/**
 * システム設定値を設定/更新
 * @param {string} configKey - 設定キー
 * @param {string} configValue - 設定値
 * @param {string} description - 説明（オプション）
 * @param {boolean} isActive - 有効フラグ（デフォルト: true）
 */
function setSystemConfig(configKey, configValue, description, isActive) {
  if (!configKey || !configValue) {
    throw new Error('Configuration key and value are required');
  }
  
  try {
    var sheetName = getSheetName('SYSTEM_CONFIG');
    var sheet = getSheet(sheetName);
    var data = sheet.getDataRange().getValues();
    
    description = description || '';
    isActive = (isActive !== false); // デフォルトtrue
    
    // 既存設定を検索
    var existingRowIndex = -1;
    for (var i = 1; i < data.length; i++) { // ヘッダー行をスキップ
      if (data[i][getColumnIndex('SYSTEM_CONFIG', 'CONFIG_KEY')] === configKey) {
        existingRowIndex = i;
        break;
      }
    }
    
    var rowData = [configKey, configValue, description, isActive];
    
    if (existingRowIndex !== -1) {
      // 既存設定を更新
      var range = sheet.getRange(existingRowIndex + 1, 1, 1, 4);
      range.setValues([rowData]);
      console.log('設定値更新: ' + configKey + ' = ' + configValue);
    } else {
      // 新規設定を追加
      appendDataToSheet(sheetName, rowData);
      console.log('設定値追加: ' + configKey + ' = ' + configValue);
    }
    
    // キャッシュを無効化
    invalidateConfigCache();
    
  } catch (error) {
    throw new Error('設定値設定エラー: ' + error.message);
  }
}

// === キャッシュ管理関数 ===

/**
 * 設定値キャッシュが有効かチェック
 * @returns {boolean} 有効な場合true
 */
function isConfigCacheValid() {
  if (!CONFIG_CACHE_TIMESTAMP) {
    return false;
  }
  
  var now = new Date().getTime();
  return (now - CONFIG_CACHE_TIMESTAMP) < CONFIG_CACHE_EXPIRY;
}

/**
 * 設定値キャッシュを読み込み
 */
function loadConfigCache() {
  try {
    CONFIG_CACHE = {};
    
    var sheetName = getSheetName('SYSTEM_CONFIG');
    if (!sheetExists(sheetName)) {
      console.log('System_Configシートが存在しません。初期化してください。');
      return;
    }
    
    var sheet = getSheet(sheetName);
    var data = sheet.getDataRange().getValues();
    
    // ヘッダー行をスキップして設定値を読み込み
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var key = row[getColumnIndex('SYSTEM_CONFIG', 'CONFIG_KEY')];
      var value = row[getColumnIndex('SYSTEM_CONFIG', 'CONFIG_VALUE')];
      var isActive = row[getColumnIndex('SYSTEM_CONFIG', 'IS_ACTIVE')];
      
      // 有効な設定のみキャッシュ
      if (key && value && isActive) {
        CONFIG_CACHE[key] = value;
      }
    }
    
    CONFIG_CACHE_TIMESTAMP = new Date().getTime();
    console.log('設定値キャッシュ読み込み完了: ' + Object.keys(CONFIG_CACHE).length + '件');
    
  } catch (error) {
    console.log('設定値キャッシュ読み込みエラー: ' + error.message);
    CONFIG_CACHE = {};
  }
}

/**
 * 設定値キャッシュを無効化
 */
function invalidateConfigCache() {
  CONFIG_CACHE = {};
  CONFIG_CACHE_TIMESTAMP = null;
}

// === 初期設定セットアップ ===

/**
 * システム設定の初期データをセットアップ
 */
function setupSystemConfigData() {
  try {
    console.log('システム設定初期データのセットアップ開始...');
    
    // 管理者メールアドレス設定
    setSystemConfig(
      'MANAGER_EMAILS',
      'manager@example.com,dev-manager@example.com',
      '管理者権限を持つメールアドレス（カンマ区切り）',
      true
    );
    
    // セッションタイムアウト設定
    setSystemConfig(
      'SESSION_TIMEOUT',
      '28800', // 8時間（秒）
      'セッションタイムアウト時間（秒）',
      true
    );
    
    // ブルートフォース攻撃防止設定
    setSystemConfig(
      'MAX_LOGIN_ATTEMPTS',
      '5',
      '最大ログイン試行回数',
      true
    );
    
    // セキュリティログ有効化設定
    setSystemConfig(
      'ENABLE_SECURITY_LOG',
      'true',
      'セキュリティログ機能の有効化',
      true
    );
    
    console.log('システム設定初期データのセットアップ完了');
    
  } catch (error) {
    console.log('システム設定初期データセットアップエラー: ' + error.message);
    throw error;
  }
} 