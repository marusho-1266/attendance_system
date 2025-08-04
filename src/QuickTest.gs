/**
 * 簡単な動作確認用テスト関数
 */

/**
 * 現在の実装状況を確認する簡単なテスト
 */
function quickSystemCheck() {
  try {
    Logger.log('=== 簡単なシステム確認開始 ===');
    
    const results = {
      timestamp: new Date(),
      checks: {},
      success: true,
      errors: []
    };
    
    // 1. 設定ファイル確認
    try {
      const systemConfig = getConfig('SYSTEM');
      results.checks.config = {
        success: true,
        message: `システム設定OK: ${systemConfig.SYSTEM_NAME} v${systemConfig.VERSION}`
      };
    } catch (error) {
      results.checks.config = {
        success: false,
        message: '設定ファイルエラー',
        error: error.message
      };
      results.success = false;
      results.errors.push(error.message);
    }
    
    // 2. スプレッドシート確認
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheetCount = spreadsheet.getSheets().length;
      results.checks.spreadsheet = {
        success: true,
        message: `スプレッドシートアクセスOK: ${sheetCount}シート存在`
      };
    } catch (error) {
      results.checks.spreadsheet = {
        success: false,
        message: 'スプレッドシートアクセスエラー',
        error: error.message
      };
      results.success = false;
      results.errors.push(error.message);
    }
    
    // 3. 認証機能確認
    try {
      const currentUser = Session.getActiveUser().getEmail();
      results.checks.auth = {
        success: true,
        message: `認証機能OK: 現在のユーザー ${currentUser}`
      };
    } catch (error) {
      results.checks.auth = {
        success: false,
        message: '認証機能エラー',
        error: error.message
      };
      results.success = false;
      results.errors.push(error.message);
    }
    
    // 4. ユーティリティ関数確認
    try {
      const testTime = timeToMinutes('09:30');
      const testDate = formatDate(new Date());
      results.checks.utils = {
        success: true,
        message: `ユーティリティ関数OK: timeToMinutes(09:30)=${testTime}, formatDate=${testDate}`
      };
    } catch (error) {
      results.checks.utils = {
        success: false,
        message: 'ユーティリティ関数エラー',
        error: error.message
      };
      results.success = false;
      results.errors.push(error.message);
    }
    
    Logger.log('確認結果:', JSON.stringify(results, null, 2));
    Logger.log('=== 簡単なシステム確認完了 ===');
    
    return results;
    
  } catch (error) {
    Logger.log('システム確認エラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}