/**
 * クイックテストモジュール
 * システム全体の動作確認用の簡易テスト機能を提供
 */

/**
 * システム全体の動作確認を実行
 */
function quickSystemCheck() {
  Logger.log('=== システム全体動作確認開始 ===');
  const startTime = new Date();
  
  const results = {
    success: true,
    checks: {},
    errors: [],
    duration: 0
  };
  
  try {
    // 1. 基本設定の確認
    Logger.log('1. 基本設定の確認...');
    results.checks.config = testBasicConfiguration();
    
    // 2. スプレッドシート構造の確認
    Logger.log('2. スプレッドシート構造の確認...');
    results.checks.spreadsheet = testSpreadsheetStructure();
    
    // 3. 認証システムの確認
    Logger.log('3. 認証システムの確認...');
    results.checks.authentication = testAuthenticationSystem();
    
    // 4. エラーハンドリングの確認
    Logger.log('4. エラーハンドリングの確認...');
    results.checks.errorHandling = testErrorHandlingSystem();
    
    // 5. 基本機能の確認
    Logger.log('5. 基本機能の確認...');
    results.checks.basicFunctions = testBasicFunctions();
    
    // 結果の集計
    const allChecks = Object.values(results.checks);
    results.success = allChecks.every(check => check.success);
    results.duration = new Date().getTime() - startTime.getTime();
    
    // 結果のログ出力
    Logger.log('=== システム全体動作確認結果 ===');
    Logger.log(`実行時間: ${results.duration}ms`);
    Logger.log(`全体結果: ${results.success ? 'SUCCESS' : 'FAILED'}`);
    
    Object.entries(results.checks).forEach(([name, check]) => {
      Logger.log(`${name}: ${check.success ? 'PASS' : 'FAIL'} - ${check.message || 'No message'}`);
    });
    
    if (results.errors.length > 0) {
      Logger.log('エラー詳細:');
      results.errors.forEach(error => Logger.log(`- ${error}`));
    }
    
  } catch (error) {
    results.success = false;
    results.errors.push(error.toString());
    Logger.log(`システム確認中にエラーが発生: ${error.toString()}`);
  }
  
  return results;
}

/**
 * 基本設定の確認
 */
function testBasicConfiguration() {
  try {
    // Config.gsの基本設定確認
    const config = getConfig('BUSINESS');
    if (!config) {
      return { success: false, message: 'BUSINESS設定が見つかりません' };
    }
    
    // 必須設定項目の確認
    const requiredSettings = ['STANDARD_WORK_HOURS', 'OVERTIME_WARNING_THRESHOLD'];
    const missingSettings = requiredSettings.filter(setting => !config[setting]);
    
    if (missingSettings.length > 0) {
      return { 
        success: false, 
        message: `必須設定が不足: ${missingSettings.join(', ')}` 
      };
    }
    
    return { success: true, message: '基本設定が正常です' };
    
  } catch (error) {
    return { success: false, message: `設定確認エラー: ${error.message}` };
  }
}

/**
 * スプレッドシート構造の確認
 */
function testSpreadsheetStructure() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      return { success: false, message: 'アクティブなスプレッドシートが見つかりません' };
    }
    
    // 必須シートの確認
    const requiredSheets = ['Master_Employee', 'Log_Raw', 'Daily_Summary', 'Monthly_Summary'];
    const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    const missingSheets = requiredSheets.filter(sheet => !existingSheets.includes(sheet));
    
    if (missingSheets.length > 0) {
      return { 
        success: false, 
        message: `必須シートが不足: ${missingSheets.join(', ')}` 
      };
    }
    
    return { success: true, message: 'スプレッドシート構造が正常です' };
    
  } catch (error) {
    return { success: false, message: `スプレッドシート確認エラー: ${error.message}` };
  }
}

/**
 * 認証システムの確認
 */
function testAuthenticationSystem() {
  try {
    // 認証関数の存在確認
    if (typeof getEmployeeInfo !== 'function') {
      return { success: false, message: 'getEmployeeInfo関数が見つかりません' };
    }
    
    if (typeof authenticateUser !== 'function') {
      return { success: false, message: 'authenticateUser関数が見つかりません' };
    }
    
    // 基本的な認証テスト（エラーが発生しないことを確認）
    try {
      const testResult = getEmployeeInfo('TEST001');
      // テスト用IDなので存在しない可能性が高いが、関数自体は動作するはず
    } catch (error) {
      // 認証エラーは正常（テスト用IDなので）
      if (!error.message.includes('認証') && !error.message.includes('見つかりません')) {
        return { success: false, message: `認証システムエラー: ${error.message}` };
      }
    }
    
    return { success: true, message: '認証システムが正常です' };
    
  } catch (error) {
    return { success: false, message: `認証確認エラー: ${error.message}` };
  }
}

/**
 * エラーハンドリングシステムの確認
 */
function testErrorHandlingSystem() {
  try {
    // エラーハンドリング関数の存在確認
    if (typeof withErrorHandling !== 'function') {
      return { success: false, message: 'withErrorHandling関数が見つかりません' };
    }
    
    if (typeof logError !== 'function') {
      return { success: false, message: 'logError関数が見つかりません' };
    }
    
    // エラーハンドリングのテスト
    const testResult = withErrorHandling(() => {
      throw new Error('テストエラー');
    }, 'testErrorHandling', 'LOW');
    
    if (testResult.success) {
      return { success: false, message: 'エラーハンドリングが正常に動作していません' };
    }
    
    return { success: true, message: 'エラーハンドリングシステムが正常です' };
    
  } catch (error) {
    return { success: false, message: `エラーハンドリング確認エラー: ${error.message}` };
  }
}

/**
 * 基本機能の確認
 */
function testBasicFunctions() {
  try {
    // 基本関数の存在確認
    const requiredFunctions = [
      'calculateDailyWorkTime',
      'timeToMinutes',
      'minutesToTime',
      'formatDate'
    ];
    
    const missingFunctions = requiredFunctions.filter(funcName => typeof globalThis[funcName] !== 'function');
    
    if (missingFunctions.length > 0) {
      return { 
        success: false, 
        message: `基本関数が不足: ${missingFunctions.join(', ')}` 
      };
    }
    
    // 基本的な時間計算のテスト
    try {
      const minutes = timeToMinutes('09:00');
      const time = minutesToTime(540);
      
      if (minutes !== 540 || time !== '09:00') {
        return { success: false, message: '時間計算機能に問題があります' };
      }
    } catch (error) {
      return { success: false, message: `時間計算エラー: ${error.message}` };
    }
    
    return { success: true, message: '基本機能が正常です' };
    
  } catch (error) {
    return { success: false, message: `基本機能確認エラー: ${error.message}` };
  }
}

/**
 * 各モジュールの機能テスト実行
 */
function runModuleTests() {
  Logger.log('=== 各モジュール機能テスト実行開始 ===');
  const startTime = new Date();
  
  const results = {
    success: true,
    modules: {},
    errors: [],
    duration: 0
  };
  
  try {
    // 1. Authentication.gs
    Logger.log('1. Authentication.gs テスト...');
    results.modules.authentication = testAuthenticationModule();
    
    // 2. BusinessLogic.gs
    Logger.log('2. BusinessLogic.gs テスト...');
    results.modules.businessLogic = testBusinessLogicModule();
    
    // 3. Utils.gs
    Logger.log('3. Utils.gs テスト...');
    results.modules.utils = testUtilsModule();
    
    // 4. ErrorHandler.gs
    Logger.log('4. ErrorHandler.gs テスト...');
    results.modules.errorHandler = testErrorHandlerModule();
    
    // 5. WebApp.gs
    Logger.log('5. WebApp.gs テスト...');
    results.modules.webApp = testWebAppModule();
    
    // 結果の集計
    const allModules = Object.values(results.modules);
    results.success = allModules.every(module => module.success);
    results.duration = new Date().getTime() - startTime.getTime();
    
    // 結果のログ出力
    Logger.log('=== モジュールテスト結果 ===');
    Logger.log(`実行時間: ${results.duration}ms`);
    Logger.log(`全体結果: ${results.success ? 'SUCCESS' : 'FAILED'}`);
    
    Object.entries(results.modules).forEach(([name, module]) => {
      Logger.log(`${name}: ${module.success ? 'PASS' : 'FAIL'} - ${module.message || 'No message'}`);
    });
    
  } catch (error) {
    results.success = false;
    results.errors.push(error.toString());
    Logger.log(`モジュールテスト中にエラーが発生: ${error.toString()}`);
  }
  
  return results;
}

/**
 * Authentication.gs モジュールテスト
 */
function testAuthenticationModule() {
  try {
    // 認証関連関数の存在確認
    const authFunctions = ['getEmployeeInfo', 'authenticateUser', 'getEmployeeById'];
    const missingFunctions = authFunctions.filter(funcName => typeof globalThis[funcName] !== 'function');
    
    if (missingFunctions.length > 0) {
      return { 
        success: false, 
        message: `認証関数が不足: ${missingFunctions.join(', ')}` 
      };
    }
    
    return { success: true, message: 'Authentication.gs が正常です' };
    
  } catch (error) {
    return { success: false, message: `Authentication.gs エラー: ${error.message}` };
  }
}

/**
 * BusinessLogic.gs モジュールテスト
 */
function testBusinessLogicModule() {
  try {
    // ビジネスロジック関数の存在確認
    const businessFunctions = ['calculateDailyWorkTime', 'validateTimeEntry'];
    const missingFunctions = businessFunctions.filter(funcName => typeof globalThis[funcName] !== 'function');
    
    if (missingFunctions.length > 0) {
      return { 
        success: false, 
        message: `ビジネスロジック関数が不足: ${missingFunctions.join(', ')}` 
      };
    }
    
    return { success: true, message: 'BusinessLogic.gs が正常です' };
    
  } catch (error) {
    return { success: false, message: `BusinessLogic.gs エラー: ${error.message}` };
  }
}

/**
 * Utils.gs モジュールテスト
 */
function testUtilsModule() {
  try {
    // ユーティリティ関数の存在確認
    const utilFunctions = ['timeToMinutes', 'minutesToTime', 'formatDate', 'calculateTimeDifference'];
    const missingFunctions = utilFunctions.filter(funcName => typeof globalThis[funcName] !== 'function');
    
    if (missingFunctions.length > 0) {
      return { 
        success: false, 
        message: `ユーティリティ関数が不足: ${missingFunctions.join(', ')}` 
      };
    }
    
    return { success: true, message: 'Utils.gs が正常です' };
    
  } catch (error) {
    return { success: false, message: `Utils.gs エラー: ${error.message}` };
  }
}

/**
 * ErrorHandler.gs モジュールテスト
 */
function testErrorHandlerModule() {
  try {
    // エラーハンドリング関数の存在確認
    const errorFunctions = ['withErrorHandling', 'logError', 'registerFunction', 'unregisterFunction'];
    const missingFunctions = errorFunctions.filter(funcName => typeof globalThis[funcName] !== 'function');
    
    if (missingFunctions.length > 0) {
      return { 
        success: false, 
        message: `エラーハンドリング関数が不足: ${missingFunctions.join(', ')}` 
      };
    }
    
    return { success: true, message: 'ErrorHandler.gs が正常です' };
    
  } catch (error) {
    return { success: false, message: `ErrorHandler.gs エラー: ${error.message}` };
  }
}

/**
 * WebApp.gs モジュールテスト
 */
function testWebAppModule() {
  try {
    // WebApp関数の存在確認
    const webAppFunctions = ['doGet', 'doPost', 'processTimeEntry'];
    const missingFunctions = webAppFunctions.filter(funcName => typeof globalThis[funcName] !== 'function');
    
    if (missingFunctions.length > 0) {
      return { 
        success: false, 
        message: `WebApp関数が不足: ${missingFunctions.join(', ')}` 
      };
    }
    
    return { success: true, message: 'WebApp.gs が正常です' };
    
  } catch (error) {
    return { success: false, message: `WebApp.gs エラー: ${error.message}` };
  }
}

/**
 * エラーハンドリングシステムの動作確認
 */
function testErrorHandlingSystem() {
  Logger.log('=== エラーハンドリングシステム動作確認開始 ===');
  const startTime = new Date();
  
  const results = {
    success: true,
    tests: {},
    errors: [],
    duration: 0
  };
  
  try {
    // 1. 基本的なエラーハンドリングテスト
    Logger.log('1. 基本的なエラーハンドリングテスト...');
    results.tests.basicErrorHandling = testBasicErrorHandling();
    
    // 2. エラーログ機能テスト
    Logger.log('2. エラーログ機能テスト...');
    results.tests.errorLogging = testErrorLogging();
    
    // 3. 関数登録・削除機能テスト
    Logger.log('3. 関数登録・削除機能テスト...');
    results.tests.functionRegistration = testFunctionRegistration();
    
    // 4. エラー分類テスト
    Logger.log('4. エラー分類テスト...');
    results.tests.errorClassification = testErrorClassification();
    
    // 結果の集計
    const allTests = Object.values(results.tests);
    results.success = allTests.every(test => test.success);
    results.duration = new Date().getTime() - startTime.getTime();
    
    // 結果のログ出力
    Logger.log('=== エラーハンドリングテスト結果 ===');
    Logger.log(`実行時間: ${results.duration}ms`);
    Logger.log(`全体結果: ${results.success ? 'SUCCESS' : 'FAILED'}`);
    
    Object.entries(results.tests).forEach(([name, test]) => {
      Logger.log(`${name}: ${test.success ? 'PASS' : 'FAIL'} - ${test.message || 'No message'}`);
    });
    
  } catch (error) {
    results.success = false;
    results.errors.push(error.toString());
    Logger.log(`エラーハンドリングテスト中にエラーが発生: ${error.toString()}`);
  }
  
  return results;
}

/**
 * 基本的なエラーハンドリングテスト
 */
function testBasicErrorHandling() {
  try {
    // withErrorHandling関数のテスト
    const result = withErrorHandling(() => {
      throw new Error('テストエラー');
    }, 'testBasicErrorHandling', 'LOW');
    
    if (result.success) {
      return { success: false, message: 'エラーが正常に処理されていません' };
    }
    
    if (!result.error || !result.error.includes('テストエラー')) {
      return { success: false, message: 'エラーメッセージが正しく記録されていません' };
    }
    
    return { success: true, message: '基本的なエラーハンドリングが正常です' };
    
  } catch (error) {
    return { success: false, message: `基本エラーハンドリングテストエラー: ${error.message}` };
  }
}

/**
 * エラーログ機能テスト
 */
function testErrorLogging() {
  try {
    // logError関数のテスト
    const testError = new Error('ログテストエラー');
    const logResult = logError(testError, 'testErrorLogging', 'MEDIUM');
    
    if (!logResult) {
      return { success: false, message: 'エラーログ機能が正常に動作していません' };
    }
    
    return { success: true, message: 'エラーログ機能が正常です' };
    
  } catch (error) {
    return { success: false, message: `エラーログテストエラー: ${error.message}` };
  }
}

/**
 * 関数登録・削除機能テスト
 */
function testFunctionRegistration() {
  try {
    // テスト用関数
    const testFunc = function() { return 'test'; };
    
    // 関数登録テスト
    registerFunction('testRegistrationFunc', testFunc);
    
    // 登録確認
    if (!FUNCTION_MAPPING['testRegistrationFunc']) {
      return { success: false, message: '関数登録が正常に動作していません' };
    }
    
    // 関数削除テスト
    const deleteResult = unregisterFunction('testRegistrationFunc');
    
    if (!deleteResult) {
      return { success: false, message: '関数削除が正常に動作していません' };
    }
    
    // 削除確認
    if (FUNCTION_MAPPING['testRegistrationFunc']) {
      return { success: false, message: '関数削除が完了していません' };
    }
    
    return { success: true, message: '関数登録・削除機能が正常です' };
    
  } catch (error) {
    return { success: false, message: `関数登録テストエラー: ${error.message}` };
  }
}

/**
 * エラー分類テスト
 */
function testErrorClassification() {
  try {
    // ValidationErrorテスト
    try {
      throw new ValidationError('バリデーションエラーテスト');
    } catch (error) {
      if (!(error instanceof ValidationError)) {
        return { success: false, message: 'ValidationErrorが正しく分類されていません' };
      }
    }
    
    // NotImplementedErrorテスト
    try {
      throw new NotImplementedError('未実装エラーテスト');
    } catch (error) {
      if (!(error instanceof NotImplementedError)) {
        return { success: false, message: 'NotImplementedErrorが正しく分類されていません' };
      }
    }
    
    return { success: true, message: 'エラー分類が正常です' };
    
  } catch (error) {
    return { success: false, message: `エラー分類テストエラー: ${error.message}` };
  }
}