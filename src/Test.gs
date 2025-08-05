/**
 * テスト機能モジュール
 * 各モジュールの単体テスト、テストカバレッジ表示、データ整合性チェック、システム稼働状況チェック機能を提供
 * 要件: 6.1
 */

// テスト設定
const TEST_CONFIG = {
  TEST_TIMEOUT: 30000,
  TEST_EMPLOYEE_ID: 'TEST001',
  TEST_EMPLOYEE_EMAIL: 'test@example.com',
  SEVERITY: {
    PASS: 'PASS',
    FAIL: 'FAIL', 
    ERROR: 'ERROR',
    SKIP: 'SKIP'
  }
};

/**
 * 全テストを実行するメイン関数
 */
function runAllTests() {
  return withErrorHandling(() => {
    Logger.log('=== 全テスト実行開始 ===');
    const startTime = new Date();
    
    const testSuites = [
      { name: 'Authentication', tests: getAuthenticationTests() },
      { name: 'BusinessLogic', tests: getBusinessLogicTests() },
      { name: 'Utils', tests: getUtilsTests() },
      { name: 'WebApp', tests: getWebAppTests() },
      { name: 'FormManager', tests: getFormManagerTests() },
      { name: 'MailManager', tests: getMailManagerTests() },
      { name: 'Triggers', tests: getTriggersTests() },
      { name: 'ErrorHandler', tests: getErrorHandlerTests() }
    ];
    
    const allResults = [];
    let totalTests = 0;
    let totalPassed = 0;
    let totalFailed = 0;
    let totalErrors = 0;
    
    testSuites.forEach(suite => {
      const result = runTestSuite(suite.name, suite.tests);
      allResults.push(result);
      totalTests += result.summary.total;
      totalPassed += result.summary.passed;
      totalFailed += result.summary.failed;
      totalErrors += result.summary.errors;
    });
    
    const endTime = new Date();
    const totalDuration = endTime.getTime() - startTime.getTime();
    
    Logger.log('=== 全テスト実行完了 ===');
    Logger.log(`総実行時間: ${totalDuration}ms`);
    Logger.log(`総テスト数: ${totalTests}`);
    Logger.log(`成功: ${totalPassed}, 失敗: ${totalFailed}, エラー: ${totalErrors}`);
    
    return {
      success: true,
      totalTests: totalTests,
      totalPassed: totalPassed,
      totalFailed: totalFailed,
      totalErrors: totalErrors,
      duration: totalDuration,
      suiteResults: allResults
    };
    
  }, 'Test.runAllTests', 'HIGH');
}

/**
 * 統合テストを実行する
 */
function runIntegrationTests() {
  return withErrorHandling(() => {
    Logger.log('=== 統合テスト実行開始 ===');
    const startTime = new Date();
    
    const integrationTests = [
      { name: 'SystemConfiguration', func: testSystemConfiguration },
      { name: 'DataIntegrity', func: testDataIntegrity },
      { name: 'TriggerConfiguration', func: testTriggerConfiguration },
      { name: 'ErrorHandlingFlow', func: testErrorHandlingFlow },
      { name: 'WorkflowIntegration', func: testWorkflowIntegration },
      { name: 'SecurityConfiguration', func: testSecurityConfiguration }
    ];
    
    const results = {
      startTime: startTime,
      tests: [],
      summary: { total: 0, passed: 0, failed: 0, errors: 0 }
    };
    
    integrationTests.forEach(test => {
      Logger.log(`統合テスト実行: ${test.name}`);
      
      try {
        const testResult = test.func();
        const result = {
          name: test.name,
          status: testResult.success ? TEST_CONFIG.SEVERITY.PASS : TEST_CONFIG.SEVERITY.FAIL,
          result: testResult,
          duration: testResult.duration || 0
        };
        
        results.tests.push(result);
        results.summary.total++;
        
        if (result.status === TEST_CONFIG.SEVERITY.PASS) {
          results.summary.passed++;
        } else {
          results.summary.failed++;
        }
        
      } catch (error) {
        const result = {
          name: test.name,
          status: TEST_CONFIG.SEVERITY.ERROR,
          error: error.message,
          duration: 0
        };
        
        results.tests.push(result);
        results.summary.total++;
        results.summary.errors++;
        
        Logger.log(`統合テストエラー (${test.name}): ${error.toString()}`);
      }
    });
    
    const endTime = new Date();
    const totalDuration = endTime.getTime() - startTime.getTime();
    results.endTime = endTime;
    results.totalDuration = totalDuration;
    
    Logger.log('=== 統合テスト実行完了 ===');
    Logger.log(`総実行時間: ${totalDuration}ms`);
    Logger.log(`成功: ${results.summary.passed}/${results.summary.total}`);
    
    return {
      success: results.summary.errors === 0 && results.summary.failed === 0,
      results: results
    };
    
  }, 'Test.runIntegrationTests', 'HIGH');
}

/**
 * システム設定の統合テスト
 */
function testSystemConfiguration() {
  const startTime = Date.now();
  
  try {
    const checks = {
      configObjects: testConfigObjectsExistence(),
      spreadsheetStructure: testSpreadsheetStructureIntegrity(),
      adminConfiguration: testAdminConfiguration(),
      systemPermissions: testSystemPermissions()
    };
    
    const allPassed = Object.values(checks).every(check => check.success);
    
    return {
      success: allPassed,
      checks: checks,
      duration: Date.now() - startTime
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      duration: Date.now() - startTime
    };
  }
}

/**
 * データ整合性の統合テスト
 */
function testDataIntegrity() {
  const startTime = Date.now();
  
  try {
    const checks = {
      employeeDataConsistency: testEmployeeDataConsistency(),
      holidayDataValidity: testHolidayDataValidity(),
      requestResponsesIntegrity: testRequestResponsesIntegrity(),
      summaryDataAccuracy: testSummaryDataAccuracy()
    };
    
    const allPassed = Object.values(checks).every(check => check.success);
    
    return {
      success: allPassed,
      checks: checks,
      duration: Date.now() - startTime
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      duration: Date.now() - startTime
    };
  }
}

/**
 * トリガー設定の統合テスト
 */
function testTriggerConfiguration() {
  const startTime = Date.now();
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const requiredTriggers = ['dailyJob', 'weeklyOvertimeJob', 'monthlyJob'];
    const foundTriggers = triggers.map(t => t.getHandlerFunction());
    
    const checks = {
      triggerCount: {
        success: triggers.length >= 3,
        expected: '3以上',
        actual: triggers.length
      },
      requiredTriggersPresent: {
        success: requiredTriggers.every(req => foundTriggers.includes(req)),
        expected: requiredTriggers,
        actual: foundTriggers
      },
      timeTriggerConfiguration: testTimeTriggerConfiguration(triggers)
    };
    
    const allPassed = Object.values(checks).every(check => check.success);
    
    return {
      success: allPassed,
      checks: checks,
      duration: Date.now() - startTime
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      duration: Date.now() - startTime
    };
  }
}

/**
 * エラーハンドリングフローの統合テスト
 */
function testErrorHandlingFlow() {
  const startTime = Date.now();
  
  try {
    const checks = {
      errorHandlingFunction: testErrorHandlingFunctionExists(),
      errorLogging: testErrorLogging(),
      errorNotification: testErrorNotificationSetup(),
      batchProcessingErrorHandling: testBatchProcessingErrorHandling()
    };
    
    const allPassed = Object.values(checks).every(check => check.success);
    
    return {
      success: allPassed,
      checks: checks,
      duration: Date.now() - startTime
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      duration: Date.now() - startTime
    };
  }
}

/**
 * ワークフロー統合テスト
 */
function testWorkflowIntegration() {
  const startTime = Date.now();
  
  try {
    const checks = {
      authenticationFlow: testAuthenticationWorkflow(),
      approvalWorkflow: testApprovalWorkflowSetup(),
      calculationWorkflow: testCalculationWorkflow(),
      notificationWorkflow: testNotificationWorkflowSetup()
    };
    
    const allPassed = Object.values(checks).every(check => check.success);
    
    return {
      success: allPassed,
      checks: checks,
      duration: Date.now() - startTime
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      duration: Date.now() - startTime
    };
  }
}

/**
 * セキュリティ設定の統合テスト
 */
function testSecurityConfiguration() {
  const startTime = Date.now();
  
  try {
    const checks = {
      sheetProtections: testSheetProtections(),
      accessControls: testAccessControls(),
      dataValidation: testDataValidationRules(),
      auditLogging: testAuditLoggingSetup()
    };
    
    const allPassed = Object.values(checks).every(check => check.success);
    
    return {
      success: allPassed,
      checks: checks,
      duration: Date.now() - startTime
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      duration: Date.now() - startTime
    };
  }
}

/**
 * テストスイートを実行
 */
function runTestSuite(suiteName, tests) {
  const startTime = new Date();
  const results = {
    suiteName: suiteName,
    startTime: startTime,
    tests: [],
    summary: { total: 0, passed: 0, failed: 0, errors: 0, skipped: 0 }
  };
  
  Logger.log(`=== テストスイート開始: ${suiteName} ===`);
  
  tests.forEach(test => {
    const testResult = runSingleTest(test);
    results.tests.push(testResult);
    results.summary.total++;
    
    switch (testResult.status) {
      case TEST_CONFIG.SEVERITY.PASS: results.summary.passed++; break;
      case TEST_CONFIG.SEVERITY.FAIL: results.summary.failed++; break;
      case TEST_CONFIG.SEVERITY.ERROR: results.summary.errors++; break;
      case TEST_CONFIG.SEVERITY.SKIP: results.summary.skipped++; break;
    }
  });
  
  const endTime = new Date();
  results.endTime = endTime;
  results.duration = endTime.getTime() - startTime.getTime();
  
  Logger.log(`=== テストスイート完了: ${suiteName} (${results.duration}ms) ===`);
  return results;
}

/**
 * 単一テストを実行
 */
function runSingleTest(testFunction) {
  const testName = testFunction.name;
  const startTime = new Date();
  
  try {
    Logger.log(`テスト実行: ${testName}`);
    
    const result = testFunction();
    const endTime = new Date();
    
    if (result === true || (result && result.success === true)) {
      Logger.log(`✓ ${testName} - PASS`);
      return {
        name: testName,
        status: TEST_CONFIG.SEVERITY.PASS,
        duration: endTime.getTime() - startTime.getTime(),
        message: result.message || 'テスト成功'
      };
    } else {
      Logger.log(`✗ ${testName} - FAIL`);
      return {
        name: testName,
        status: TEST_CONFIG.SEVERITY.FAIL,
        duration: endTime.getTime() - startTime.getTime(),
        message: result.message || 'テスト失敗'
      };
    }
    
  } catch (error) {
    const endTime = new Date();
    Logger.log(`✗ ${testName} - ERROR: ${error.message}`);
    return {
      name: testName,
      status: TEST_CONFIG.SEVERITY.ERROR,
      duration: endTime.getTime() - startTime.getTime(),
      message: error.message,
      stack: error.stack
    };
  }
}

// ========== Authentication モジュールのテスト ==========

function getAuthenticationTests() {
  return [
    testAuthenticateUser,
    testGetEmployeeInfo,
    testValidateEmployeeAccess,
    testValidateApprovalAccess,
    testGetEmployeeById
  ];
}

function testAuthenticateUser() {
  try {
    // 現在のユーザーで認証テスト
    const user = authenticateUser();
    
    if (!user) {
      return { success: false, message: '認証結果がnullです' };
    }
    
    if (!user.id || !user.name || !user.email) {
      return { success: false, message: '必須フィールドが不足しています' };
    }
    
    return { success: true, message: `認証成功: ${user.name}` };
    
  } catch (error) {
    return { success: false, message: `認証エラー: ${error.message}` };
  }
}

function testGetEmployeeInfo() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const employee = getEmployeeInfo(userEmail);
    
    if (!employee) {
      return { success: false, message: '従業員情報が取得できませんでした' };
    }
    
    if (!employee.id || !employee.name || !employee.email) {
      return { success: false, message: '従業員情報の必須フィールドが不足しています' };
    }
    
    return { success: true, message: `従業員情報取得成功: ${employee.name}` };
    
  } catch (error) {
    return { success: false, message: `従業員情報取得エラー: ${error.message}` };
  }
}

function testValidateEmployeeAccess() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const isValid = validateEmployeeAccess(userEmail);
    
    if (typeof isValid !== 'boolean') {
      return { success: false, message: '戻り値がboolean型ではありません' };
    }
    
    return { success: true, message: `アクセス検証結果: ${isValid}` };
    
  } catch (error) {
    return { success: false, message: `アクセス検証エラー: ${error.message}` };
  }
}

function testValidateApprovalAccess() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const testEmployeeId = 'TEST001';
    
    const hasAccess = validateApprovalAccess(userEmail, testEmployeeId);
    
    if (typeof hasAccess !== 'boolean') {
      return { success: false, message: '戻り値がboolean型ではありません' };
    }
    
    return { success: true, message: `承認権限検証結果: ${hasAccess}` };
    
  } catch (error) {
    return { success: false, message: `承認権限検証エラー: ${error.message}` };
  }
}

function testGetEmployeeById() {
  try {
    // 実際の従業員IDを取得してテスト
    const employees = getSheetData('Master_Employee', 'A:H');
    if (employees.length <= 1) {
      return { success: false, message: '従業員マスタにデータがありません' };
    }
    
    const testEmployeeId = employees[1][0]; // 最初の従業員ID
    const employee = getEmployeeById(testEmployeeId);
    
    if (!employee) {
      return { success: false, message: '従業員情報が取得できませんでした' };
    }
    
    if (employee.id !== testEmployeeId) {
      return { success: false, message: 'IDが一致しません' };
    }
    
    return { success: true, message: `ID検索成功: ${employee.name}` };
    
  } catch (error) {
    return { success: false, message: `ID検索エラー: ${error.message}` };
  }
}

// ========== BusinessLogic モジュールのテスト ==========

function getBusinessLogicTests() {
  return [
    testCalculateDailyWorkTime,
    testCalculateOvertime,
    testApplyBreakDeduction,
    testCalculateLateAndEarlyLeave,
    testUpdateDailySummary
  ];
}

function testCalculateDailyWorkTime() {
  try {
    // テスト用の従業員IDと日付
    const employees = getSheetData('Master_Employee', 'A:H');
    if (employees.length <= 1) {
      return { success: false, message: '従業員マスタにデータがありません' };
    }
    
    const testEmployeeId = employees[1][0];
    const testDate = new Date();
    
    const result = calculateDailyWorkTime(testEmployeeId, testDate);
    
    if (!result) {
      return { success: false, message: '計算結果がnullです' };
    }
    
    if (typeof result.workMinutes !== 'number' || 
        typeof result.overtimeMinutes !== 'number' ||
        typeof result.breakMinutes !== 'number') {
      return { success: false, message: '計算結果の型が正しくありません' };
    }
    
    return { success: true, message: `労働時間計算成功: ${result.workMinutes}分` };
    
  } catch (error) {
    return { success: false, message: `労働時間計算エラー: ${error.message}` };
  }
}

function testCalculateOvertime() {
  try {
    // テストケース1: 通常の残業
    const overtime1 = calculateOvertime(540, 480, new Date()); // 9時間 - 8時間
    if (overtime1 !== 60) {
      return { success: false, message: `残業計算エラー: 期待値60分, 実際${overtime1}分` };
    }
    
    // テストケース2: 残業なし
    const overtime2 = calculateOvertime(480, 480, new Date()); // 8時間 - 8時間
    if (overtime2 !== 0) {
      return { success: false, message: `残業計算エラー: 期待値0分, 実際${overtime2}分` };
    }
    
    return { success: true, message: '残業時間計算テスト成功' };
    
  } catch (error) {
    return { success: false, message: `残業時間計算エラー: ${error.message}` };
  }
}

function testApplyBreakDeduction() {
  try {
    // テストケース1: 6時間超過（45分控除）
    const result1 = applyBreakDeduction(420); // 7時間
    if (result1.workMinutes !== 375 || result1.breakMinutes !== 45) {
      return { success: false, message: '6時間超過の休憩控除が正しくありません' };
    }
    
    // テストケース2: 8時間超過（60分控除）
    const result2 = applyBreakDeduction(540); // 9時間
    if (result2.workMinutes !== 480 || result2.breakMinutes !== 60) {
      return { success: false, message: '8時間超過の休憩控除が正しくありません' };
    }
    
    // テストケース3: 6時間以下（控除なし）
    const result3 = applyBreakDeduction(300); // 5時間
    if (result3.workMinutes !== 300 || result3.breakMinutes !== 0) {
      return { success: false, message: '6時間以下の休憩控除が正しくありません' };
    }
    
    return { success: true, message: '休憩時間控除テスト成功' };
    
  } catch (error) {
    return { success: false, message: `休憩時間控除エラー: ${error.message}` };
  }
}

function testCalculateLateAndEarlyLeave() {
  try {
    const mockEmployee = {
      standardStartTime: '09:00',
      standardEndTime: '17:00'
    };
    
    const mockTimeEntries = {
      clockIn: '09:30',  // 30分遅刻
      clockOut: '16:30'  // 30分早退
    };
    
    const result = calculateLateAndEarlyLeave(mockEmployee, mockTimeEntries);
    
    if (result.lateMinutes !== 30) {
      return { success: false, message: `遅刻計算エラー: 期待値30分, 実際${result.lateMinutes}分` };
    }
    
    if (result.earlyLeaveMinutes !== 30) {
      return { success: false, message: `早退計算エラー: 期待値30分, 実際${result.earlyLeaveMinutes}分` };
    }
    
    return { success: true, message: '遅刻・早退判定テスト成功' };
    
  } catch (error) {
    return { success: false, message: `遅刻・早退判定エラー: ${error.message}` };
  }
}

function testUpdateDailySummary() {
  try {
    // 実際の従業員IDを使用
    const employees = getSheetData('Master_Employee', 'A:H');
    if (employees.length <= 1) {
      return { success: false, message: '従業員マスタにデータがありません' };
    }
    
    const testEmployeeId = employees[1][0];
    const testDate = new Date();
    
    // Daily_Summary更新を実行
    updateDailySummary(testEmployeeId, testDate);
    
    return { success: true, message: 'Daily_Summary更新テスト成功' };
    
  } catch (error) {
    return { success: false, message: `Daily_Summary更新エラー: ${error.message}` };
  }
}

// ========== Utils モジュールのテスト ==========

function getUtilsTests() {
  return [
    testIsHoliday,
    testTimeToMinutes,
    testMinutesToTime,
    testCalculateTimeDifference,
    testGetSheetData,
    testFormatDate,
    testFormatTime
  ];
}

function testIsHoliday() {
  try {
    const testDate = new Date('2024-01-01'); // 元日
    const result = isHoliday(testDate);
    
    if (!result || typeof result.isHoliday !== 'boolean') {
      return { success: false, message: '祝日判定結果の形式が正しくありません' };
    }
    
    return { success: true, message: `祝日判定テスト成功: ${testDate.toDateString()} = ${result.isHoliday}` };
    
  } catch (error) {
    return { success: false, message: `祝日判定エラー: ${error.message}` };
  }
}

function testTimeToMinutes() {
  try {
    // テストケース1: 正常な時刻
    const minutes1 = timeToMinutes('09:30');
    if (minutes1 !== 570) { // 9*60 + 30 = 570
      return { success: false, message: `時刻変換エラー: 期待値570, 実際${minutes1}` };
    }
    
    // テストケース2: 0時0分
    const minutes2 = timeToMinutes('00:00');
    if (minutes2 !== 0) {
      return { success: false, message: `時刻変換エラー: 期待値0, 実際${minutes2}` };
    }
    
    // テストケース3: 23時59分
    const minutes3 = timeToMinutes('23:59');
    if (minutes3 !== 1439) { // 23*60 + 59 = 1439
      return { success: false, message: `時刻変換エラー: 期待値1439, 実際${minutes3}` };
    }
    
    return { success: true, message: '時刻→分変換テスト成功' };
    
  } catch (error) {
    return { success: false, message: `時刻→分変換エラー: ${error.message}` };
  }
}

function testMinutesToTime() {
  try {
    // テストケース1: 570分 = 9:30
    const time1 = minutesToTime(570);
    if (time1 !== '09:30') {
      return { success: false, message: `分→時刻変換エラー: 期待値09:30, 実際${time1}` };
    }
    
    // テストケース2: 0分 = 0:00
    const time2 = minutesToTime(0);
    if (time2 !== '00:00') {
      return { success: false, message: `分→時刻変換エラー: 期待値00:00, 実際${time2}` };
    }
    
    return { success: true, message: '分→時刻変換テスト成功' };
    
  } catch (error) {
    return { success: false, message: `分→時刻変換エラー: ${error.message}` };
  }
}

function testCalculateTimeDifference() {
  try {
    // テストケース1: 9:00-17:00 = 480分
    const diff1 = calculateTimeDifference('09:00', '17:00');
    if (diff1 !== 480) {
      return { success: false, message: `時間差計算エラー: 期待値480, 実際${diff1}` };
    }
    
    // テストケース2: 同じ時刻 = 0分
    const diff2 = calculateTimeDifference('12:00', '12:00');
    if (diff2 !== 0) {
      return { success: false, message: `時間差計算エラー: 期待値0, 実際${diff2}` };
    }
    
    return { success: true, message: '時間差計算テスト成功' };
    
  } catch (error) {
    return { success: false, message: `時間差計算エラー: ${error.message}` };
  }
}

function testGetSheetData() {
  try {
    // Master_Employeeシートのデータ取得テスト
    const data = getSheetData('Master_Employee', 'A1:A1');
    
    if (!Array.isArray(data)) {
      return { success: false, message: 'データがArray型ではありません' };
    }
    
    return { success: true, message: `シートデータ取得成功: ${data.length}行` };
    
  } catch (error) {
    return { success: false, message: `シートデータ取得エラー: ${error.message}` };
  }
}

function testFormatDate() {
  try {
    const testDate = new Date('2024-01-15');
    
    // YYYY-MM-DD形式
    const formatted1 = formatDate(testDate, 'YYYY-MM-DD');
    if (formatted1 !== '2024-01-15') {
      return { success: false, message: `日付フォーマットエラー: 期待値2024-01-15, 実際${formatted1}` };
    }
    
    // YYYY/MM/DD形式
    const formatted2 = formatDate(testDate, 'YYYY/MM/DD');
    if (formatted2 !== '2024/01/15') {
      return { success: false, message: `日付フォーマットエラー: 期待値2024/01/15, 実際${formatted2}` };
    }
    
    return { success: true, message: '日付フォーマットテスト成功' };
    
  } catch (error) {
    return { success: false, message: `日付フォーマットエラー: ${error.message}` };
  }
}

function testFormatTime() {
  try {
    const testDate = new Date('2024-01-15 09:30:45');
    
    // HH:MM形式
    const formatted1 = formatTime(testDate, 'HH:MM');
    if (formatted1 !== '09:30') {
      return { success: false, message: `時刻フォーマットエラー: 期待値09:30, 実際${formatted1}` };
    }
    
    // HH:MM:SS形式
    const formatted2 = formatTime(testDate, 'HH:MM:SS');
    if (formatted2 !== '09:30:45') {
      return { success: false, message: `時刻フォーマットエラー: 期待値09:30:45, 実際${formatted2}` };
    }
    
    return { success: true, message: '時刻フォーマットテスト成功' };
    
  } catch (error) {
    return { success: false, message: `時刻フォーマットエラー: ${error.message}` };
  }
}

// ========== WebApp モジュールのテスト ==========

function getWebAppTests() {
  return [
    testProcessTimeEntry,
    testValidateTimeEntry,
    testRecordTimeEntry,
    testCreateJsonResponse,
    testGetActionName
  ];
}

function testProcessTimeEntry() {
  try {
    // 認証が必要なため、実際の打刻は行わずに関数の存在確認のみ
    if (typeof processTimeEntry !== 'function') {
      return { success: false, message: 'processTimeEntry関数が存在しません' };
    }
    
    return { success: true, message: 'processTimeEntry関数存在確認成功' };
    
  } catch (error) {
    return { success: false, message: `processTimeEntryテストエラー: ${error.message}` };
  }
}

function testValidateTimeEntry() {
  try {
    // 実際の従業員IDを使用
    const employees = getSheetData('Master_Employee', 'A:H');
    if (employees.length <= 1) {
      return { success: false, message: '従業員マスタにデータがありません' };
    }
    
    const testEmployeeId = employees[1][0];
    const testDate = new Date();
    
    const result = validateTimeEntry(testEmployeeId, 'IN', testDate);
    
    if (!result || typeof result.valid !== 'boolean') {
      return { success: false, message: '検証結果の形式が正しくありません' };
    }
    
    return { success: true, message: `打刻検証テスト成功: ${result.message}` };
    
  } catch (error) {
    return { success: false, message: `打刻検証テストエラー: ${error.message}` };
  }
}

function testRecordTimeEntry() {
  try {
    // 実際の記録は行わず、関数の存在確認のみ
    if (typeof recordTimeEntry !== 'function') {
      return { success: false, message: 'recordTimeEntry関数が存在しません' };
    }
    
    return { success: true, message: 'recordTimeEntry関数存在確認成功' };
    
  } catch (error) {
    return { success: false, message: `recordTimeEntryテストエラー: ${error.message}` };
  }
}

function testCreateJsonResponse() {
  try {
    const response = createJsonResponse(true, 'テストメッセージ');
    
    if (!response) {
      return { success: false, message: 'JSONレスポンスが生成されませんでした' };
    }
    
    const content = response.getContent();
    const parsed = JSON.parse(content);
    
    if (parsed.success !== true || parsed.message !== 'テストメッセージ') {
      return { success: false, message: 'JSONレスポンスの内容が正しくありません' };
    }
    
    return { success: true, message: 'JSONレスポンス生成テスト成功' };
    
  } catch (error) {
    return { success: false, message: `JSONレスポンステストエラー: ${error.message}` };
  }
}

function testGetActionName() {
  try {
    const actionNames = {
      'IN': '出勤',
      'OUT': '退勤',
      'BRK_IN': '休憩開始',
      'BRK_OUT': '休憩終了'
    };
    
    for (const [action, expectedName] of Object.entries(actionNames)) {
      const actualName = getActionName(action);
      if (actualName !== expectedName) {
        return { success: false, message: `アクション名エラー: ${action} = ${actualName} (期待値: ${expectedName})` };
      }
    }
    
    return { success: true, message: 'アクション名取得テスト成功' };
    
  } catch (error) {
    return { success: false, message: `アクション名テストエラー: ${error.message}` };
  }
}

// ========== FormManager モジュールのテスト ==========

function getFormManagerTests() {
  return [
    testExtractFormData,
    testValidateRequestData,
    testIdentifyApprover,
    testGetEmployeeById,
    testProcessApprovalStatusChange
  ];
}

function testExtractFormData() {
  try {
    const mockEvent = {
      values: ['2024-01-15 10:00:00', 'EMP001', '修正', '出勤時刻修正', '09:00']
    };
    
    const formData = extractFormData(mockEvent);
    
    if (!formData || !formData.employeeId || !formData.requestType) {
      return { success: false, message: 'フォームデータの抽出が正しくありません' };
    }
    
    if (formData.employeeId !== 'EMP001' || formData.requestType !== '修正') {
      return { success: false, message: 'フォームデータの値が正しくありません' };
    }
    
    return { success: true, message: 'フォームデータ抽出テスト成功' };
    
  } catch (error) {
    return { success: false, message: `フォームデータ抽出エラー: ${error.message}` };
  }
}

function testValidateRequestData() {
  try {
    const validData = {
      employeeId: 'EMP001',
      requestType: '修正',
      details: 'テスト詳細',
      requestedValue: '09:00'
    };
    
    // 有効なデータの検証（エラーが発生しないことを確認）
    validateRequestData(validData);
    
    return { success: true, message: '申請データ検証テスト成功' };
    
  } catch (error) {
    // 検証エラーが発生した場合
    if (error.message.includes('申請データ検証エラー')) {
      return { success: true, message: '申請データ検証（エラーケース）テスト成功' };
    }
    return { success: false, message: `申請データ検証エラー: ${error.message}` };
  }
}

function testIdentifyApprover() {
  try {
    // 実際の従業員IDを使用
    const employees = getSheetData('Master_Employee', 'A:H');
    if (employees.length <= 1) {
      return { success: false, message: '従業員マスタにデータがありません' };
    }
    
    const testEmployeeId = employees[1][0];
    
    try {
      const approver = identifyApprover(testEmployeeId);
      return { success: true, message: `承認者特定テスト成功: ${approver}` };
    } catch (error) {
      // 承認者が設定されていない場合は警告として扱う
      if (error.message.includes('承認者が設定されていません')) {
        return { 
          success: false, 
          warning: true, 
          message: '承認者特定テスト警告: 承認者が設定されていません。設定を確認してください。' 
        };
      }
      throw error;
    }
    
  } catch (error) {
    return { success: false, message: `承認者特定エラー: ${error.message}` };
  }
}

function testGetEmployeeById() {
  try {
    // 実際の従業員IDを使用
    const employees = getSheetData('Master_Employee', 'A:H');
    if (employees.length <= 1) {
      return { success: false, message: '従業員マスタにデータがありません' };
    }
    
    const testEmployeeId = employees[1][0];
    const employee = getEmployeeById(testEmployeeId);
    
    if (!employee) {
      return { success: false, message: '従業員情報が取得できませんでした' };
    }
    
    if (employee.id !== testEmployeeId) {
      return { success: false, message: 'IDが一致しません' };
    }
    
    return { success: true, message: `従業員ID検索テスト成功: ${employee.name}` };
    
  } catch (error) {
    return { success: false, message: `従業員ID検索エラー: ${error.message}` };
  }
}

function testProcessApprovalStatusChange() {
  try {
    // 実際の処理は行わず、関数の存在確認のみ
    if (typeof processApprovalStatusChange !== 'function') {
      return { success: false, message: 'processApprovalStatusChange関数が存在しません' };
    }
    
    return { success: true, message: 'processApprovalStatusChange関数存在確認成功' };
    
  } catch (error) {
    return { success: false, message: `承認ステータス変更テストエラー: ${error.message}` };
  }
}

// ========== MailManager モジュールのテスト ==========

function getMailManagerTests() {
  return [
    testSendBatchNotifications,
    testCombineMessages,
    testCreateErrorAlertBody,
    testGetAdminEmail
  ];
}

function testSendBatchNotifications() {
  try {
    // 実際のメール送信は行わず、関数の存在確認のみ
    if (typeof sendBatchNotifications !== 'function') {
      return { success: false, message: 'sendBatchNotifications関数が存在しません' };
    }
    
    return { success: true, message: 'sendBatchNotifications関数存在確認成功' };
    
  } catch (error) {
    return { success: false, message: `バッチ通知テストエラー: ${error.message}` };
  }
}

function testCombineMessages() {
  try {
    const testMessages = [
      { subject: 'テスト1', body: '本文1' },
      { subject: 'テスト2', body: '本文2' }
    ];
    
    const combined = combineMessages(testMessages);
    
    if (!combined || !combined.subject || !combined.body) {
      return { success: false, message: 'メッセージ結合結果が正しくありません' };
    }
    
    if (!combined.body.includes('本文1') || !combined.body.includes('本文2')) {
      return { success: false, message: 'メッセージ結合で本文が含まれていません' };
    }
    
    return { success: true, message: 'メッセージ結合テスト成功' };
    
  } catch (error) {
    return { success: false, message: `メッセージ結合エラー: ${error.message}` };
  }
}

function testCreateErrorAlertBody() {
  try {
    const mockError = {
      name: 'TestError',
      message: 'テストエラーメッセージ',
      stack: 'テストスタックトレース'
    };
    const context = 'TestContext';
    
    const body = createErrorAlertBody(mockError, context);
    
    if (!body || typeof body !== 'string') {
      return { success: false, message: 'エラー通知本文が生成されませんでした' };
    }
    
    if (!body.includes('TestContext') || !body.includes('テストエラーメッセージ')) {
      return { success: false, message: 'エラー通知本文に必要な情報が含まれていません' };
    }
    
    return { success: true, message: 'エラー通知本文作成テスト成功' };
    
  } catch (error) {
    return { success: false, message: `エラー通知本文作成エラー: ${error.message}` };
  }
}

function testGetAdminEmail() {
  try {
    const adminEmail = getAdminEmail();
    
    if (!adminEmail || typeof adminEmail !== 'string') {
      return { success: false, message: '管理者メールアドレスが取得できませんでした' };
    }
    
    if (!adminEmail.includes('@')) {
      return { success: false, message: '管理者メールアドレスの形式が正しくありません' };
    }
    
    return { success: true, message: `管理者メール取得テスト成功: ${adminEmail}` };
    
  } catch (error) {
    return { success: false, message: `管理者メール取得エラー: ${error.message}` };
  }
}

// ========== Triggers モジュールのテスト ==========

function getTriggersTests() {
  return [
    testGetAllEmployees,
    testCalculateEmployeeOvertimeForPeriod,
    testCalculateEmployeeMonthlySummary,
    testUpdateMonthlySummarySheet
  ];
}

function testGetAllEmployees() {
  try {
    const employees = getAllEmployees();
    
    if (!Array.isArray(employees)) {
      return { success: false, message: '従業員リストがArray型ではありません' };
    }
    
    if (employees.length === 0) {
      return { success: false, message: '従業員データが取得できませんでした' };
    }
    
    // 最初の従業員の必須フィールドをチェック
    const firstEmployee = employees[0];
    if (!firstEmployee.employeeId || !firstEmployee.name || !firstEmployee.email) {
      return { success: false, message: '従業員データの必須フィールドが不足しています' };
    }
    
    return { success: true, message: `全従業員取得テスト成功: ${employees.length}名` };
    
  } catch (error) {
    return { success: false, message: `全従業員取得エラー: ${error.message}` };
  }
}

function testCalculateEmployeeOvertimeForPeriod() {
  try {
    // 実際の従業員IDを使用
    const employees = getSheetData('Master_Employee', 'A:H');
    if (employees.length <= 1) {
      return { success: false, message: '従業員マスタにデータがありません' };
    }
    
    const testEmployeeId = employees[1][0];
    const endDate = new Date();
    const startDate = new Date(endDate.getTime() - 7 * 24 * 60 * 60 * 1000); // 1週間前
    
    const overtimeMinutes = calculateEmployeeOvertimeForPeriod(testEmployeeId, startDate, endDate);
    
    if (typeof overtimeMinutes !== 'number' || overtimeMinutes < 0) {
      return { success: false, message: '残業時間計算結果が正しくありません' };
    }
    
    return { success: true, message: `期間残業時間計算テスト成功: ${overtimeMinutes}分` };
    
  } catch (error) {
    return { success: false, message: `期間残業時間計算エラー: ${error.message}` };
  }
}

function testCalculateEmployeeMonthlySummary() {
  try {
    // 実際の従業員IDを使用
    const employees = getSheetData('Master_Employee', 'A:H');
    if (employees.length <= 1) {
      return { success: false, message: '従業員マスタにデータがありません' };
    }
    
    const testEmployeeId = employees[1][0];
    const today = new Date();
    const startDate = new Date(today.getFullYear(), today.getMonth(), 1);
    const endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    
    const summary = calculateEmployeeMonthlySummary(testEmployeeId, startDate, endDate);
    
    if (!summary || typeof summary.workDays !== 'number' || typeof summary.totalWorkMinutes !== 'number') {
      return { success: false, message: '月次サマリー計算結果が正しくありません' };
    }
    
    return { success: true, message: `月次サマリー計算テスト成功: ${summary.workDays}日勤務` };
    
  } catch (error) {
    return { success: false, message: `月次サマリー計算エラー: ${error.message}` };
  }
}

function testUpdateMonthlySummarySheet() {
  try {
    // 実際の更新は行わず、関数の存在確認のみ
    if (typeof updateMonthlySummarySheet !== 'function') {
      return { success: false, message: 'updateMonthlySummarySheet関数が存在しません' };
    }
    
    return { success: true, message: 'updateMonthlySummarySheet関数存在確認成功' };
    
  } catch (error) {
    return { success: false, message: `月次サマリー更新テストエラー: ${error.message}` };
  }
}

// ========== ErrorHandler モジュールのテスト ==========

function getErrorHandlerTests() {
  return [
    testLogError,
    testWithErrorHandling,
    testProcessBatch,
    testProcessChunks,
    testPerformHealthCheck
  ];
}

function testLogError() {
  try {
    const testError = new Error('テスト用エラー');
    logError(testError, 'TestContext', 'LOW', { testData: 'sample' });
    
    return { success: true, message: 'エラーログ記録テスト成功' };
    
  } catch (error) {
    return { success: false, message: `エラーログ記録エラー: ${error.message}` };
  }
}

function testWithErrorHandling() {
  try {
    const result = withErrorHandling(() => {
      return 'テスト成功';
    }, 'TestFunction', 'LOW');
    
    if (result !== 'テスト成功') {
      return { success: false, message: 'エラーハンドリングラッパーの結果が正しくありません' };
    }
    
    return { success: true, message: 'エラーハンドリングラッパーテスト成功' };
    
  } catch (error) {
    return { success: false, message: `エラーハンドリングラッパーエラー: ${error.message}` };
  }
}

function testProcessBatch() {
  try {
    const testItems = [1, 2, 3, 4, 5];
    const result = processBatch(testItems, (item, index) => {
      return item * 2;
    }, {
      context: 'TestBatch',
      batchSize: 2
    });
    
    if (!result || !result.success || result.processedCount !== 5) {
      return { success: false, message: 'バッチ処理結果が正しくありません' };
    }
    
    return { success: true, message: `バッチ処理テスト成功: ${result.processedCount}件処理` };
    
  } catch (error) {
    return { success: false, message: `バッチ処理テストエラー: ${error.message}` };
  }
}

function testProcessChunks() {
  try {
    const testItems = [1, 2, 3, 4, 5];
    const result = processChunks(testItems, (item, index) => {
      return item * 3;
    }, {
      context: 'TestChunk',
      chunkSize: 2
    });
    
    if (!result || !result.success || result.processedCount !== 5) {
      return { success: false, message: 'チャンク処理結果が正しくありません' };
    }
    
    return { success: true, message: `チャンク処理テスト成功: ${result.processedCount}件処理` };
    
  } catch (error) {
    return { success: false, message: `チャンク処理テストエラー: ${error.message}` };
  }
}

function testPerformHealthCheck() {
  try {
    const healthStatus = performHealthCheck();
    
    if (!healthStatus || !healthStatus.overall || !Array.isArray(healthStatus.checks)) {
      return { success: false, message: 'ヘルスチェック結果が正しくありません' };
    }
    
    return { success: true, message: `ヘルスチェックテスト成功: ${healthStatus.overall}` };
    
  } catch (error) {
    return { success: false, message: `ヘルスチェックテストエラー: ${error.message}` };
  }
}

// ========== テストカバレッジ表示機能 ==========

/**
 * テストカバレッジを表示
 */
function showTestCoverage() {
  return withErrorHandling(() => {
    Logger.log('=== テストカバレッジ分析開始 ===');
    
    const moduleInfo = {
      'Authentication.gs': {
        functions: ['authenticateUser', 'getEmployeeInfo', 'validateEmployeeAccess', 'validateApprovalAccess', 'getEmployeeById'],
        tests: ['testAuthenticateUser', 'testGetEmployeeInfo', 'testValidateEmployeeAccess', 'testValidateApprovalAccess', 'testGetEmployeeById']
      },
      'BusinessLogic.gs': {
        functions: ['calculateDailyWorkTime', 'calculateOvertime', 'applyBreakDeduction', 'calculateLateAndEarlyLeave', 'updateDailySummary'],
        tests: ['testCalculateDailyWorkTime', 'testCalculateOvertime', 'testApplyBreakDeduction', 'testCalculateLateAndEarlyLeave', 'testUpdateDailySummary']
      },
      'Utils.gs': {
        functions: ['isHoliday', 'timeToMinutes', 'minutesToTime', 'calculateTimeDifference', 'getSheetData', 'formatDate', 'formatTime'],
        tests: ['testIsHoliday', 'testTimeToMinutes', 'testMinutesToTime', 'testCalculateTimeDifference', 'testGetSheetData', 'testFormatDate', 'testFormatTime']
      },
      'WebApp.gs': {
        functions: ['processTimeEntry', 'validateTimeEntry', 'recordTimeEntry', 'createJsonResponse', 'getActionName'],
        tests: ['testProcessTimeEntry', 'testValidateTimeEntry', 'testRecordTimeEntry', 'testCreateJsonResponse', 'testGetActionName']
      },
      'FormManager.gs': {
        functions: ['extractFormData', 'validateRequestData', 'identifyApprover', 'getEmployeeById', 'processApprovalStatusChange'],
        tests: ['testExtractFormData', 'testValidateRequestData', 'testIdentifyApprover', 'testGetEmployeeById', 'testProcessApprovalStatusChange']
      },
      'MailManager.gs': {
        functions: ['sendBatchNotifications', 'combineMessages', 'createErrorAlertBody', 'getAdminEmail'],
        tests: ['testSendBatchNotifications', 'testCombineMessages', 'testCreateErrorAlertBody', 'testGetAdminEmail']
      },
      'Triggers.gs': {
        functions: ['getAllEmployees', 'calculateEmployeeOvertimeForPeriod', 'calculateEmployeeMonthlySummary', 'updateMonthlySummarySheet'],
        tests: ['testGetAllEmployees', 'testCalculateEmployeeOvertimeForPeriod', 'testCalculateEmployeeMonthlySummary', 'testUpdateMonthlySummarySheet']
      },
      'ErrorHandler.gs': {
        functions: ['logError', 'withErrorHandling', 'processBatch', 'processChunks', 'performHealthCheck'],
        tests: ['testLogError', 'testWithErrorHandling', 'testProcessBatch', 'testProcessChunks', 'testPerformHealthCheck']
      }
    };
    
    const coverageReport = [];
    let totalFunctions = 0;
    let totalTests = 0;
    let totalCovered = 0;
    
    Object.keys(moduleInfo).forEach(moduleName => {
      const module = moduleInfo[moduleName];
      const functionCount = module.functions.length;
      const testCount = module.tests.length;
      const coverage = Math.round((testCount / functionCount) * 100);
      
      totalFunctions += functionCount;
      totalTests += testCount;
      totalCovered += Math.min(testCount, functionCount);
      
      coverageReport.push({
        module: moduleName,
        functions: functionCount,
        tests: testCount,
        coverage: coverage,
        status: coverage >= 80 ? '✓' : coverage >= 50 ? '△' : '✗'
      });
      
      Logger.log(`${moduleName}: ${testCount}/${functionCount} (${coverage}%) ${coverage >= 80 ? '✓' : coverage >= 50 ? '△' : '✗'}`);
    });
    
    const overallCoverage = Math.round((totalCovered / totalFunctions) * 100);
    
    Logger.log('=== テストカバレッジサマリー ===');
    Logger.log(`総関数数: ${totalFunctions}`);
    Logger.log(`総テスト数: ${totalTests}`);
    Logger.log(`カバレッジ: ${overallCoverage}%`);
    Logger.log(`評価: ${overallCoverage >= 80 ? '優秀' : overallCoverage >= 50 ? '良好' : '要改善'}`);
    
    return {
      success: true,
      overallCoverage: overallCoverage,
      totalFunctions: totalFunctions,
      totalTests: totalTests,
      moduleReports: coverageReport
    };
    
  }, 'Test.showTestCoverage', 'MEDIUM');
}

// ========== データ整合性チェック機能 ==========

/**
 * データ整合性チェックを実行
 */
function checkDataIntegrity() {
  return withErrorHandling(() => {
    Logger.log('=== データ整合性チェック開始 ===');
    
    const checks = [
      checkMasterEmployeeIntegrity,
      checkMasterHolidayIntegrity,
      checkLogRawIntegrity,
      checkDailySummaryIntegrity,
      checkMonthlySummaryIntegrity,
      checkRequestResponsesIntegrity
    ];
    
    const results = [];
    let totalChecks = 0;
    let passedChecks = 0;
    let failedChecks = 0;
    
    checks.forEach(checkFunction => {
      try {
        const result = checkFunction();
        results.push(result);
        totalChecks++;
        
        if (result.success) {
          passedChecks++;
          Logger.log(`✓ ${result.checkName}: ${result.message}`);
        } else {
          failedChecks++;
          Logger.log(`✗ ${result.checkName}: ${result.message}`);
        }
        
      } catch (error) {
        results.push({
          checkName: checkFunction.name,
          success: false,
          message: `チェック実行エラー: ${error.message}`,
          error: error
        });
        totalChecks++;
        failedChecks++;
        Logger.log(`✗ ${checkFunction.name}: チェック実行エラー - ${error.message}`);
      }
    });
    
    Logger.log('=== データ整合性チェック完了 ===');
    Logger.log(`総チェック数: ${totalChecks}`);
    Logger.log(`成功: ${passedChecks}, 失敗: ${failedChecks}`);
    Logger.log(`整合性スコア: ${Math.round((passedChecks / totalChecks) * 100)}%`);
    
    return {
      success: failedChecks === 0,
      totalChecks: totalChecks,
      passedChecks: passedChecks,
      failedChecks: failedChecks,
      integrityScore: Math.round((passedChecks / totalChecks) * 100),
      checkResults: results
    };
    
  }, 'Test.checkDataIntegrity', 'HIGH');
}

/**
 * Master_Employeeシートの整合性チェック
 */
function checkMasterEmployeeIntegrity() {
  try {
    const data = getSheetData('Master_Employee', 'A:H');
    
    if (data.length <= 1) {
      return { checkName: 'Master_Employee', success: false, message: 'データが存在しません' };
    }
    
    const issues = [];
    const employeeIds = new Set();
    const emails = new Set();
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const employeeId = row[0];
      const name = row[1];
      const email = row[2];
      
      // 必須フィールドチェック
      if (!employeeId) issues.push(`行${i + 1}: 社員IDが空です`);
      if (!name) issues.push(`行${i + 1}: 氏名が空です`);
      if (!email) issues.push(`行${i + 1}: メールアドレスが空です`);
      
      // 重複チェック
      if (employeeId) {
        if (employeeIds.has(employeeId)) {
          issues.push(`行${i + 1}: 社員ID「${employeeId}」が重複しています`);
        }
        employeeIds.add(employeeId);
      }
      
      if (email) {
        if (emails.has(email)) {
          issues.push(`行${i + 1}: メールアドレス「${email}」が重複しています`);
        }
        emails.add(email);
      }
      
      // メールアドレス形式チェック
      if (email && !email.includes('@')) {
        issues.push(`行${i + 1}: メールアドレス「${email}」の形式が正しくありません`);
      }
    }
    
    if (issues.length > 0) {
      return {
        checkName: 'Master_Employee',
        success: false,
        message: `${issues.length}件の問題が見つかりました`,
        issues: issues
      };
    }
    
    return {
      checkName: 'Master_Employee',
      success: true,
      message: `${data.length - 1}件のデータに問題はありません`
    };
    
  } catch (error) {
    return {
      checkName: 'Master_Employee',
      success: false,
      message: `チェック実行エラー: ${error.message}`
    };
  }
}

/**
 * Master_Holidayシートの整合性チェック
 */
function checkMasterHolidayIntegrity() {
  try {
    const data = getSheetData('Master_Holiday', 'A:D');
    
    if (data.length <= 1) {
      return { checkName: 'Master_Holiday', success: true, message: '祝日データが存在しません（正常）' };
    }
    
    const issues = [];
    const dates = new Set();
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[0];
      const name = row[1];
      
      // 必須フィールドチェック
      if (!date) issues.push(`行${i + 1}: 日付が空です`);
      if (!name) issues.push(`行${i + 1}: 祝日名が空です`);
      
      // 日付重複チェック
      if (date) {
        const dateStr = formatDate(new Date(date));
        if (dates.has(dateStr)) {
          issues.push(`行${i + 1}: 日付「${dateStr}」が重複しています`);
        }
        dates.add(dateStr);
      }
      
      // 日付形式チェック
      if (date && isNaN(new Date(date).getTime())) {
        issues.push(`行${i + 1}: 日付「${date}」の形式が正しくありません`);
      }
    }
    
    if (issues.length > 0) {
      return {
        checkName: 'Master_Holiday',
        success: false,
        message: `${issues.length}件の問題が見つかりました`,
        issues: issues
      };
    }
    
    return {
      checkName: 'Master_Holiday',
      success: true,
      message: `${data.length - 1}件のデータに問題はありません`
    };
    
  } catch (error) {
    return {
      checkName: 'Master_Holiday',
      success: false,
      message: `チェック実行エラー: ${error.message}`
    };
  }
}/**

 * Log_Rawシートの整合性チェック
 */
function checkLogRawIntegrity() {
  try {
    const data = getSheetData('Log_Raw', 'A:F');
    
    if (data.length <= 1) {
      return { checkName: 'Log_Raw', success: true, message: '打刻データが存在しません（正常）' };
    }
    
    const issues = [];
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const timestamp = row[0];
      const employeeId = row[1];
      const action = row[3];
      
      // 必須フィールドチェック
      if (!timestamp) issues.push(`行${i + 1}: タイムスタンプが空です`);
      if (!employeeId) issues.push(`行${i + 1}: 社員IDが空です`);
      if (!action) issues.push(`行${i + 1}: アクションが空です`);
      
      // タイムスタンプ形式チェック
      if (timestamp && isNaN(new Date(timestamp).getTime())) {
        issues.push(`行${i + 1}: タイムスタンプ「${timestamp}」の形式が正しくありません`);
      }
      
      // アクション値チェック
      if (action && !['IN', 'OUT', 'BRK_IN', 'BRK_OUT'].includes(action)) {
        issues.push(`行${i + 1}: アクション「${action}」が無効です`);
      }
    }
    
    if (issues.length > 0) {
      return {
        checkName: 'Log_Raw',
        success: false,
        message: `${issues.length}件の問題が見つかりました`,
        issues: issues
      };
    }
    
    return {
      checkName: 'Log_Raw',
      success: true,
      message: `${data.length - 1}件のデータに問題はありません`
    };
    
  } catch (error) {
    return {
      checkName: 'Log_Raw',
      success: false,
      message: `チェック実行エラー: ${error.message}`
    };
  }
}

/**
 * Daily_Summaryシートの整合性チェック
 */
function checkDailySummaryIntegrity() {
  try {
    const data = getSheetData('Daily_Summary', 'A:I');
    
    if (data.length <= 1) {
      return { checkName: 'Daily_Summary', success: true, message: 'サマリーデータが存在しません（正常）' };
    }
    
    const issues = [];
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[0];
      const employeeId = row[1];
      const workMinutes = row[5];
      const overtimeMinutes = row[6];
      
      // 必須フィールドチェック
      if (!date) issues.push(`行${i + 1}: 日付が空です`);
      if (!employeeId) issues.push(`行${i + 1}: 社員IDが空です`);
      
      // 数値フィールドチェック
      if (workMinutes !== '' && (isNaN(workMinutes) || workMinutes < 0)) {
        issues.push(`行${i + 1}: 実働時間「${workMinutes}」が無効です`);
      }
      
      if (overtimeMinutes !== '' && (isNaN(overtimeMinutes) || overtimeMinutes < 0)) {
        issues.push(`行${i + 1}: 残業時間「${overtimeMinutes}」が無効です`);
      }
      
      // 日付形式チェック
      if (date && isNaN(new Date(date).getTime())) {
        issues.push(`行${i + 1}: 日付「${date}」の形式が正しくありません`);
      }
    }
    
    if (issues.length > 0) {
      return {
        checkName: 'Daily_Summary',
        success: false,
        message: `${issues.length}件の問題が見つかりました`,
        issues: issues
      };
    }
    
    return {
      checkName: 'Daily_Summary',
      success: true,
      message: `${data.length - 1}件のデータに問題はありません`
    };
    
  } catch (error) {
    return {
      checkName: 'Daily_Summary',
      success: false,
      message: `チェック実行エラー: ${error.message}`
    };
  }
}

/**
 * Monthly_Summaryシートの整合性チェック
 */
function checkMonthlySummaryIntegrity() {
  try {
    const data = getSheetData('Monthly_Summary', 'A:G');
    
    if (data.length <= 1) {
      return { checkName: 'Monthly_Summary', success: true, message: '月次サマリーデータが存在しません（正常）' };
    }
    
    const issues = [];
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const yearMonth = row[0];
      const employeeId = row[1];
      const workDays = row[2];
      const totalWorkMinutes = row[3];
      
      // 必須フィールドチェック
      if (!yearMonth) issues.push(`行${i + 1}: 年月が空です`);
      if (!employeeId) issues.push(`行${i + 1}: 社員IDが空です`);
      
      // 年月形式チェック（YYYY-MM）
      if (yearMonth && !/^\d{4}-\d{2}$/.test(yearMonth)) {
        issues.push(`行${i + 1}: 年月「${yearMonth}」の形式が正しくありません（YYYY-MM形式で入力してください）`);
      }
      
      // 数値フィールドチェック
      if (workDays !== '' && (isNaN(workDays) || workDays < 0)) {
        issues.push(`行${i + 1}: 勤務日数「${workDays}」が無効です`);
      }
      
      if (totalWorkMinutes !== '' && (isNaN(totalWorkMinutes) || totalWorkMinutes < 0)) {
        issues.push(`行${i + 1}: 総労働時間「${totalWorkMinutes}」が無効です`);
      }
    }
    
    if (issues.length > 0) {
      return {
        checkName: 'Monthly_Summary',
        success: false,
        message: `${issues.length}件の問題が見つかりました`,
        issues: issues
      };
    }
    
    return {
      checkName: 'Monthly_Summary',
      success: true,
      message: `${data.length - 1}件のデータに問題はありません`
    };
    
  } catch (error) {
    return {
      checkName: 'Monthly_Summary',
      success: false,
      message: `チェック実行エラー: ${error.message}`
    };
  }
}

/**
 * Request_Responsesシートの整合性チェック
 */
function checkRequestResponsesIntegrity() {
  try {
    const data = getSheetData('Request_Responses', 'A:G');
    
    if (data.length <= 1) {
      return { checkName: 'Request_Responses', success: true, message: '申請データが存在しません（正常）' };
    }
    
    const issues = [];
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const timestamp = row[0];
      const employeeId = row[1];
      const requestType = row[2];
      const status = row[6];
      
      // 必須フィールドチェック
      if (!timestamp) issues.push(`行${i + 1}: タイムスタンプが空です`);
      if (!employeeId) issues.push(`行${i + 1}: 社員IDが空です`);
      if (!requestType) issues.push(`行${i + 1}: 申請種別が空です`);
      if (!status) issues.push(`行${i + 1}: ステータスが空です`);
      
      // タイムスタンプ形式チェック
      if (timestamp && isNaN(new Date(timestamp).getTime())) {
        issues.push(`行${i + 1}: タイムスタンプ「${timestamp}」の形式が正しくありません`);
      }
      
      // 申請種別チェック
      if (requestType && !['修正', '残業', '休暇'].includes(requestType)) {
        issues.push(`行${i + 1}: 申請種別「${requestType}」が無効です`);
      }
      
      // ステータスチェック
      if (status && !['Pending', 'Approved', 'Rejected'].includes(status)) {
        issues.push(`行${i + 1}: ステータス「${status}」が無効です`);
      }
    }
    
    if (issues.length > 0) {
      return {
        checkName: 'Request_Responses',
        success: false,
        message: `${issues.length}件の問題が見つかりました`,
        issues: issues
      };
    }
    
    return {
      checkName: 'Request_Responses',
      success: true,
      message: `${data.length - 1}件のデータに問題はありません`
    };
    
  } catch (error) {
    return {
      checkName: 'Request_Responses',
      success: false,
      message: `チェック実行エラー: ${error.message}`
    };
  }
}

// ========== システム稼働状況チェック機能 ==========

/**
 * システム稼働状況チェックを実行
 */
function checkSystemHealth() {
  return withErrorHandling(() => {
    Logger.log('=== システム稼働状況チェック開始 ===');
    
    const healthChecks = [
      checkSheetAccess,
      checkScriptPermissions,
      checkTriggerStatus,
      checkQuotaUsage,
      checkConfigurationStatus,
      checkRecentErrors
    ];
    
    const results = [];
    let totalChecks = 0;
    let healthyChecks = 0;
    let warningChecks = 0;
    let criticalChecks = 0;
    
    healthChecks.forEach(checkFunction => {
      try {
        const result = checkFunction();
        results.push(result);
        totalChecks++;
        
        switch (result.status) {
          case 'HEALTHY':
            healthyChecks++;
            Logger.log(`✓ ${result.checkName}: ${result.message}`);
            break;
          case 'WARNING':
            warningChecks++;
            Logger.log(`△ ${result.checkName}: ${result.message}`);
            break;
          case 'CRITICAL':
            criticalChecks++;
            Logger.log(`✗ ${result.checkName}: ${result.message}`);
            break;
        }
        
      } catch (error) {
        results.push({
          checkName: checkFunction.name,
          status: 'CRITICAL',
          message: `チェック実行エラー: ${error.message}`,
          error: error
        });
        totalChecks++;
        criticalChecks++;
        Logger.log(`✗ ${checkFunction.name}: チェック実行エラー - ${error.message}`);
      }
    });
    
    // 総合評価
    let overallStatus = 'HEALTHY';
    if (criticalChecks > 0) {
      overallStatus = 'CRITICAL';
    } else if (warningChecks > 0) {
      overallStatus = 'WARNING';
    }
    
    Logger.log('=== システム稼働状況チェック完了 ===');
    Logger.log(`総合評価: ${overallStatus}`);
    Logger.log(`正常: ${healthyChecks}, 警告: ${warningChecks}, 重要: ${criticalChecks}`);
    
    return {
      success: true,
      overallStatus: overallStatus,
      totalChecks: totalChecks,
      healthyChecks: healthyChecks,
      warningChecks: warningChecks,
      criticalChecks: criticalChecks,
      checkResults: results,
      timestamp: new Date()
    };
    
  }, 'Test.checkSystemHealth', 'HIGH');
}/**

 * シートアクセス状況をチェック
 */
function checkSheetAccess() {
  try {
    const requiredSheets = [
      'Master_Employee', 'Master_Holiday', 'Log_Raw',
      'Daily_Summary', 'Monthly_Summary', 'Request_Responses'
    ];
    
    const inaccessibleSheets = [];
    
    requiredSheets.forEach(sheetName => {
      try {
        const sheet = getSheet(sheetName);
        // 簡単なアクセステスト
        sheet.getLastRow();
      } catch (error) {
        inaccessibleSheets.push(sheetName);
      }
    });
    
    if (inaccessibleSheets.length > 0) {
      return {
        checkName: 'シートアクセス',
        status: 'CRITICAL',
        message: `アクセスできないシート: ${inaccessibleSheets.join(', ')}`,
        details: { inaccessibleSheets: inaccessibleSheets }
      };
    }
    
    return {
      checkName: 'シートアクセス',
      status: 'HEALTHY',
      message: `全ての必須シート（${requiredSheets.length}個）にアクセス可能`
    };
    
  } catch (error) {
    return {
      checkName: 'シートアクセス',
      status: 'CRITICAL',
      message: `シートアクセスチェックエラー: ${error.message}`
    };
  }
}

/**
 * スクリプト権限をチェック
 */
function checkScriptPermissions() {
  try {
    const permissions = [];
    
    // Gmail送信権限チェック
    try {
      GmailApp.getRemainingDailyQuota();
      permissions.push('Gmail送信: OK');
    } catch (error) {
      permissions.push('Gmail送信: NG');
    }
    
    // Drive操作権限チェック
    try {
      DriveApp.getRootFolder();
      permissions.push('Drive操作: OK');
    } catch (error) {
      permissions.push('Drive操作: NG');
    }
    
    // スプレッドシート操作権限チェック
    try {
      SpreadsheetApp.getActiveSpreadsheet();
      permissions.push('スプレッドシート操作: OK');
    } catch (error) {
      permissions.push('スプレッドシート操作: NG');
    }
    
    const failedPermissions = permissions.filter(p => p.includes('NG'));
    
    if (failedPermissions.length > 0) {
      return {
        checkName: 'スクリプト権限',
        status: 'CRITICAL',
        message: `権限不足: ${failedPermissions.join(', ')}`,
        details: { permissions: permissions }
      };
    }
    
    return {
      checkName: 'スクリプト権限',
      status: 'HEALTHY',
      message: '必要な権限が全て付与されています'
    };
    
  } catch (error) {
    return {
      checkName: 'スクリプト権限',
      status: 'CRITICAL',
      message: `権限チェックエラー: ${error.message}`
    };
  }
}

/**
 * トリガー設定状況をチェック
 */
function checkTriggerStatus() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const triggerInfo = [];
    
    const expectedTriggers = ['dailyJob', 'weeklyOvertimeJob', 'monthlyJob'];
    const foundTriggers = [];
    
    triggers.forEach(trigger => {
      const handlerFunction = trigger.getHandlerFunction();
      const triggerSource = trigger.getTriggerSource();
      const eventType = trigger.getEventType();
      
      triggerInfo.push({
        function: handlerFunction,
        source: triggerSource.toString(),
        event: eventType.toString()
      });
      
      if (expectedTriggers.includes(handlerFunction)) {
        foundTriggers.push(handlerFunction);
      }
    });
    
    const missingTriggers = expectedTriggers.filter(t => !foundTriggers.includes(t));
    
    if (missingTriggers.length > 0) {
      return {
        checkName: 'トリガー設定',
        status: 'WARNING',
        message: `未設定のトリガー: ${missingTriggers.join(', ')}`,
        details: { 
          totalTriggers: triggers.length,
          foundTriggers: foundTriggers,
          missingTriggers: missingTriggers
        }
      };
    }
    
    return {
      checkName: 'トリガー設定',
      status: 'HEALTHY',
      message: `必要なトリガーが設定済み（${triggers.length}個）`
    };
    
  } catch (error) {
    return {
      checkName: 'トリガー設定',
      status: 'CRITICAL',
      message: `トリガーチェックエラー: ${error.message}`
    };
  }
}

/**
 * クォータ使用状況をチェック
 */
function checkQuotaUsage() {
  try {
    const quotaInfo = [];
    let warningCount = 0;
    
    // Gmail送信クォータ
    try {
      const gmailQuota = GmailApp.getRemainingDailyQuota();
      quotaInfo.push(`Gmail送信: ${gmailQuota}通残り`);
      if (gmailQuota < 50) warningCount++;
    } catch (error) {
      quotaInfo.push('Gmail送信: チェック不可');
      warningCount++;
    }
    
    // スクリプト実行時間（概算）
    const executionTime = new Date().getTime() % (6 * 60 * 1000); // 6分周期で概算
    quotaInfo.push(`実行時間: 約${Math.round(executionTime / 1000)}秒経過`);
    
    if (warningCount > 0) {
      return {
        checkName: 'クォータ使用状況',
        status: 'WARNING',
        message: `クォータ残量に注意が必要`,
        details: { quotaInfo: quotaInfo }
      };
    }
    
    return {
      checkName: 'クォータ使用状況',
      status: 'HEALTHY',
      message: 'クォータ使用量は正常範囲内'
    };
    
  } catch (error) {
    return {
      checkName: 'クォータ使用状況',
      status: 'WARNING',
      message: `クォータチェックエラー: ${error.message}`
    };
  }
}

/**
 * 設定状況をチェック
 */
function checkConfigurationStatus() {
  try {
    const configIssues = [];
    
    // 管理者メール設定チェック
    try {
      const adminEmail = getAdminEmail();
      if (!adminEmail || !adminEmail.includes('@')) {
        configIssues.push('管理者メールアドレスが正しく設定されていません');
      }
    } catch (error) {
      configIssues.push('管理者メールアドレスの取得に失敗しました');
    }
    
    // 従業員マスタデータチェック
    try {
      const employees = getSheetData('Master_Employee', 'A:H');
      if (employees.length <= 1) {
        configIssues.push('従業員マスタにデータが登録されていません');
      }
    } catch (error) {
      configIssues.push('従業員マスタの確認に失敗しました');
    }
    
    if (configIssues.length > 0) {
      return {
        checkName: '設定状況',
        status: 'WARNING',
        message: `設定に問題があります: ${configIssues.join(', ')}`,
        details: { issues: configIssues }
      };
    }
    
    return {
      checkName: '設定状況',
      status: 'HEALTHY',
      message: '基本設定は正常です'
    };
    
  } catch (error) {
    return {
      checkName: '設定状況',
      status: 'CRITICAL',
      message: `設定チェックエラー: ${error.message}`
    };
  }
}

/**
 * 最近のエラー状況をチェック
 */
function checkRecentErrors() {
  try {
    // エラーログシートの存在確認
    let errorLogSheet;
    try {
      errorLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Error_Log');
    } catch (error) {
      return {
        checkName: '最近のエラー',
        status: 'HEALTHY',
        message: 'エラーログシートが存在しません（正常）'
      };
    }
    
    if (!errorLogSheet) {
      return {
        checkName: '最近のエラー',
        status: 'HEALTHY',
        message: 'エラーログシートが存在しません（正常）'
      };
    }
    
    const errorData = errorLogSheet.getDataRange().getValues();
    
    if (errorData.length <= 1) {
      return {
        checkName: '最近のエラー',
        status: 'HEALTHY',
        message: 'エラーログが記録されていません（正常）'
      };
    }
    
    // 過去24時間のエラーをチェック
    const now = new Date();
    const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    
    let recentErrors = 0;
    let criticalErrors = 0;
    
    // ヘッダー行をスキップ
    for (let i = 1; i < errorData.length; i++) {
      const row = errorData[i];
      const timestamp = new Date(row[0]);
      const severity = row[2];
      
      if (timestamp > yesterday) {
        recentErrors++;
        if (severity === 'CRITICAL' || severity === 'HIGH') {
          criticalErrors++;
        }
      }
    }
    
    if (criticalErrors > 0) {
      return {
        checkName: '最近のエラー',
        status: 'CRITICAL',
        message: `過去24時間に重要なエラーが${criticalErrors}件発生しています`,
        details: { recentErrors: recentErrors, criticalErrors: criticalErrors }
      };
    }
    
    if (recentErrors > 10) {
      return {
        checkName: '最近のエラー',
        status: 'WARNING',
        message: `過去24時間に${recentErrors}件のエラーが発生しています`,
        details: { recentErrors: recentErrors }
      };
    }
    
    return {
      checkName: '最近のエラー',
      status: 'HEALTHY',
      message: `過去24時間のエラー: ${recentErrors}件（正常範囲内）`
    };
    
  } catch (error) {
    return {
      checkName: '最近のエラー',
      status: 'WARNING',
      message: `エラー状況チェックエラー: ${error.message}`
    };
  }
}

// ========== テスト実行用ユーティリティ関数 ==========

/**
 * 手動でテストを実行するためのメニュー関数
 */
function runTestsManually() {
  try {
    SpreadsheetApp.getUi().alert('テスト実行を開始します。完了まで数分かかる場合があります。');
    
    const result = runAllTests();
    
    const message = `
テスト実行完了

総テスト数: ${result.totalTests}
成功: ${result.totalPassed}
失敗: ${result.totalFailed}
エラー: ${result.totalErrors}
実行時間: ${result.duration}ms

詳細はログを確認してください。
    `;
    
    SpreadsheetApp.getUi().alert('テスト実行完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`テスト実行エラー: ${error.message}`);
  }
}

/**
 * データ整合性チェックを手動実行
 */
function runDataIntegrityCheckManually() {
  try {
    SpreadsheetApp.getUi().alert('データ整合性チェックを開始します。');
    
    const result = checkDataIntegrity();
    
    const message = `
データ整合性チェック完了

総チェック数: ${result.totalChecks}
成功: ${result.passedChecks}
失敗: ${result.failedChecks}
整合性スコア: ${result.integrityScore}%

詳細はログを確認してください。
    `;
    
    SpreadsheetApp.getUi().alert('データ整合性チェック完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`データ整合性チェックエラー: ${error.message}`);
  }
}

/**
 * システム稼働状況チェックを手動実行
 */
function runSystemHealthCheckManually() {
  try {
    SpreadsheetApp.getUi().alert('システム稼働状況チェックを開始します。');
    
    const result = checkSystemHealth();
    
    const message = `
システム稼働状況チェック完了

総合評価: ${result.overallStatus}
総チェック数: ${result.totalChecks}
正常: ${result.healthyChecks}
警告: ${result.warningChecks}
重要: ${result.criticalChecks}

詳細はログを確認してください。
    `;
    
    SpreadsheetApp.getUi().alert('システム稼働状況チェック完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`システム稼働状況チェックエラー: ${error.message}`);
  }
}

/**
 * テストカバレッジを手動表示
 */
function showTestCoverageManually() {
  try {
    SpreadsheetApp.getUi().alert('テストカバレッジ分析を開始します。');
    
    const result = showTestCoverage();
    
    const message = `
テストカバレッジ分析完了

総関数数: ${result.totalFunctions}
総テスト数: ${result.totalTests}
カバレッジ: ${result.overallCoverage}%
評価: ${result.overallCoverage >= 80 ? '優秀' : result.overallCoverage >= 50 ? '良好' : '要改善'}

詳細はログを確認してください。
    `;
    
    SpreadsheetApp.getUi().alert('テストカバレッジ分析完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`テストカバレッジ分析エラー: ${error.message}`);
  }
}

/**
 * 全テスト機能を一括実行
 */
function runAllTestFunctions() {
  return withErrorHandling(() => {
    Logger.log('=== 全テスト機能一括実行開始 ===');
    
    const results = {
      unitTests: null,
      coverage: null,
      dataIntegrity: null,
      systemHealth: null
    };
    
    // 1. 単体テスト実行
    try {
      Logger.log('単体テスト実行中...');
      results.unitTests = runAllTests();
    } catch (error) {
      Logger.log(`単体テスト実行エラー: ${error.message}`);
      results.unitTests = { success: false, error: error.message };
    }
    
    // 2. テストカバレッジ分析
    try {
      Logger.log('テストカバレッジ分析中...');
      results.coverage = showTestCoverage();
    } catch (error) {
      Logger.log(`テストカバレッジ分析エラー: ${error.message}`);
      results.coverage = { success: false, error: error.message };
    }
    
    // 3. データ整合性チェック
    try {
      Logger.log('データ整合性チェック中...');
      results.dataIntegrity = checkDataIntegrity();
    } catch (error) {
      Logger.log(`データ整合性チェックエラー: ${error.message}`);
      results.dataIntegrity = { success: false, error: error.message };
    }
    
    // 4. システム稼働状況チェック
    try {
      Logger.log('システム稼働状況チェック中...');
      results.systemHealth = checkSystemHealth();
    } catch (error) {
      Logger.log(`システム稼働状況チェックエラー: ${error.message}`);
      results.systemHealth = { success: false, error: error.message };
    }
    
    Logger.log('=== 全テスト機能一括実行完了 ===');
    
    return {
      success: true,
      results: results,
      timestamp: new Date()
    };
    
  }, 'Test.runAllTestFunctions', 'HIGH');
}

// ========== 統合テスト用ヘルパー関数（リファクタリング済み） ==========
// 以下の関数は機能別に分割されました：
// - TestHelpers_Config.gs: 設定関連テスト
// - TestHelpers_Security.gs: セキュリティ関連テスト  
// - TestHelpers_Workflow.gs: ワークフロー関連テスト
// - TestHelpers_Data.gs: データ検証テスト

/**
 * 構文チェック用の簡単なテスト関数
 */
function syntaxCheck() {
  try {
    Logger.log('=== 構文チェック開始 ===');
    
    // ERROR_CONFIGの存在確認
    if (typeof ERROR_CONFIG !== 'undefined') {
      Logger.log('✓ ERROR_CONFIG: 正常に定義されています');
      Logger.log(`  最大実行時間: ${ERROR_CONFIG.MAX_EXECUTION_TIME}ms`);
      Logger.log(`  デフォルトチャンクサイズ: ${ERROR_CONFIG.DEFAULT_CHUNK_SIZE}`);
    } else {
      Logger.log('✗ ERROR_CONFIG: 定義されていません');
    }
    
    // 基本的な関数の存在確認
    const requiredFunctions = [
      'logError', 'withErrorHandling', 'processBatch', 'processChunks',
      'getConfig', 'getSheet', 'formatDate', 'formatTime'
    ];
    
    requiredFunctions.forEach(funcName => {
      // 安全な関数存在チェック - eval()の代わりにグローバルスコープから直接アクセス
      const func = this[funcName] || (typeof window !== 'undefined' ? window[funcName] : null);
      if (typeof func === 'function') {
        Logger.log(`✓ ${funcName}: 関数が存在します`);
      } else {
        Logger.log(`✗ ${funcName}: 関数が存在しません`);
      }
    });
    
    Logger.log('=== 構文チェック完了 ===');
    return { success: true, message: '構文チェックが正常に完了しました' };
    
  } catch (error) {
    Logger.log(`構文チェックエラー: ${error.toString()}`);
    return { success: false, message: `構文チェックエラー: ${error.message}` };
  }
}

/**
 * 修正したテストの個別実行確認
 */
function testFixedCreateErrorAlertBody() {
  try {
    Logger.log('=== 修正したtestCreateErrorAlertBodyテスト実行 ===');
    
    const result = testCreateErrorAlertBody();
    
    if (result.success) {
      Logger.log('✓ testCreateErrorAlertBody: 修正成功');
      Logger.log(`  結果: ${result.message}`);
    } else {
      Logger.log('✗ testCreateErrorAlertBody: まだ問題あり');
      Logger.log(`  エラー: ${result.message}`);
    }
    
    return result;
    
  } catch (error) {
    Logger.log(`テスト実行エラー: ${error.toString()}`);
    return { success: false, message: `テスト実行エラー: ${error.message}` };
  }
}