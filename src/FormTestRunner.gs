/**
 * FormManager.gs のテストランナー
 * TDD実装: Red-Green-Refactorサイクル支援
 * 
 * 機能:
 * - FormManagerTest.gsの全テスト実行
 * - 個別テスト実行
 * - パフォーマンステスト
 * - テスト結果レポート
 */

/**
 * FormManager TDDテスト実行
 * 全テストケースを実行し、結果をレポート
 */
function runFormManagerTests() {
  // システム設定の初期化
  initializeSystemConfigIfNeeded();
  
  // テストケースの定義
  var testFunctions = [
    'testProcessFormResponse_ValidData_ReturnsSuccess',
    'testProcessFormResponse_InvalidData_ReturnsError',
    'testConvertFormDataToLogRaw_ValidData_ReturnsCorrectFormat',
    'testValidateFormData_ValidData_ReturnsTrue',
    'testValidateFormData_InvalidData_ReturnsFalse',
    'testSaveFormDataToSheet_ValidData_SavesSuccessfully',
    'testCheckDuplicateFormResponse_NewData_ReturnsFalse',
    'testCheckDuplicateFormResponse_DuplicateData_ReturnsTrue',
    'testProcessFormResponse_IntegrationTest_CompleteWorkflow',
    'testProcessFormResponse_ExceptionHandling_ReturnsError',
    'testProcessMultipleFormResponses_ValidData_ProcessesAll',
    'testGetFormResponseStats_ValidData_ReturnsStats'
  ];
  
  // 動的にテスト数を計算
  var testCount = testFunctions.length;
  
  console.log('=== FormManager TDDテスト実行開始 ===');
  console.log('実行日時: ' + new Date().toLocaleString('ja-JP'));
  console.log('実行テスト数: ' + testCount);
  console.log('---');
  
  testFunctions.forEach(function(testFunction) {
    try {
      var testStartTime = new Date().getTime();
      eval(testFunction + '()');
      var testEndTime = new Date().getTime();
      var testDuration = testEndTime - testStartTime;
      
      testResults.push({
        name: testFunction,
        success: true,
        duration: testDuration
      });
      
      console.log('✅ ' + testFunction + ' - 成功 (' + testDuration + 'ms)');
      
    } catch (error) {
      var testEndTime = new Date().getTime();
      var testDuration = testEndTime - testStartTime;
      
      testResults.push({
        name: testFunction,
        success: false,
        error: error.message,
        duration: testDuration
      });
      
      console.log('❌ ' + testFunction + ' - 失敗 (' + testDuration + 'ms)');
      console.log('   エラー: ' + error.message);
    }
  });
  
  var endTime = new Date().getTime();
  var totalDuration = endTime - startTime;
  
  // 結果サマリー
  var successCount = testResults.filter(function(result) { return result.success; }).length;
  var failureCount = testResults.filter(function(result) { return !result.success; }).length;
  var successRate = (successCount / testResults.length * 100).toFixed(1);
  
  console.log('---');
  console.log('=== FormManager TDDテスト実行結果 ===');
  console.log('総実行時間: ' + totalDuration + 'ms');
  console.log('成功: ' + successCount + '/' + testResults.length);
  console.log('失敗: ' + failureCount + '/' + testResults.length);
  console.log('成功率: ' + successRate + '%');
  console.log('---');
  
  if (failureCount > 0) {
    console.log('失敗したテスト:');
    testResults.filter(function(result) { return !result.success; }).forEach(function(result) {
      console.log('  - ' + result.name + ': ' + result.error);
    });
  }
  
  return {
    totalTests: testResults.length,
    successCount: successCount,
    failureCount: failureCount,
    successRate: successRate,
    totalDuration: totalDuration,
    results: testResults
  };
}

/**
 * システム設定の初期化確認
 * ADMIN_EMAILSなどの必要な設定が存在しない場合に追加
 */
function initializeSystemConfigIfNeeded() {
  try {
    var sheet = getSheet(getSheetName('SYSTEM_CONFIG'));
    var data = sheet.getDataRange().getValues();
    
    // ADMIN_EMAILS設定の存在確認
    var hasAdminEmails = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'ADMIN_EMAILS') {
        hasAdminEmails = true;
        break;
      }
    }
    
    if (!hasAdminEmails) {
      console.log('ADMIN_EMAILS設定を追加中...');
      var configData = [
        ['ADMIN_EMAILS', 'manager@example.com,dev-manager@example.com', '管理者メールアドレス（カンマ区切り）', 'TRUE'],
        ['EMAIL_MOCK_ENABLED', 'TRUE', 'メール送信のモック有効化', 'TRUE'],
        ['EMAIL_ACTUAL_SEND', 'FALSE', '実際のメール送信フラグ', 'TRUE']
      ];
      
      configData.forEach(function(rowData) {
        appendDataToSheet(getSheetName('SYSTEM_CONFIG'), rowData);
      });
      
      console.log('✓ システム設定を初期化しました');
    }
  } catch (error) {
    console.log('システム設定初期化エラー: ' + error.message);
  }
}

/**
 * 個別のFormManagerテストを実行
 * @param {string} testName - 実行するテスト名
 */
function runSingleFormTest(testName) {
  console.log('=== FormManager 個別テスト実行 ===');
  console.log('テスト名: ' + testName);
  console.log('実行日時: ' + new Date().toLocaleString('ja-JP'));
  
  var startTime = new Date().getTime();
  
  try {
    // テスト実行
    eval(testName + '()');
    
    var endTime = new Date().getTime();
    var duration = endTime - startTime;
    
    console.log('✅ テスト成功 (' + duration + 'ms)');
    return {
      success: true,
      duration: duration,
      error: null
    };
    
  } catch (error) {
    var endTime = new Date().getTime();
    var duration = endTime - startTime;
    
    console.log('❌ テスト失敗 (' + duration + 'ms)');
    console.log('エラー: ' + error.message);
    return {
      success: false,
      duration: duration,
      error: error.message
    };
  }
}

/**
 * FormManagerのクイックテスト実行
 * 基本的な機能のみをテスト
 */
function runFormManagerQuickTest() {
  console.log('=== FormManager クイックテスト実行 ===');
  
  var quickTests = [
    'testValidateFormData_ValidData_ReturnsTrue',
    'testValidateFormData_InvalidData_ReturnsFalse',
    'testConvertFormDataToLogRaw_ValidData_ReturnsCorrectFormat'
  ];
  
  var passedTests = 0;
  var failedTests = 0;
  
  for (var i = 0; i < quickTests.length; i++) {
    try {
      eval(quickTests[i] + '()');
      console.log('✅ ' + quickTests[i]);
      passedTests++;
    } catch (error) {
      console.log('❌ ' + quickTests[i] + ': ' + error.message);
      failedTests++;
    }
  }
  
  console.log('---');
  console.log('クイックテスト結果: ' + passedTests + '/' + (passedTests + failedTests) + ' 成功');
  
  return {
    passedTests: passedTests,
    failedTests: failedTests,
    totalTests: passedTests + failedTests
  };
}

/**
 * FormManagerのパフォーマンステスト
 */
function runFormManagerPerformanceTest() {
  console.log('=== FormManager パフォーマンステスト ===');
  
  var testData = {
    timestamp: new Date(),
    employeeId: 'PERF001',
    employeeName: 'パフォーマンス太郎',
    action: 'IN',
    ipAddress: '192.168.1.200',
    remarks: 'パフォーマンステスト'
  };
  
  var iterations = 10;
  var totalTime = 0;
  var successCount = 0;
  
  console.log('実行回数: ' + iterations);
  
  for (var i = 0; i < iterations; i++) {
    var startTime = new Date().getTime();
    
    try {
      var result = processFormResponse(testData);
      var endTime = new Date().getTime();
      var duration = endTime - startTime;
      
      totalTime += duration;
      if (result.success) {
        successCount++;
      }
      
      console.log('実行 ' + (i + 1) + ': ' + duration + 'ms (' + (result.success ? '成功' : '失敗') + ')');
      
    } catch (error) {
      var endTime = new Date().getTime();
      var duration = endTime - startTime;
      totalTime += duration;
      console.log('実行 ' + (i + 1) + ': ' + duration + 'ms (エラー)');
    }
  }
  
  var averageTime = totalTime / iterations;
  var successRate = (successCount / iterations) * 100;
  
  console.log('---');
  console.log('平均実行時間: ' + averageTime.toFixed(2) + 'ms');
  console.log('成功率: ' + successRate.toFixed(1) + '%');
  console.log('総実行時間: ' + totalTime + 'ms');
  
  return {
    iterations: iterations,
    averageTime: averageTime,
    successRate: successRate,
    totalTime: totalTime
  };
}

/**
 * FormManagerの統合テスト実行
 */
function runFormManagerIntegrationTest() {
  console.log('=== FormManager 統合テスト実行 ===');
  
  var integrationTests = [
    'testProcessFormResponse_IntegrationTest_CompleteWorkflow',
    'testProcessMultipleFormResponses_ValidData_ProcessesAll',
    'testGetFormResponseStats_ValidData_ReturnsStats'
  ];
  
  var passedTests = 0;
  var failedTests = 0;
  
  for (var i = 0; i < integrationTests.length; i++) {
    try {
      eval(integrationTests[i] + '()');
      console.log('✅ ' + integrationTests[i]);
      passedTests++;
    } catch (error) {
      console.log('❌ ' + integrationTests[i] + ': ' + error.message);
      failedTests++;
    }
  }
  
  console.log('---');
  console.log('統合テスト結果: ' + passedTests + '/' + (passedTests + failedTests) + ' 成功');
  
  return {
    passedTests: passedTests,
    failedTests: failedTests,
    totalTests: passedTests + failedTests
  };
}

/**
 * FormManagerのTDDサイクル実行
 * Red-Green-Refactorサイクルを支援
 */
function runFormManagerTDDCycle() {
  console.log('=== FormManager TDDサイクル実行 ===');
  console.log('1. Red: テスト実行（期待される失敗）');
  console.log('2. Green: 最小限の実装');
  console.log('3. Refactor: コード改善');
  console.log('---');
  
  // 現在のテスト状況を確認
  var testResult = runFormManagerTests();
  
  if (testResult.failureCount > 0) {
    console.log('🔴 Redフェーズ: ' + testResult.failureCount + '個のテストが失敗');
    console.log('次のステップ: Greenフェーズで実装を完了');
  } else {
    console.log('🟢 Greenフェーズ: 全テストが成功');
    console.log('次のステップ: Refactorフェーズでコード改善');
  }
  
  return testResult;
} 