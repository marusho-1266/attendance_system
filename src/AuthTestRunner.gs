/**
 * 認証機能専用テストランナー
 * TDD実装フェーズ auth-1 用の高速テスト実行
 */

/**
 * 認証機能の全テストを実行
 */
function runAuthenticationTests() {
  console.log('=== 認証機能テスト開始 ===');
  var startTime = new Date().getTime();
  
  var testResults = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: [],
    duration: 0
  };
  
  // テスト関数リスト
  var authTestFunctions = [
    'testAuthenticateUser_ValidUser_ReturnsTrue',
    'testAuthenticateUser_InvalidUser_ReturnsFalse',
    'testAuthenticateUser_EmptyEmail_ReturnsFalse',
    'testAuthenticateUser_NullEmail_ReturnsFalse',
    'testAuthenticateUser_InvalidEmailFormat_ReturnsFalse',
    'testCheckPermission_ValidUser_ClockIn_ReturnsTrue',
    'testCheckPermission_InvalidUser_ClockIn_ReturnsFalse',
    'testCheckPermission_Manager_AdminAction_ReturnsTrue',
    'testCheckPermission_RegularUser_AdminAction_ReturnsFalse',
    'testCheckPermission_ValidUser_InvalidAction_ReturnsFalse',
    'testGetSessionInfo_ValidSession_ReturnsUserInfo',
    'testGetSessionInfo_AuthenticatedUser_ReturnsEmployeeData',
    'testGetSessionInfo_UnauthenticatedUser_ReturnsLimitedInfo',
    'testAuthenticationFlow_ValidUser_FullFlow'
  ];
  
  console.log('実行予定テスト数: ' + authTestFunctions.length);
  
  // 各テスト実行
  authTestFunctions.forEach(function(testFunctionName) {
    testResults.total++;
    
    try {
      console.log('実行中: ' + testFunctionName);
      
      // テスト関数を実行
      var testFunction = eval(testFunctionName);
      testFunction();
      
      testResults.passed++;
      console.log('✓ PASS: ' + testFunctionName);
      
    } catch (error) {
      testResults.failed++;
      testResults.errors.push({
        test: testFunctionName,
        error: error.message,
        stack: error.stack
      });
      console.log('✗ FAIL: ' + testFunctionName + ' - ' + error.message);
    }
  });
  
  // 実行時間計算
  var endTime = new Date().getTime();
  testResults.duration = endTime - startTime;
  
  // 結果サマリー
  console.log('\n=== 認証機能テスト結果 ===');
  console.log('総テスト数: ' + testResults.total);
  console.log('成功: ' + testResults.passed);
  console.log('失敗: ' + testResults.failed);
  console.log('実行時間: ' + testResults.duration + 'ms');
  console.log('成功率: ' + Math.round((testResults.passed / testResults.total) * 100) + '%');
  
  // エラー詳細
  if (testResults.failed > 0) {
    console.log('\n=== エラー詳細 ===');
    testResults.errors.forEach(function(error, index) {
      console.log((index + 1) + '. ' + error.test);
      console.log('   エラー: ' + error.error);
    });
  }
  
  // TDD進捗レポート
  if (testResults.failed === 0) {
    console.log('\n🎉 全テスト通過！Greenフェーズ完了、Refactorフェーズに進むことができます。');
  } else {
    console.log('\n⚠️ 失敗したテストがあります。実装を修正してください。');
  }
  
  return testResults;
}

/**
 * 特定の認証テストのみを実行（デバッグ用）
 */
function runSingleAuthTest(testFunctionName) {
  console.log('=== 単体認証テスト: ' + testFunctionName + ' ===');
  
  try {
    var testFunction = eval(testFunctionName);
    testFunction();
    console.log('✓ テスト成功: ' + testFunctionName);
    return true;
  } catch (error) {
    console.log('✗ テスト失敗: ' + testFunctionName);
    console.log('エラー: ' + error.message);
    console.log('スタック: ' + error.stack);
    return false;
  }
}

/**
 * 認証機能のクイックテスト（最重要な機能のみ）
 */
function quickAuthTest() {
  console.log('=== 認証機能クイックテスト ===');
  
  var quickTests = [
    'testAuthenticateUser_ValidUser_ReturnsTrue',
    'testCheckPermission_ValidUser_ClockIn_ReturnsTrue',
    'testGetSessionInfo_ValidSession_ReturnsUserInfo'
  ];
  
  var results = { passed: 0, failed: 0 };
  
  quickTests.forEach(function(testName) {
    try {
      var testFunction = eval(testName);
      testFunction();
      results.passed++;
      console.log('✓ ' + testName);
    } catch (error) {
      results.failed++;
      console.log('✗ ' + testName + ': ' + error.message);
    }
  });
  
  console.log('クイックテスト結果: ' + results.passed + '/' + quickTests.length + ' 通過');
  
  return results.failed === 0;
} 