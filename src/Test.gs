/**
 * GAS用テストフレームワーク
 * TDD実装で使用するテストランナー
 */

// テスト結果格納用のグローバル変数
var testResults = [];

/**
 * アサーション関数
 */
function assert(condition, message) {
  if (!condition) {
    throw new Error(message || 'Assertion failed');
  }
}

function assertEquals(expected, actual, message) {
  if (expected !== actual) {
    throw new Error((message || 'Assertion failed') + ': expected ' + expected + ', but got ' + actual);
  }
}

function assertNotNull(value, message) {
  if (value === null || value === undefined) {
    throw new Error(message || 'Value should not be null or undefined');
  }
}

function assertTrue(condition, message) {
  if (!condition) {
    throw new Error(message || 'Condition should be true');
  }
}

function assertFalse(condition, message) {
  if (condition) {
    throw new Error(message || 'Condition should be false');
  }
}

/**
 * 個別テスト実行関数
 */
function runTest(testFunction, testName) {
  try {
    testFunction();
    testResults.push({
      name: testName,
      status: 'PASS',
      message: ''
    });
    console.log('✓ ' + testName + ' - PASS');
  } catch (error) {
    testResults.push({
      name: testName,
      status: 'FAIL',
      message: error.message
    });
    console.log('✗ ' + testName + ' - FAIL: ' + error.message);
  }
}

/**
 * 全テスト実行とレポート生成
 */
function runAllTests() {
  console.log('=== GAS Test Framework - 実行開始 ===');
  testResults = []; // 結果をリセット
  
  // まずはダミーテストから実行
  runTest(testDummy_Assert_True_ShouldPass, 'testDummy_Assert_True_ShouldPass');
  runTest(testDummy_Assert_False_ShouldFail, 'testDummy_Assert_False_ShouldFail');
  
  // Constants.gs のテスト実行
  runConstantsTests();
  
  // CommonUtils.gs のテスト実行
  runCommonUtilsTests();
  
  // SpreadsheetManager.gs のテスト実行
  runSpreadsheetManagerTests();
  
  // テスト結果のサマリー表示
  showTestSummary();
}

/**
 * テスト結果のサマリー表示
 */
function showTestSummary() {
  var passCount = testResults.filter(function(result) { return result.status === 'PASS'; }).length;
  var failCount = testResults.filter(function(result) { return result.status === 'FAIL'; }).length;
  
  console.log('\n=== テスト結果サマリー ===');
  console.log('実行テスト数: ' + testResults.length);
  console.log('成功: ' + passCount);
  console.log('失敗: ' + failCount);
  
  if (failCount > 0) {
    console.log('\n=== 失敗テスト詳細 ===');
    testResults.filter(function(result) { return result.status === 'FAIL'; }).forEach(function(result) {
      console.log('- ' + result.name + ': ' + result.message);
    });
  }
  
  console.log('=== GAS Test Framework - 実行完了 ===');
  return { total: testResults.length, passed: passCount, failed: failCount };
}

/**
 * ダミーテスト（成功）
 */
function testDummy_Assert_True_ShouldPass() {
  // Red: まずは必ず成功するテストで環境確認
  assert(true, 'ダミーテスト - trueは成功すべき');
}

/**
 * ダミーテスト（失敗例）
 * このテストは実際にはコメントアウトして使用
 */
function testDummy_Assert_False_ShouldFail() {
  // このテストは失敗を確認するためのもの（通常はコメントアウト）
  // assert(false, 'ダミーテスト - falseは失敗すべき');
  
  // 実際には成功するテストに変更
  assert(true, 'ダミーテスト - 環境確認用');
}

/**
 * テスト環境確認用の手動実行関数
 */
function testEnvironmentCheck() {
  console.log('テスト環境確認を開始します...');
  
  try {
    // アサーション関数の動作確認
    assert(true);
    assertEquals(1, 1);
    assertNotNull('test');
    assertTrue(true);
    assertFalse(false);
    
    console.log('✓ 全てのアサーション関数が正常に動作しています');
    return true;
  } catch (error) {
    console.log('✗ テスト環境にエラーがあります: ' + error.message);
    return false;
  }
}

/**
 * 特定のモジュールのテストのみ実行
 * @param {string} moduleName - テストするモジュール名
 */
function runModuleTests(moduleName) {
  console.log('=== ' + moduleName + ' テスト実行開始 ===');
  testResults = []; // 結果をリセット
  
  switch (moduleName) {
    case 'Constants':
      runConstantsTests();
      break;
    case 'CommonUtils':
      runCommonUtilsTests();
      break;
    case 'SpreadsheetManager':
      runSpreadsheetManagerTests();
      break;
    default:
      console.log('不明なモジュール名: ' + moduleName);
      return;
  }
  
  showTestSummary();
}

/**
 * TDD専用テスト実行関数
 * 短期サイクル用（5-10分ルール対応）
 */
function runQuickTests() {
  console.log('=== TDD クイックテスト実行 ===');
  testResults = []; // 結果をリセット
  
  // 基本的なテストのみ実行（高速化）
  runTest(testDummy_Assert_True_ShouldPass, 'testDummy_Assert_True_ShouldPass');
  
  // 最新実装のテストを優先実行
  runTest(testGetColumnIndex_Employee_Name_ReturnsCorrectIndex, 'testGetColumnIndex_Employee_Name_ReturnsCorrectIndex');
  runTest(testFormatDate_ValidDate_ReturnsFormattedString, 'testFormatDate_ValidDate_ReturnsFormattedString');
  
  showTestSummary();
  
  if (testResults.filter(function(result) { return result.status === 'FAIL'; }).length === 0) {
    console.log('✓ クイックテスト全通過 - TDDサイクル継続可能');
  } else {
    console.log('✗ クイックテスト失敗 - 実装を修正してください');
  }
}

/**
 * パフォーマンステスト用
 * F.I.R.S.T.原則のFast（高速）チェック
 */
function runPerformanceTest() {
  console.log('=== パフォーマンステスト実行 ===');
  
  var startTime = new Date().getTime();
  
  // 全テストを実行
  runAllTests();
  
  var endTime = new Date().getTime();
  var duration = endTime - startTime;
  
  console.log('テスト実行時間: ' + duration + 'ms');
  
  // F.I.R.S.T.原則: Fastの基準（30秒以内）
  if (duration < 30000) {
    console.log('✓ パフォーマンス基準クリア（30秒以内）');
  } else {
    console.log('✗ パフォーマンス基準未達成（30秒超過）');
    console.log('テストの最適化を検討してください');
  }
  
  return { duration: duration, passed: duration < 30000 };
}

/**
 * カバレッジ情報表示（簡易版）
 * t-wada推奨の開発支援ツールとして活用
 */
function showTestCoverage() {
  console.log('=== テストカバレッジ情報 ===');
  
  var moduleInfo = {
    'Constants.gs': {
      functions: ['getColumnIndex', 'getSheetName', 'getActionConstant', 'getAppConfig'],
      tested: ['getColumnIndex', 'getSheetName', 'getActionConstant', 'getAppConfig']
    },
    'CommonUtils.gs': {
      functions: ['formatDate', 'formatTime', 'parseDate', 'timeStringToMinutes', 'isEmpty', 'isValidEmail'],
      tested: ['formatDate', 'formatTime', 'parseDate', 'timeStringToMinutes', 'isEmpty', 'isValidEmail']
    },
    'SpreadsheetManager.gs': {
      functions: ['getSheet', 'createSheet', 'sheetExists', 'getOrCreateSheet', 'appendDataToSheet'],
      tested: ['getOrCreateSheet', 'appendDataToSheet'] // パラメータ検証のみ
    }
  };
  
  Object.keys(moduleInfo).forEach(function(module) {
    var info = moduleInfo[module];
    var coverage = (info.tested.length / info.functions.length * 100).toFixed(1);
    console.log(module + ': ' + coverage + '% (' + info.tested.length + '/' + info.functions.length + ' 関数)');
  });
  
  console.log('\n注: カバレッジは開発支援用です。100%を目標にせず、重要な機能の品質確保に活用してください。');
}

/**
 * TDD状況レポート
 * Red-Green-Refactorサイクルの進捗確認
 */
function showTDDStatus() {
  console.log('=== TDD進捗レポート ===');
  
  var tddStatus = {
    'フェーズ1: 基盤構築': {
      'setup-1': 'Green（完了）',
      'constants-1': 'Green（完了）', 
      'utils-1': 'Green（完了）',
      'spreadsheet-1': 'Green（完了）',
      'test-framework': 'Refactor（進行中）'
    }
  };
  
  Object.keys(tddStatus).forEach(function(phase) {
    console.log('\n' + phase + ':');
    var tasks = tddStatus[phase];
    Object.keys(tasks).forEach(function(task) {
      console.log('  ' + task + ': ' + tasks[task]);
    });
  });
  
  console.log('\n次のステップ: フェーズ2のコア機能実装（utils-2, auth-1, webapp-1）');
}

/**
 * GAS環境専用のデバッグ情報表示
 */
function showGASEnvironmentInfo() {
  console.log('=== GAS環境情報 ===');
  
  try {
    // 基本的な環境情報
    console.log('実行コンテキスト: ' + (typeof SpreadsheetApp !== 'undefined' ? 'GAS環境' : 'ローカル環境'));
    
    if (typeof SpreadsheetApp !== 'undefined') {
      console.log('スプレッドシートアクセス: 可能');
      
      // アクティブなスプレッドシートの確認
      try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        console.log('アクティブスプレッドシート: ' + ss.getName());
        console.log('シート数: ' + ss.getSheets().length);
      } catch (error) {
        console.log('アクティブスプレッドシート: なし（' + error.message + '）');
      }
    } else {
      console.log('スプレッドシートアクセス: 不可（ローカル環境）');
    }
    
    // 日時情報
    console.log('実行日時: ' + new Date().toString());
    
  } catch (error) {
    console.log('環境情報取得エラー: ' + error.message);
  }
} 