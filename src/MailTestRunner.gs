/**
 * MailManager TDDテスト実行
 * 全テストケースを実行し、結果をレポート
 */
function runMailManagerTests() {
  // テストケースの定義
  var testFunctions = [
    testSendUnfinishedClockOutEmail_ValidData_SendsEmail,
    testSendUnfinishedClockOutEmail_NoUnfinished_ReturnsSuccess,
    testSendMonthlyReportEmail_ValidData_SendsEmail,
    testGenerateUnfinishedClockOutEmailTemplate_ValidData_ReturnsCorrectTemplate,
    testGenerateMonthlyReportEmailTemplate_ValidData_ReturnsCorrectTemplate,
    testCheckEmailQuota_WithinLimit_ReturnsTrue,
    testCheckEmailQuota_ExceededLimit_ReturnsFalse,
    testCheckEmailQuota_BoundaryValues_ReturnsCorrectResults,
    testSendEmail_ErrorOccurs_ReturnsError,
    testGetEmailStats_ValidData_ReturnsStats,
    testSendUnfinishedClockOutEmail_IntegrationTest_CompleteWorkflow,
    testSendMultipleEmails_ValidData_SendsAllEmails,
    testSendEmail_ExceptionHandling_ReturnsError
  ];
  
  // 動的にテスト数を計算
  var testCount = testFunctions.length;
  
  console.log('=== MailManager TDDテスト実行開始 ===');
  console.log('実行日時: ' + new Date().toLocaleString('ja-JP'));
  console.log('実行テスト数: ' + testCount);
  console.log('---');
  
  // 統一されたテストランナーを使用
  var result = runAllTests(testFunctions, 'MailManager TDDテスト');
  
  console.log('---');
  console.log('=== MailManager TDDテスト実行結果 ===');
  console.log('総実行時間: ' + result.totalDuration + 'ms');
  console.log('成功: ' + result.passedTests + '/' + result.totalTests);
  console.log('失敗: ' + result.failedTests + '/' + result.totalTests);
  console.log('成功率: ' + result.successRate + '%');
  console.log('---');
  
  return {
    totalTests: result.totalTests,
    successCount: result.passedTests,
    failureCount: result.failedTests,
    successRate: result.successRate,
    totalDuration: result.totalDuration,
    results: result.results
  };
}

/**
 * メール機能のクイックテスト実行
 */
function runMailManagerQuickTests() {
  // 基本的な機能テスト
  var quickTests = [
    testGenerateUnfinishedClockOutEmailTemplate_ValidData_ReturnsCorrectTemplate,
    testGenerateMonthlyReportEmailTemplate_ValidData_ReturnsCorrectTemplate,
    testCheckEmailQuota_WithinLimit_ReturnsTrue,
    testCheckEmailQuota_ExceededLimit_ReturnsFalse,
    testCheckEmailQuota_BoundaryValues_ReturnsCorrectResults
  ];
  
  // 統一されたテストランナーを使用
  return runQuickTests(quickTests, 'MailManager.gs クイックテスト');
}

/**
 * メール機能のパフォーマンステスト実行
 */
function runMailManagerPerformanceTests() {
  // パフォーマンステスト
  var performanceTests = [
    testGenerateUnfinishedClockOutEmailTemplate_ValidData_ReturnsCorrectTemplate,
    testGenerateMonthlyReportEmailTemplate_ValidData_ReturnsCorrectTemplate,
    testCheckEmailQuota_WithinLimit_ReturnsTrue
  ];
  
  // 統一されたテストランナーを使用
  return runPerformanceTests(performanceTests, 'MailManager.gs パフォーマンステスト', 10);
}

/**
 * メール機能の統合テスト実行
 */
function runMailManagerIntegrationTests() {
  // 統合テスト
  var integrationTests = [
    testSendUnfinishedClockOutEmail_IntegrationTest_CompleteWorkflow,
    testSendMultipleEmails_ValidData_SendsAllEmails
  ];
  
  // 統一されたテストランナーを使用
  return runIntegrationTests(integrationTests, 'MailManager.gs 統合テスト');
}

/**
 * メール機能の全テスト実行（メイン関数）
 */
function runMailManagerAllTests() {
  console.log('お知らせ	実行開始');
  
  // クイックテスト実行
  var quickResult = runMailManagerQuickTests();
  
  // 統合テスト実行
  var integrationResult = runMailManagerIntegrationTests();
  
  // パフォーマンステスト実行
  var performanceResult = runMailManagerPerformanceTests();
  
  // 全テスト実行
  var fullResult = runMailManagerTests();
  
  console.log('お知らせ	実行完了');
  
  return {
    quickResult: quickResult,
    integrationResult: integrationResult,
    performanceResult: performanceResult,
    fullResult: fullResult
  };
} 