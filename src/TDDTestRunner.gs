/**
 * フェーズ2 TDDサイクル専用テストランナー
 * Red-Green-Refactorの高速サイクルを支援
 */

/**
 * Phase2 業務ロジックの基本テスト実行
 * TDDサイクル用の高速テスト
 */
function runPhase2BasicTests() {
  console.log('=== Phase2 業務ロジック 基本テスト実行 ===');
  testResults = []; // 結果をリセット
  
  // isHoliday関数の基本テスト
  runTest(testIsHoliday_Weekday_ReturnsFalse, 'testIsHoliday_Weekday_ReturnsFalse');
  runTest(testIsHoliday_Saturday_ReturnsTrue, 'testIsHoliday_Saturday_ReturnsTrue');
  runTest(testIsHoliday_Sunday_ReturnsTrue, 'testIsHoliday_Sunday_ReturnsTrue');
  
  // calcWorkTime関数の基本テスト  
  runTest(testCalcWorkTime_StandardWork_ReturnsCorrectTime, 'testCalcWorkTime_StandardWork_ReturnsCorrectTime');
  runTest(testCalcWorkTime_Overtime_ReturnsCorrectTime, 'testCalcWorkTime_Overtime_ReturnsCorrectTime');
  
  // getEmployee関数の基本テスト
  runTest(testGetEmployee_ValidEmail_ReturnsEmployeeInfo, 'testGetEmployee_ValidEmail_ReturnsEmployeeInfo');
  runTest(testGetEmployee_InvalidEmail_ReturnsNull, 'testGetEmployee_InvalidEmail_ReturnsNull');
  
  showTestSummary();
  
  var failCount = testResults.filter(function(result) { return result.status === 'FAIL'; }).length;
  if (failCount === 0) {
    console.log('✅ Phase2基本テスト全通過 - Greenステップ完了');
    console.log('次: Refactorステップに進んでください');
  } else {
    console.log('❌ ' + failCount + '件のテスト失敗 - 実装を修正してください');
  }
  
  return failCount === 0;
}

/**
 * Phase2 業務ロジックの例外テスト実行
 * エラーハンドリングの確認用
 */
function runPhase2ExceptionTests() {
  console.log('=== Phase2 業務ロジック 例外テスト実行 ===');
  testResults = []; // 結果をリセット
  
  // isHoliday関数の例外テスト
  runTest(testIsHoliday_NullDate_ThrowsError, 'testIsHoliday_NullDate_ThrowsError');
  
  // calcWorkTime関数の例外テスト
  runTest(testCalcWorkTime_InvalidTimeFormat_ThrowsError, 'testCalcWorkTime_InvalidTimeFormat_ThrowsError');
  runTest(testCalcWorkTime_NegativeBreak_ThrowsError, 'testCalcWorkTime_NegativeBreak_ThrowsError');
  
  // getEmployee関数の例外テスト
  runTest(testGetEmployee_EmptyEmail_ThrowsError, 'testGetEmployee_EmptyEmail_ThrowsError');
  runTest(testGetEmployee_NullEmail_ThrowsError, 'testGetEmployee_NullEmail_ThrowsError');
  runTest(testGetEmployee_InvalidEmailFormat_ThrowsError, 'testGetEmployee_InvalidEmailFormat_ThrowsError');
  
  showTestSummary();
  
  var failCount = testResults.filter(function(result) { return result.status === 'FAIL'; }).length;
  if (failCount === 0) {
    console.log('✅ Phase2例外テスト全通過 - エラーハンドリング正常');
  } else {
    console.log('❌ ' + failCount + '件のテスト失敗 - エラーハンドリングを修正してください');
  }
  
  return failCount === 0;
}

/**
 * Phase2 業務ロジックの全テスト実行
 * 完成版の総合テスト
 */
function runPhase2AllTests() {
  console.log('=== Phase2 業務ロジック 全テスト実行 ===');
  
  var basicPassed = runPhase2BasicTests();
  console.log('');
  var exceptionPassed = runPhase2ExceptionTests();
  
  console.log('\n=== Phase2 総合結果 ===');
  if (basicPassed && exceptionPassed) {
    console.log('🎉 Phase2業務ロジック実装完了！');
    console.log('次: Phase2-auth-1（認証機能）の実装に進んでください');
  } else {
    if (!basicPassed) console.log('⚠️  基本機能に問題があります');
    if (!exceptionPassed) console.log('⚠️  例外処理に問題があります');
  }
  
  return basicPassed && exceptionPassed;
}

/**
 * TDDサイクル確認用の超簡易テスト
 * 5分以内でのRed-Green確認
 */
function quickTDDCheck() {
  console.log('=== TDDサイクル確認（5分以内） ===');
  testResults = []; // 結果をリセット
  
  // 各関数1つずつのクイックテスト
  runTest(testIsHoliday_Weekday_ReturnsFalse, 'isHoliday基本動作');
  runTest(testCalcWorkTime_StandardWork_ReturnsCorrectTime, 'calcWorkTime基本動作');
  runTest(testGetEmployee_ValidEmail_ReturnsEmployeeInfo, 'getEmployee基本動作');
  
  showTestSummary();
  
  var failCount = testResults.filter(function(result) { return result.status === 'FAIL'; }).length;
  if (failCount === 0) {
    console.log('⚡ クイックチェック通過 - TDDサイクル継続可能');
  } else {
    console.log('🔴 ' + failCount + '件失敗 - 基本実装を確認してください');
  }
  
  return failCount === 0;
} 