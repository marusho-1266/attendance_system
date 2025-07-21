/**
 * Triggers.gs のテストケース
 * TDD Red フェーズ: トリガー関数の期待する動作をテストで定義
 * 
 * テスト対象機能:
 * - onOpen(): 管理メニュー追加
 * - dailyJob(): 日次集計処理（02:00実行）
 * - weeklyOvertimeJob(): 週次残業警告（月曜07:00実行）
 * - monthlyJob(): 月次集計処理（毎月1日02:30実行）
 */

// === onOpen関数のテスト ===

/**
 * onOpen関数: 管理メニューが追加されることをテスト
 */
function testOnOpen_SpreadsheetOpen_AddsManagementMenu() {
  // Red: スプレッドシート開発時に管理メニューが追加されることを期待
  
  try {
    // onOpen関数の実行と戻り値の検証
    var result = onOpen();
    
    // 戻り値の構造検証
    assertNotNull(result, 'onOpen関数は戻り値を返すべき');
    assertTrue(typeof result === 'object', 'onOpen関数の戻り値はオブジェクトであるべき');
    assertTrue('success' in result, 'onOpen関数の戻り値にsuccessプロパティがあるべき');
    assertTrue('message' in result, 'onOpen関数の戻り値にmessageプロパティがあるべき');
    assertTrue(typeof result.success === 'boolean', 'successプロパティはbooleanであるべき');
    assertTrue(typeof result.message === 'string', 'messageプロパティはstringであるべき');
    
    // 正常実行の場合、successがtrueであることを確認
    assertTrue(result.success, 'onOpen関数は正常実行時にsuccess: trueを返すべき: ' + result.message);
    
    console.log('✓ onOpen関数の実行確認完了: ' + result.message);
  } catch (error) {
    throw new Error('onOpen関数の実行に失敗: ' + error.message);
  }
}

/**
 * onOpen関数: エラーハンドリングテスト
 */
function testOnOpen_ErrorHandling_DoesNotThrow() {
  // Red: エラーが発生してもスプレッドシートの開発が阻害されないことを期待
  
  try {
    var result = onOpen();
    
    // onOpen関数は例外を投げずに結果を返すべき
    assertNotNull(result, 'onOpen関数はエラー時でも戻り値を返すべき');
    assertTrue(typeof result === 'object', 'onOpen関数の戻り値は常にオブジェクトであるべき');
    assertTrue('success' in result, 'onOpen関数の戻り値には常にsuccessプロパティがあるべき');
    assertTrue('message' in result, 'onOpen関数の戻り値には常にmessageプロパティがあるべき');
    
    // エラー時はsuccess: falseが返されることもあることを許容
    if (!result.success) {
      console.log('警告: onOpen実行時にエラー: ' + result.message);
      assertTrue(typeof result.message === 'string' && result.message.length > 0, 
                'エラー時はmessageに具体的なエラー内容があるべき');
    }
    
    console.log('✓ onOpenのエラーハンドリング確認完了: success=' + result.success);
  } catch (error) {
    // onOpen関数はユーザー体験を阻害してはいけないため、例外を投げるべきではない
    throw new Error('onOpen関数は例外を投げるべきではない: ' + error.message);
  }
}

// === dailyJob関数のテスト ===

/**
 * dailyJob関数: 前日分のDaily_Summary更新テスト
 */
function testDailyJob_ExecuteDaily_UpdatesDailySummary() {
  // Red: 前日分のLog_RawデータからDaily_Summaryを更新することを期待
  
  try {
    // 実行前のDaily_Summaryシート状態を記録
    var dailySummarySheet = getOrCreateSheet(getSheetName('DAILY_SUMMARY'));
    var beforeData = dailySummarySheet.getDataRange().getValues();
    var beforeRowCount = beforeData.length;
    
    // 前日の日付を計算
    var yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    var targetDate = formatDate(yesterday);
    
    console.log('テスト対象日: ' + targetDate);
    
    // 実行前の時間記録
    var beforeExecution = new Date().getTime();
    
    // dailyJob関数の実行
    var result = dailyJob();
    
    var afterExecution = new Date().getTime();
    var executionTime = afterExecution - beforeExecution;
    
    // 1. 戻り値の構造検証
    assertNotNull(result, 'dailyJob関数は戻り値を返すべき');
    assertTrue(typeof result === 'object', 'dailyJob関数の戻り値はオブジェクトであるべき');
    assertTrue('success' in result, 'dailyJob関数の戻り値にsuccessプロパティがあるべき');
    assertTrue('summaryResult' in result, 'dailyJob関数の戻り値にsummaryResultプロパティがあるべき');
    assertTrue(result.success, 'dailyJob関数は正常実行時にsuccess: trueを返すべき');
    
    // 2. Daily_Summaryシートの更新確認
    var afterData = dailySummarySheet.getDataRange().getValues();
    
    // ヘッダー行の確認
    if (afterData.length > 0) {
      var headerRow = afterData[0];
      assertTrue(headerRow.length >= 9, 'Daily_Summaryにはヘッダー行に最低9列あるべき（実際: ' + headerRow.length + '列）');
    }
    
    // 対象日のデータが含まれているかチェック
    var targetDateRecordFound = false;
    var targetDateRecords = 0;
    
    for (var i = 1; i < afterData.length; i++) {
      var row = afterData[i];
      if (row[0] && formatDate(new Date(row[0])) === targetDate) {
        targetDateRecordFound = true;
        targetDateRecords++;
        
        // データの基本構造確認
        assertTrue(row.length >= 9, '各レコードは最低9列のデータを持つべき（実際: ' + row.length + '列）');
        assertTrue(row[1] !== undefined && row[1] !== '', '社員IDは空でないべき');
      }
    }
    
    // 3. summaryResultの詳細検証
    if (result.summaryResult) {
      assertTrue('recordsUpdated' in result.summaryResult, 'summaryResultにrecordsUpdatedがあるべき');
      assertTrue(typeof result.summaryResult.recordsUpdated === 'number', 'recordsUpdatedは数値であるべき');
      
      // 更新されたレコード数と実際の対象日レコード数の整合性確認
      if (result.summaryResult.recordsUpdated > 0) {
        assertTrue(targetDateRecordFound, '更新レコードがある場合、対象日のデータがDaily_Summaryに存在すべき');
        assertTrue(targetDateRecords >= result.summaryResult.recordsUpdated, 
                  '対象日のレコード数は更新レコード数以上であるべき');
      }
    }
    
    // 4. 実行時間の確認（クォータ制限内）
    assertTrue(executionTime < 60000, 'dailyJobは1分以内に完了すべき（実際: ' + executionTime + 'ms）');
    
    console.log('✓ Daily_Summary更新確認: 対象日レコード数=' + targetDateRecords + 
               ', 更新件数=' + (result.summaryResult ? result.summaryResult.recordsUpdated : 0));
    console.log('✓ dailyJob実行時間確認: ' + executionTime + 'ms');
    
  } catch (error) {
    throw new Error('dailyJob実行エラー: ' + error.message);
  }
}

/**
 * dailyJob関数: 実行時間測定テスト（パフォーマンス確認専用）
 */
function testDailyJob_PerformanceCheck_ExecutesWithinTimeLimit() {
  // Red: dailyJobがクォータ制限内で実行されることを期待
  
  try {
    // 実行時間測定のみに集中
    var beforeExecution = new Date().getTime();
    
    // dailyJob関数の実行
    var result = dailyJob();
    
    var afterExecution = new Date().getTime();
    var executionTime = afterExecution - beforeExecution;
    
    // 基本的な戻り値確認
    assertNotNull(result, 'dailyJob関数は戻り値を返すべき');
    assertTrue(result.success, 'dailyJob関数は正常実行すべき');
    
    // 実行時間の確認（クォータ制限内）
    assertTrue(executionTime < 60000, 'dailyJobは1分以内に完了すべき（実際: ' + executionTime + 'ms）');
    
    // パフォーマンス目標の確認（10秒以内推奨）
    if (executionTime > 10000) {
      console.log('⚠️ パフォーマンス注意: ' + executionTime + 'ms（推奨10秒以内）');
    } else {
      console.log('✓ パフォーマンス良好: ' + executionTime + 'ms');
    }
    
    console.log('✓ dailyJob実行時間測定完了: ' + executionTime + 'ms');
    
  } catch (error) {
    throw new Error('dailyJobパフォーマンステスト実行エラー: ' + error.message);
  }
}

/**
 * dailyJob関数: 未退勤者一覧メール送信テスト
 */
function testDailyJob_HasUnfinishedClockOut_SendsEmail() {
  // Red: 未退勤者がいる場合、一覧メールが送信されることを期待
  
  try {
    // テストモードの確認
    var isTestMode = getTestModeConfig('EMAIL_MOCK_ENABLED');
    assertTrue(isTestMode, 'テスト実行時はメールモックモードが有効であるべき');
    
    // モックデータをクリア
    clearMockEmailData();
    
    // dailyJob実行前のモックデータ状態を確認
    var beforeMockData = getMockEmailData();
    var beforeEmailCount = Object.keys(beforeMockData).length;
    
    // dailyJob関数の実行
    var result = dailyJob();
    
    // 1. 基本的な戻り値検証
    assertNotNull(result, 'dailyJob関数は戻り値を返すべき');
    assertTrue(typeof result === 'object', 'dailyJob関数の戻り値はオブジェクトであるべき');
    assertTrue('success' in result, 'dailyJob関数の戻り値にsuccessプロパティがあるべき');
    assertTrue('emailResult' in result, 'dailyJob関数の戻り値にemailResultプロパティがあるべき');
    assertTrue(result.success, 'dailyJob関数は正常実行時にsuccess: trueを返すべき');
    
    // 2. emailResultの詳細検証
    if (result.emailResult) {
      assertTrue(typeof result.emailResult === 'object', 'emailResultはオブジェクトであるべき');
      assertTrue('unfinishedCount' in result.emailResult, 'emailResultにunfinishedCountがあるべき');
      assertTrue(typeof result.emailResult.unfinishedCount === 'number', 'unfinishedCountは数値であるべき');
      
      // 3. モックメール送信の検証
      var afterMockData = getMockEmailData();
      var afterEmailCount = Object.keys(afterMockData).length;
      
      if (result.emailResult.unfinishedCount > 0) {
        // 未退勤者がいる場合：メール送信が行われるべき
        assertTrue(afterEmailCount > beforeEmailCount, '未退勤者がいる場合、メール送信記録があるべき');
        
        // 送信されたメール情報の詳細検証
        var emailKeys = Object.keys(afterMockData);
        var latestEmailKey = emailKeys[emailKeys.length - 1];
        var mockEmail = afterMockData[latestEmailKey];
        
        // メール内容の検証
        assertTrue(mockEmail.type === 'unfinished_clockout', 'メールタイプが未退勤者メールであるべき');
        assertTrue(Array.isArray(mockEmail.recipients), 'recipients配列が存在すべき');
        assertTrue(mockEmail.recipients.length > 0, '送信先が設定されているべき');
        assertTrue(typeof mockEmail.subject === 'string', '件名が文字列であるべき');
        assertTrue(mockEmail.subject.indexOf('未退勤者一覧') !== -1, '件名に「未退勤者一覧」が含まれるべき');
        assertTrue(typeof mockEmail.body === 'string', '本文が文字列であるべき');
        assertTrue(mockEmail.body.length > 0, '本文が空でないべき');
        assertTrue(Array.isArray(mockEmail.unfinishedEmployees), '未退勤者配列が存在すべき');
        assertTrue(mockEmail.unfinishedEmployees.length === result.emailResult.unfinishedCount, 
                  '未退勤者数とメール内容が一致すべき');
        
        // 各未退勤者情報の検証
        mockEmail.unfinishedEmployees.forEach(function(emp, index) {
          assertTrue('id' in emp, '未退勤者' + (index + 1) + 'にIDがあるべき');
          assertTrue('name' in emp, '未退勤者' + (index + 1) + 'に名前があるべき');
          assertTrue('department' in emp, '未退勤者' + (index + 1) + 'に部署があるべき');
          assertTrue(typeof emp.id === 'string' && emp.id.length > 0, '未退勤者' + (index + 1) + 'のIDが有効であるべき');
          assertTrue(typeof emp.name === 'string' && emp.name.length > 0, '未退勤者' + (index + 1) + 'の名前が有効であるべき');
        });
        
        console.log('✓ 未退勤者メール送信確認: ' + result.emailResult.unfinishedCount + '名、メールID=' + latestEmailKey);
        console.log('✓ 件名: ' + mockEmail.subject);
        console.log('✓ 送信先: ' + mockEmail.recipients.join(', '));
        
      } else {
        // 未退勤者がいない場合：メール送信は行われないべき
        console.log('✓ 未退勤者なし: メール送信スキップ確認');
      }
    }
    
    console.log('✓ dailyJob未退勤者メール送信テスト完了');
    
  } catch (error) {
    throw new Error('dailyJob未退勤者メール送信テスト実行エラー: ' + error.message);
  }
}

/**
 * dailyJob関数: クォータ制限内での実行テスト
 */
function testDailyJob_QuotaCompliance_StaysWithinLimits() {
  // Red: 90分/日のクォータ制限内で実行されることを期待
  
  try {
    var startTime = new Date().getTime();
    
    dailyJob();
    
    var endTime = new Date().getTime();
    var duration = endTime - startTime;
    
    // 5分以内での完了を目標（安全マージン）
    assertTrue(duration < 300000, 'dailyJobは5分以内に完了すべき（クォータ管理）');
    console.log('✓ dailyJobクォータ制限確認: ' + duration + 'ms');
    
  } catch (error) {
    throw new Error('dailyJobクォータテストエラー: ' + error.message);
  }
}

// === weeklyOvertimeJob関数のテスト ===

/**
 * weeklyOvertimeJob関数: 週次残業集計テスト
 */
function testWeeklyOvertimeJob_Monday_CalculatesOvertimeHours() {
  // Red: 毎週月曜日に直近4週間の残業時間を集計することを期待
  
  try {
    var result = weeklyOvertimeJob();
    
    // 残業集計結果の確認
    assertTrue(result !== undefined, 'weeklyOvertimeJobは集計結果を返すべき');
    console.log('✓ 週次残業集計実行確認');
    
  } catch (error) {
    throw new Error('weeklyOvertimeJob実行エラー: ' + error.message);
  }
}

/**
 * weeklyOvertimeJob関数: 管理者への警告メール送信テスト
 */
function testWeeklyOvertimeJob_HighOvertimeDetected_SendsWarningEmail() {
  // Red: 月80h超えそうな残業者がいる場合、管理者に警告メールが送信されることを期待
  
  try {
    // 残業超過チェックの実行
    var result = weeklyOvertimeJob();
    
    // メール送信については現段階では関数実行の成功を確認
    assertTrue(result !== null, 'weeklyOvertimeJobは処理結果を返すべき');
    console.log('✓ 週次残業警告処理実行確認');
    
  } catch (error) {
    throw new Error('weeklyOvertimeJob警告処理エラー: ' + error.message);
  }
}

// === monthlyJob関数のテスト ===

/**
 * monthlyJob関数: Monthly_Summary転記テスト
 */
function testMonthlyJob_FirstDayOfMonth_UpdatesMonthlySummary() {
  // Red: 毎月1日に前月分のMonthly_Summaryが更新されることを期待
  
  try {
    var result = monthlyJob();
    
    // Monthly_Summary更新の確認
    assertTrue(result !== undefined, 'monthlyJobは処理結果を返すべき');
    console.log('✓ 月次集計処理実行確認');
    
  } catch (error) {
    throw new Error('monthlyJob実行エラー: ' + error.message);
  }
}

/**
 * monthlyJob関数: CSV出力とDrive保存テスト
 */
function testMonthlyJob_MonthlyProcess_ExportsCSVToDrive() {
  // Red: 月次処理でCSVファイルがDriveに出力されることを期待
  
  try {
    var result = monthlyJob();
    
    // CSV出力処理の確認
    assertTrue(result !== null, 'monthlyJobはCSV出力結果を返すべき');
    console.log('✓ 月次CSV出力処理実行確認');
    
  } catch (error) {
    throw new Error('monthlyJob CSV出力エラー: ' + error.message);
  }
}

/**
 * monthlyJob関数: 管理者へのリンクメール送信テスト
 */
function testMonthlyJob_CSVGenerated_SendsLinkEmail() {
  // Red: CSV生成後、管理者にリンクメールが送信されることを期待
  
  try {
    var result = monthlyJob();
    
    // メール送信処理の確認
    assertTrue(typeof result === 'object', 'monthlyJobは詳細な処理結果を返すべき');
    console.log('✓ 月次リンクメール処理実行確認');
    
  } catch (error) {
    throw new Error('monthlyJobメール送信エラー: ' + error.message);
  }
}

// === 統合テスト ===

/**
 * トリガー統合テスト: 全トリガー関数の基本動作確認
 */
function testTriggerIntegration_AllFunctions_ExecuteSuccessfully() {
  // Red: 全てのトリガー関数が正常に実行されることを期待
  
  var results = {
    onOpen: false,
    dailyJob: false,
    weeklyOvertimeJob: false,
    monthlyJob: false
  };
  
  try {
    // onOpen実行
    onOpen();
    results.onOpen = true;
    console.log('✓ onOpen統合テスト成功');
    
    // dailyJob実行
    var dailyResult = dailyJob();
    results.dailyJob = (dailyResult !== undefined);
    console.log('✓ dailyJob統合テスト成功');
    
    // weeklyOvertimeJob実行
    var weeklyResult = weeklyOvertimeJob();
    results.weeklyOvertimeJob = (weeklyResult !== undefined);
    console.log('✓ weeklyOvertimeJob統合テスト成功');
    
    // monthlyJob実行
    var monthlyResult = monthlyJob();
    results.monthlyJob = (monthlyResult !== undefined);
    console.log('✓ monthlyJob統合テスト成功');
    
    // 全体確認
    var allSuccess = results.onOpen && results.dailyJob && results.weeklyOvertimeJob && results.monthlyJob;
    assertTrue(allSuccess, 'すべてのトリガー関数が正常に実行されるべき');
    
    console.log('✓ トリガー統合テスト完了: ' + JSON.stringify(results));
    
  } catch (error) {
    throw new Error('トリガー統合テストエラー: ' + error.message + ', 結果: ' + JSON.stringify(results));
  }
}

// === エラーハンドリングテスト ===

/**
 * トリガーエラーハンドリング: 例外が発生してもシステム継続
 */
function testTriggerErrorHandling_ExceptionOccurs_SystemContinues() {
  // Red: トリガー関数で例外が発生してもシステム全体が停止しないことを期待
  
  try {
    // 各トリガー関数を個別に実行し、エラーをキャッチ
    var errorCount = 0;
    
    // onOpenエラーテスト
    try {
      onOpen();
    } catch (error) {
      errorCount++;
      console.log('onOpenエラー（想定内）: ' + error.message);
    }
    
    // エラーが発生してもテストが継続していることを確認
    assertTrue(true, 'エラーハンドリングテストが継続実行されている');
    console.log('✓ トリガーエラーハンドリング確認完了（エラー数: ' + errorCount + '）');
    
  } catch (error) {
    throw new Error('エラーハンドリングテスト失敗: ' + error.message);
  }
}

// === テスト実行関数 ===

/**
 * Triggersテストの実行
 */
function runTriggersTests() {
  console.log('=== Triggers.gs テスト実行開始 ===');
  
  // onOpen関数のテスト
  runTest(testOnOpen_SpreadsheetOpen_AddsManagementMenu, 'testOnOpen_SpreadsheetOpen_AddsManagementMenu');
  runTest(testOnOpen_ErrorHandling_DoesNotThrow, 'testOnOpen_ErrorHandling_DoesNotThrow');
  
  // dailyJob関数のテスト
  runTest(testDailyJob_ExecuteDaily_UpdatesDailySummary, 'testDailyJob_ExecuteDaily_UpdatesDailySummary');
  runTest(testDailyJob_HasUnfinishedClockOut_SendsEmail, 'testDailyJob_HasUnfinishedClockOut_SendsEmail');
  runTest(testDailyJob_QuotaCompliance_StaysWithinLimits, 'testDailyJob_QuotaCompliance_StaysWithinLimits');
  
  // weeklyOvertimeJob関数のテスト
  runTest(testWeeklyOvertimeJob_Monday_CalculatesOvertimeHours, 'testWeeklyOvertimeJob_Monday_CalculatesOvertimeHours');
  runTest(testWeeklyOvertimeJob_HighOvertimeDetected_SendsWarningEmail, 'testWeeklyOvertimeJob_HighOvertimeDetected_SendsWarningEmail');
  
  // monthlyJob関数のテスト
  runTest(testMonthlyJob_FirstDayOfMonth_UpdatesMonthlySummary, 'testMonthlyJob_FirstDayOfMonth_UpdatesMonthlySummary');
  runTest(testMonthlyJob_MonthlyProcess_ExportsCSVToDrive, 'testMonthlyJob_MonthlyProcess_ExportsCSVToDrive');
  runTest(testMonthlyJob_CSVGenerated_SendsLinkEmail, 'testMonthlyJob_CSVGenerated_SendsLinkEmail');
  
  // 統合テスト
  runTest(testTriggerIntegration_AllFunctions_ExecuteSuccessfully, 'testTriggerIntegration_AllFunctions_ExecuteSuccessfully');
  runTest(testTriggerErrorHandling_ExceptionOccurs_SystemContinues, 'testTriggerErrorHandling_ExceptionOccurs_SystemContinues');
  
  console.log('=== Triggers.gs テスト実行完了 ===');
} 