/**
 * MailManager.gs のテストファイル
 * TDD実装: Red-Green-Refactorサイクル
 * 
 * テスト対象機能:
 * - 未退勤者メール送信
 * - 月次レポートメール送信
 * - メールテンプレート生成
 * - クォータ管理
 * - エラーハンドリング
 */

/**
 * テスト用のAPP_CONFIG初期化
 */
function initializeTestAppConfig() {
  if (typeof APP_CONFIG === 'undefined') {
    APP_CONFIG = {};
  }
  APP_CONFIG['ADMIN_EMAILS'] = ['manager@example.com', 'hr@example.com'];
  APP_CONFIG['EMAIL_DAILY_QUOTA'] = 100;
  // getTestModeConfigのダミー化
  globalThis.getTestModeConfig = function() { return false; };
  
  // メールモードを必ずMOCKに設定
  setEmailMode('MOCK');
  clearMockEmailData();
}

/**
 * 未退勤者メール送信の基本テスト
 */
function testSendUnfinishedClockOutEmail_ValidData_SendsEmail() {
  // Arrange
  initializeTestAppConfig();
  var unfinishedEmployees = [
    {
      employeeId: 'EMP001',
      employeeName: '田中太郎',
      email: 'tanaka@example.com',
      clockInTime: '09:00',
      currentTime: '18:30'
    },
    {
      employeeId: 'EMP002',
      employeeName: '佐藤花子',
      email: 'sato@example.com',
      clockInTime: '08:30',
      currentTime: '18:30'
    }
  ];
  var targetDate = new Date('2025-07-20');
  
  // Act
  var result = sendUnfinishedClockOutEmail_MailManager(unfinishedEmployees, targetDate);
  
  // Debug: 結果オブジェクトの内容確認
  console.log('テスト結果デバッグ:', JSON.stringify(result));
  console.log('result.sentCount:', result.sentCount);
  console.log('result.success:', result.success);
  
  // Assert
  assertTrue(result.success, 'メール送信が成功するべき');
  assertEquals(2, result.sentCount, '管理者2名にメールが送信されるべき');
  assertTrue(result.message.includes('送信完了'), '送信完了メッセージが含まれるべき');
}

/**
 * 未退勤者なしの場合のテスト
 */
function testSendUnfinishedClockOutEmail_NoUnfinished_ReturnsSuccess() {
  // Arrange
  initializeTestAppConfig();
  var unfinishedEmployees = [];
  var targetDate = new Date('2025-07-20');
  
  // Act
  var result = sendUnfinishedClockOutEmail_MailManager(unfinishedEmployees, targetDate);
  
  // Assert
  assertTrue(result.success, '処理が成功するべき');
  assertEquals(0, result.sentCount, '送信件数は0であるべき');
  assertTrue(result.message.includes('未退勤者なし'), '未退勤者なしメッセージが含まれるべき');
}

/**
 * 月次レポートメール送信のテスト
 */
function testSendMonthlyReportEmail_ValidData_SendsEmail() {
  // Arrange
  initializeTestAppConfig();
  var reportData = {
    month: '2025-07',
    totalEmployees: 5,
    totalWorkDays: 22,
    averageWorkHours: 8.5,
    overtimeHours: 12.5,
    csvFileName: 'monthly_report_2025-07.csv'
  };
  var adminEmails = ['manager@example.com', 'hr@example.com'];
  
  // Act
  var result = sendMonthlyReportEmail_MailManager(reportData, adminEmails);
  
  // Assert
  assertTrue(result.success, 'メール送信が成功するべき');
  assertEquals(2, result.sentCount, '2件のメールが送信されるべき');
  assertTrue(result.message.includes('月次レポート'), '月次レポートメッセージが含まれるべき');
}

/**
 * メールテンプレート生成のテスト
 */
function testGenerateUnfinishedClockOutEmailTemplate_ValidData_ReturnsCorrectTemplate() {
  // Arrange
  var employee = {
    employeeId: 'EMP001',
    employeeName: '田中太郎',
    email: 'tanaka@example.com',
    clockInTime: '09:00',
    currentTime: '18:30'
  };
  var targetDate = new Date('2025-07-20');
  
  // Act
  var template = generateUnfinishedClockOutEmailTemplate(employee, targetDate);
  
  // Assert
  assertNotNull(template, 'テンプレートが返されるべき');
  assertTrue(template.subject.includes('未退勤'), '件名に未退勤が含まれるべき');
  assertTrue(template.body.includes('田中太郎'), '本文に氏名が含まれるべき');
  assertTrue(template.body.includes('09:00'), '本文に出勤時刻が含まれるべき');
  assertTrue(template.body.includes('18:30'), '本文に現在時刻が含まれるべき');
}

/**
 * 月次レポートテンプレート生成のテスト
 */
function testGenerateMonthlyReportEmailTemplate_ValidData_ReturnsCorrectTemplate() {
  // Arrange
  var reportData = {
    month: '2025-07',
    totalEmployees: 5,
    totalWorkDays: 22,
    averageWorkHours: 8.5,
    overtimeHours: 12.5,
    csvFileName: 'monthly_report_2025-07.csv'
  };
  
  // Act
  var template = generateMonthlyReportEmailTemplate(reportData);
  
  // Assert
  assertNotNull(template, 'テンプレートが返されるべき');
  assertTrue(template.subject.includes('月次レポート'), '件名に月次レポートが含まれるべき');
  assertTrue(template.subject.includes('2025-07'), '件名に月が含まれるべき');
  assertTrue(template.body.includes('5'), '本文に従業員数が含まれるべき');
  assertTrue(template.body.includes('8.5'), '本文に平均勤務時間が含まれるべき');
  assertTrue(template.body.includes('12.5'), '本文に残業時間が含まれるべき');
}

/**
 * クォータ管理のテスト
 */
function testCheckEmailQuota_WithinLimit_ReturnsTrue() {
  // Arrange
  var currentCount = 50; // 現在の送信数
  var maxQuota = 100; // 最大クォータ
  
  // Act
  var canSend = checkEmailQuota(currentCount, maxQuota);
  
  // Assert
  assertTrue(canSend, 'クォータ内であれば送信可能であるべき');
}

/**
 * クォータ超過のテスト
 */
function testCheckEmailQuota_ExceededLimit_ReturnsFalse() {
  // Arrange
  initializeTestAppConfig();
  // クォータ超過状態を作るため、getTodayEmailCountを上書き
  var originalGetTodayEmailCount = getTodayEmailCount;
  getTodayEmailCount = function() { return 100; };

  // Act
  var result = checkEmailQuota(1);

  // Assert
  assertFalse(result, 'クォータ超過であれば送信不可であるべき');

  // 後始末
  getTodayEmailCount = originalGetTodayEmailCount;
}

/**
 * クォータ管理の境界値テスト
 * 境界値（maxQuota - 1, maxQuota, maxQuota + 1）での動作を検証
 */
function testCheckEmailQuota_BoundaryValues_ReturnsCorrectResults() {
  // Arrange
  var maxQuota = 100; // 最大クォータ
  
  // Test Case 1: クォータ未満（maxQuota - 1）
  var currentCount1 = maxQuota - 1; // 99
  
  // Test Case 2: クォータちょうど（maxQuota）
  var currentCount2 = maxQuota; // 100
  
  // Test Case 3: クォータ超過（maxQuota + 1）
  var currentCount3 = maxQuota + 1; // 101
  
  // Act & Assert - Test Case 1: クォータ未満
  var canSend1 = checkEmailQuota(currentCount1, maxQuota);
  assertTrue(canSend1, 'クォータ未満（' + currentCount1 + '）では送信可能であるべき');
  
  // Act & Assert - Test Case 2: クォータちょうど
  var canSend2 = checkEmailQuota(currentCount2, maxQuota);
  assertFalse(canSend2, 'クォータちょうど（' + currentCount2 + '）では送信不可であるべき');
  
  // Act & Assert - Test Case 3: クォータ超過
  var canSend3 = checkEmailQuota(currentCount3, maxQuota);
  assertFalse(canSend3, 'クォータ超過（' + currentCount3 + '）では送信不可であるべき');
  
  // 追加の境界値テスト: ゼロクォータ
  var canSendZero = checkEmailQuota(0, 0);
  assertFalse(canSendZero, 'ゼロクォータでは送信不可であるべき');
  
  // 追加の境界値テスト: 負の値
  var canSendNegative = checkEmailQuota(-1, 100);
  assertTrue(canSendNegative, '負の現在送信数では送信可能であるべき');
}

/**
 * メール送信エラーハンドリングのテスト
 */
function testSendEmail_ErrorOccurs_ReturnsError() {
  // Arrange
  initializeTestAppConfig();
  var emailData = {
    to: null,  // 必須パラメータ不足でエラーを発生させる
    subject: 'テスト',
    body: 'テスト本文'
  };
  
  // Act
  var result = sendEmail(emailData);
  
  // Assert
  assertFalse(result.success, 'エラーが発生した場合、失敗するべき');
  assertTrue(result.message.includes('エラー'), 'エラーメッセージが含まれるべき');
}

/**
 * メール送信統計のテスト
 */
function testGetEmailStats_ValidData_ReturnsStats() {
  // Arrange
  // テスト用のメール送信履歴を設定
  var testEmailData = {
    to: 'test@example.com',
    subject: 'テストメール',
    body: 'これはテスト用のメールです。'
  };
  
  // テストメールを送信して統計データを生成
  var sendResult = sendEmail(testEmailData);
  assertTrue(sendResult.success, 'テストメールの送信が成功するべき');
  
  // Act
  var stats = getEmailStats();
  
  // Assert
  assertNotNull(stats, '統計情報が返されるべき');
  assertTrue(stats.totalSent >= 1, '総送信数が1以上であるべき（テストメール送信後）');
  assertTrue(stats.quotaRemaining >= 0, '残りクォータが0以上であるべき');
  assertNotNull(stats.lastSentDate, '最終送信日がnullでないべき');
  assertTrue(stats.lastSentDate instanceof Date || typeof stats.lastSentDate === 'string', '最終送信日が有効な形式であるべき');
  
  // 追加の検証: 送信記録の詳細確認
  if (stats.totalSent > 0) {
    assertTrue(stats.quotaRemaining < 100, 'クォータが使用されているべき');
    assertTrue(stats.successRate >= 0 && stats.successRate <= 100, '成功率が0-100%の範囲であるべき');
  }
}

/**
 * メール送信の統合テスト
 */
function testSendUnfinishedClockOutEmail_IntegrationTest_CompleteWorkflow() {
  // Arrange
  var unfinishedEmployees = [
    {
      employeeId: 'EMP001',
      employeeName: '田中太郎',
      email: 'tanaka@example.com',
      clockInTime: '09:00',
      currentTime: '18:30'
    }
  ];
  var targetDate = new Date('2025-07-20');
  
  // Act
  var result = sendUnfinishedClockOutEmail(unfinishedEmployees, targetDate);
  
  // Assert
  assertTrue(result.success, '統合テストが成功するべき');
  assertEquals(1, result.sentCount, '1件のメールが送信されるべき');
  assertTrue(result.message.includes('送信完了'), '送信完了メッセージが含まれるべき');
  
  // 統計情報の確認
  var stats = getEmailStats();
  assertTrue(stats.totalSent >= 1, '統計に送信記録が残るべき');
}

/**
 * メール送信の一括処理テスト
 */
function testSendMultipleEmails_ValidData_SendsAllEmails() {
  // Arrange
  var emailList = [
    {
      to: 'user1@example.com',
      subject: 'テスト1',
      body: 'テスト本文1'
    },
    {
      to: 'user2@example.com',
      subject: 'テスト2',
      body: 'テスト本文2'
    }
  ];
  
  // Act
  var results = sendMultipleEmails(emailList);
  
  // Assert
  assertNotNull(results, '結果が返されるべき');
  assertEquals(2, results.length, '2件の結果が返されるべき');
  assertTrue(results[0].success, '1件目が成功するべき');
  assertTrue(results[1].success, '2件目が成功するべき');
}

/**
 * メール送信の例外処理テスト
 */
function testSendEmail_ExceptionHandling_ReturnsError() {
  // Arrange
  var invalidEmailData = {
    to: null,
    subject: null,
    body: null
  };
  
  // Act
  var result = sendEmail(invalidEmailData);
  
  // Assert
  assertFalse(result.success, '例外が発生した場合、失敗するべき');
  assertTrue(result.message.includes('エラー'), 'エラーメッセージが含まれるべき');
  assertNull(result.sentCount, '送信件数はnullであるべき');
} 