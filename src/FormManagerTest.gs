/**
 * FormManager.gs のテストファイル
 * TDD実装: Red-Green-Refactorサイクル
 * 
 * テスト対象機能:
 * - フォーム応答の受信処理
 * - フォームデータの変換処理
 * - スプレッドシートへの保存処理
 * - エラーハンドリング
 */

/**
 * フォーム応答受信の基本テスト
 */
function testProcessFormResponse_ValidData_ReturnsSuccess() {
  // Arrange
  var mockFormResponse = {
    timestamp: new Date('2025-07-18T21:00:00'), // 時刻を変更して重複を回避
    employeeId: 'EMP001',
    employeeName: '田中太郎',
    action: 'IN',
    ipAddress: '192.168.1.100',
    remarks: '通常出勤'
  };
  
  // Act
  var result = processFormResponse(mockFormResponse);
  
  // Assert
  assertTrue(result.success, '処理が成功するべき');
  assertEquals('フォーム応答が正常に処理されました', result.message);
  assertNotNull(result.data, 'データが返されるべき');
}

/**
 * 無効なフォームデータのテスト
 */
function testProcessFormResponse_InvalidData_ReturnsError() {
  // Arrange
  var invalidFormResponse = {
    timestamp: null,
    employeeId: '',
    employeeName: '',
    action: 'INVALID',
    ipAddress: '',
    remarks: ''
  };
  
  // Act
  var result = processFormResponse(invalidFormResponse);
  
  // Assert
  assertFalse(result.success, '処理が失敗するべき');
  assertTrue(result.message.includes('エラー'), 'エラーメッセージが含まれるべき');
  assertNull(result.data, 'データはnullであるべき');
}

/**
 * フォームデータ変換のテスト
 */
function testConvertFormDataToLogRaw_ValidData_ReturnsCorrectFormat() {
  // Arrange
  var formData = {
    timestamp: new Date('2025-07-18T09:30:00'),
    employeeId: 'EMP002',
    employeeName: '佐藤花子',
    action: 'OUT',
    ipAddress: '192.168.1.101',
    remarks: '定時退勤'
  };
  
  // Act
  var logData = convertFormDataToLogRaw(formData);
  
  // Assert
  assertNotNull(logData, 'ログデータが返されるべき');
  assertEquals(6, logData.length, '6つのフィールドが含まれるべき');
  assertEquals(formData.timestamp, logData[0], 'タイムスタンプが正しく設定されるべき');
  assertEquals(formData.employeeId, logData[1], '社員IDが正しく設定されるべき');
  assertEquals(formData.employeeName, logData[2], '氏名が正しく設定されるべき');
  assertEquals(formData.action, logData[3], 'アクションが正しく設定されるべき');
  assertEquals(formData.ipAddress, logData[4], 'IPアドレスが正しく設定されるべき');
  assertEquals(formData.remarks, logData[5], '備考が正しく設定されるべき');
}

/**
 * フォームデータのバリデーションテスト
 */
function testValidateFormData_ValidData_ReturnsTrue() {
  // Arrange
  var validData = {
    timestamp: new Date(),
    employeeId: 'EMP003',
    employeeName: '山田次郎',
    action: 'IN',
    ipAddress: '192.168.1.102',
    remarks: '出勤'
  };
  
  // Act
  var isValid = validateFormData(validData);
  
  // Assert
  assertTrue(isValid, '有効なデータはtrueを返すべき');
}

/**
 * フォームデータのバリデーション（無効データ）テスト
 */
function testValidateFormData_InvalidData_ReturnsFalse() {
  // Arrange
  var invalidData = {
    timestamp: null,
    employeeId: '',
    employeeName: '',
    action: 'INVALID_ACTION',
    ipAddress: '',
    remarks: ''
  };
  
  // Act
  var isValid = validateFormData(invalidData);
  
  // Assert
  assertFalse(isValid, '無効なデータはfalseを返すべき');
}

/**
 * スプレッドシート保存のテスト
 */
function testSaveFormDataToSheet_ValidData_SavesSuccessfully() {
  // Arrange
  var logData = [
    new Date('2025-07-18T08:00:00'),
    'EMP004',
    '鈴木三郎',
    'IN',
    '192.168.1.103',
    '出勤'
  ];
  
  // Act
  var result = saveFormDataToSheet(logData);
  
  // Assert
  assertTrue(result.success, '保存が成功するべき');
  assertTrue(result.message.includes('保存'), '保存メッセージが含まれるべき');
  assertNotNull(result.rowNumber, '行番号が返されるべき');
}

/**
 * フォーム応答の重複チェックテスト
 */
function testCheckDuplicateFormResponse_NewData_ReturnsFalse() {
  // Arrange
  var newFormData = {
    timestamp: new Date('2025-07-18T10:00:00'),
    employeeId: 'EMP005',
    employeeName: '高橋四郎',
    action: 'IN',
    ipAddress: '192.168.1.104',
    remarks: '出勤'
  };
  
  // Act
  var isDuplicate = checkDuplicateFormResponse(newFormData);
  
  // Assert
  assertFalse(isDuplicate, '新しいデータは重複ではないべき');
}

/**
 * フォーム応答の重複チェック（重複データ）テスト
 */
function testCheckDuplicateFormResponse_DuplicateData_ReturnsTrue() {
  // Arrange
  var duplicateFormData = {
    timestamp: new Date('2025-07-18T09:00:00'),
    employeeId: 'EMP001',
    employeeName: '田中太郎',
    action: 'IN',
    ipAddress: '192.168.1.100',
    remarks: '出勤'
  };
  
  // 事前に同じデータを保存
  saveFormDataToSheet(convertFormDataToLogRaw(duplicateFormData));
  
  // Act
  var isDuplicate = checkDuplicateFormResponse(duplicateFormData);
  
  // Assert
  assertTrue(isDuplicate, '重複データはtrueを返すべき');
}

/**
 * フォーム応答処理の統合テスト
 */
function testProcessFormResponse_IntegrationTest_CompleteWorkflow() {
  // Arrange
  var formResponse = {
    timestamp: new Date('2025-07-18T22:00:00'), // 時刻を変更して重複を回避
    employeeId: 'EMP006',
    employeeName: '伊藤五郎',
    action: 'OUT',
    ipAddress: '192.168.1.105',
    remarks: '退勤'
  };
  
  // Act
  var result = processFormResponse(formResponse);
  
  // Assert
  assertTrue(result.success, '統合テストが成功するべき');
  assertTrue(result.message.includes('処理'), '処理メッセージが含まれるべき');
  assertNotNull(result.data, 'データが返されるべき');
  
  // 保存されたデータの確認
  var savedData = result.data;
  assertEquals(formResponse.employeeId, savedData.employeeId, '保存された社員IDが一致するべき');
  assertEquals(formResponse.action, savedData.action, '保存されたアクションが一致するべき');
}

/**
 * エラーハンドリングのテスト
 */
function testProcessFormResponse_ExceptionHandling_ReturnsError() {
  // Arrange
  var malformedData = {
    timestamp: 'invalid-date',
    employeeId: 123, // 数値（文字列であるべき）
    employeeName: null,
    action: 'IN',
    ipAddress: '192.168.1.106',
    remarks: 'テスト'
  };
  
  // Act
  var result = processFormResponse(malformedData);
  
  // Assert
  assertFalse(result.success, '例外が発生した場合、失敗するべき');
  assertTrue(result.message.includes('エラー'), 'エラーメッセージが含まれるべき');
  assertNull(result.data, 'データはnullであるべき');
}

/**
 * フォーム応答の一括処理テスト
 */
function testProcessMultipleFormResponses_ValidData_ProcessesAll() {
  // Arrange
  var formResponses = [
    {
      timestamp: new Date('2025-07-18T23:00:00'), // 時刻を変更して重複を回避
      employeeId: 'EMP007',
      employeeName: '渡辺六郎',
      action: 'IN',
      ipAddress: '192.168.1.107',
      remarks: '出勤'
    },
    {
      timestamp: new Date('2025-07-18T23:30:00'), // 時刻を変更して重複を回避
      employeeId: 'EMP007',
      employeeName: '渡辺六郎',
      action: 'OUT',
      ipAddress: '192.168.1.107',
      remarks: '退勤'
    }
  ];
  
  // Act
  var results = processMultipleFormResponses(formResponses);
  
  // Assert
  assertNotNull(results, '結果が返されるべき');
  assertEquals(2, results.length, '2件の結果が返されるべき');
  assertTrue(results[0].success, '1件目が成功するべき');
  assertTrue(results[1].success, '2件目が成功するべき');
}

/**
 * フォーム応答の統計情報取得テスト
 */
function testGetFormResponseStats_ValidData_ReturnsStats() {
  // Arrange
  // テストデータを事前に投入
  var testData = [
    {
      timestamp: new Date('2025-07-18T08:00:00'),
      employeeId: 'EMP008',
      employeeName: '中村七郎',
      action: 'IN',
      ipAddress: '192.168.1.108',
      remarks: '出勤'
    }
  ];
  processFormResponse(testData[0]);
  
  // Act
  var stats = getFormResponseStats();
  
  // Assert
  assertNotNull(stats, '統計情報が返されるべき');
  assertTrue(stats.totalResponses >= 1, '総応答数が1以上であるべき');
  assertTrue(stats.successCount >= 1, '成功数が1以上であるべき');
  assertTrue(stats.errorCount >= 0, 'エラー数が0以上であるべき');
} 