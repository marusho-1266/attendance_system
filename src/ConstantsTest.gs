/**
 * Constants.gs のテストケース
 * TDD Red フェーズ: 期待する動作をテストで定義
 */

/**
 * getColumnIndex関数のテスト - EMPLOYEE シート
 */
function testGetColumnIndex_Employee_Name_ReturnsCorrectIndex() {
  // Red: 従業員シートのNAME列インデックスが1を返すことを期待
  var result = getColumnIndex('EMPLOYEE', 'NAME');
  assertEquals(1, result, 'EMPLOYEE シートのNAME列は1を返すべき');
}

function testGetColumnIndex_Employee_Gmail_ReturnsCorrectIndex() {
  // Red: 従業員シートのGMAIL列インデックスが2を返すことを期待
  var result = getColumnIndex('EMPLOYEE', 'GMAIL');
  assertEquals(2, result, 'EMPLOYEE シートのGMAIL列は2を返すべき');
}

function testGetColumnIndex_Employee_EmployeeId_ReturnsCorrectIndex() {
  // Red: 従業員シートのEMPLOYEE_ID列インデックスが0を返すことを期待
  var result = getColumnIndex('EMPLOYEE', 'EMPLOYEE_ID');
  assertEquals(0, result, 'EMPLOYEE シートのEMPLOYEE_ID列は0を返すべき');
}

/**
 * getColumnIndex関数のテスト - HOLIDAY シート
 */
function testGetColumnIndex_Holiday_Date_ReturnsCorrectIndex() {
  // Red: 休日シートのDATE列インデックスが0を返すことを期待
  var result = getColumnIndex('HOLIDAY', 'DATE');
  assertEquals(0, result, 'HOLIDAY シートのDATE列は0を返すべき');
}

function testGetColumnIndex_Holiday_Name_ReturnsCorrectIndex() {
  // Red: 休日シートのNAME列インデックスが1を返すことを期待
  var result = getColumnIndex('HOLIDAY', 'NAME');
  assertEquals(1, result, 'HOLIDAY シートのNAME列は1を返すべき');
}

/**
 * getColumnIndex関数のテスト - LOG_RAW シート
 */
function testGetColumnIndex_LogRaw_Timestamp_ReturnsCorrectIndex() {
  // Red: ログシートのTIMESTAMP列インデックスが0を返すことを期待
  var result = getColumnIndex('LOG_RAW', 'TIMESTAMP');
  assertEquals(0, result, 'LOG_RAW シートのTIMESTAMP列は0を返すべき');
}

function testGetColumnIndex_LogRaw_Action_ReturnsCorrectIndex() {
  // Red: ログシートのACTION列インデックスが3を返すことを期待
  var result = getColumnIndex('LOG_RAW', 'ACTION');
  assertEquals(3, result, 'LOG_RAW シートのACTION列は3を返すべき');
}

/**
 * getColumnIndex関数のテスト - 無効なシートタイプ
 */
function testGetColumnIndex_InvalidSheetType_ThrowsError() {
  // Red: 無効なシートタイプでエラーが発生することを期待
  try {
    getColumnIndex('INVALID_SHEET', 'ANY_COLUMN');
    assert(false, '無効なシートタイプでエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.indexOf('Unknown sheet type') !== -1, 'エラーメッセージに"Unknown sheet type"が含まれるべき');
  }
}

/**
 * getSheetName関数のテスト
 */
function testGetSheetName_MasterEmployee_ReturnsCorrectName() {
  // Red: MASTER_EMPLOYEEが正しいシート名を返すことを期待
  var result = getSheetName('MASTER_EMPLOYEE');
  assertEquals('Master_Employee', result, 'MASTER_EMPLOYEEシート名が正しく返されるべき');
}

function testGetSheetName_MasterHoliday_ReturnsCorrectName() {
  // Red: MASTER_HOLIDAYが正しいシート名を返すことを期待
  var result = getSheetName('MASTER_HOLIDAY');
  assertEquals('Master_Holiday', result, 'MASTER_HOLIDAYシート名が正しく返されるべき');
}

function testGetSheetName_LogRaw_ReturnsCorrectName() {
  // Red: LOG_RAWが正しいシート名を返すことを期待
  var result = getSheetName('LOG_RAW');
  assertEquals('Log_Raw', result, 'LOG_RAWシート名が正しく返されるべき');
}

/**
 * getActionConstant関数のテスト
 */
function testGetActionConstant_ClockIn_ReturnsCorrectValue() {
  // Red: CLOCK_INが'IN'を返すことを期待
  var result = getActionConstant('CLOCK_IN');
  assertEquals('IN', result, 'CLOCK_INアクションは"IN"を返すべき');
}

function testGetActionConstant_ClockOut_ReturnsCorrectValue() {
  // Red: CLOCK_OUTが'OUT'を返すことを期待
  var result = getActionConstant('CLOCK_OUT');
  assertEquals('OUT', result, 'CLOCK_OUTアクションは"OUT"を返すべき');
}

function testGetActionConstant_BreakStart_ReturnsCorrectValue() {
  // Red: BREAK_STARTが'BRK_IN'を返すことを期待
  var result = getActionConstant('BREAK_START');
  assertEquals('BRK_IN', result, 'BREAK_STARTアクションは"BRK_IN"を返すべき');
}

function testGetActionConstant_BreakEnd_ReturnsCorrectValue() {
  // Red: BREAK_ENDが'BRK_OUT'を返すことを期待
  var result = getActionConstant('BREAK_END');
  assertEquals('BRK_OUT', result, 'BREAK_ENDアクションは"BRK_OUT"を返すべき');
}

/**
 * getAppConfig関数のテスト
 */
function testGetAppConfig_StandardWorkHours_ReturnsCorrectValue() {
  // Red: STANDARD_WORK_HOURSが8を返すことを期待
  var result = getAppConfig('STANDARD_WORK_HOURS');
  assertEquals(8, result, 'STANDARD_WORK_HOURSは8を返すべき');
}

function testGetAppConfig_MaxWorkHoursPerDay_ReturnsCorrectValue() {
  // Red: MAX_WORK_HOURS_PER_DAYが24を返すことを期待
  var result = getAppConfig('MAX_WORK_HOURS_PER_DAY');
  assertEquals(24, result, 'MAX_WORK_HOURS_PER_DAYは24を返すべき');
}

function testGetAppConfig_BreakTimeAutoDeduct_ReturnsCorrectValue() {
  // Red: BREAK_TIME_AUTO_DEDUCTが45を返すことを期待
  var result = getAppConfig('BREAK_TIME_AUTO_DEDUCT');
  assertEquals(45, result, 'BREAK_TIME_AUTO_DEDUCTは45を返すべき');
}

/**
 * Constants テスト実行関数
 */
function runConstantsTests() {
  console.log('=== Constants.gs テスト実行開始 ===');
  
  // getColumnIndex関数のテスト
  runTest(testGetColumnIndex_Employee_Name_ReturnsCorrectIndex, 'testGetColumnIndex_Employee_Name_ReturnsCorrectIndex');
  runTest(testGetColumnIndex_Employee_Gmail_ReturnsCorrectIndex, 'testGetColumnIndex_Employee_Gmail_ReturnsCorrectIndex');
  runTest(testGetColumnIndex_Employee_EmployeeId_ReturnsCorrectIndex, 'testGetColumnIndex_Employee_EmployeeId_ReturnsCorrectIndex');
  runTest(testGetColumnIndex_Holiday_Date_ReturnsCorrectIndex, 'testGetColumnIndex_Holiday_Date_ReturnsCorrectIndex');
  runTest(testGetColumnIndex_Holiday_Name_ReturnsCorrectIndex, 'testGetColumnIndex_Holiday_Name_ReturnsCorrectIndex');
  runTest(testGetColumnIndex_LogRaw_Timestamp_ReturnsCorrectIndex, 'testGetColumnIndex_LogRaw_Timestamp_ReturnsCorrectIndex');
  runTest(testGetColumnIndex_LogRaw_Action_ReturnsCorrectIndex, 'testGetColumnIndex_LogRaw_Action_ReturnsCorrectIndex');
  runTest(testGetColumnIndex_InvalidSheetType_ThrowsError, 'testGetColumnIndex_InvalidSheetType_ThrowsError');
  
  // getSheetName関数のテスト
  runTest(testGetSheetName_MasterEmployee_ReturnsCorrectName, 'testGetSheetName_MasterEmployee_ReturnsCorrectName');
  runTest(testGetSheetName_MasterHoliday_ReturnsCorrectName, 'testGetSheetName_MasterHoliday_ReturnsCorrectName');
  runTest(testGetSheetName_LogRaw_ReturnsCorrectName, 'testGetSheetName_LogRaw_ReturnsCorrectName');
  
  // getActionConstant関数のテスト
  runTest(testGetActionConstant_ClockIn_ReturnsCorrectValue, 'testGetActionConstant_ClockIn_ReturnsCorrectValue');
  runTest(testGetActionConstant_ClockOut_ReturnsCorrectValue, 'testGetActionConstant_ClockOut_ReturnsCorrectValue');
  runTest(testGetActionConstant_BreakStart_ReturnsCorrectValue, 'testGetActionConstant_BreakStart_ReturnsCorrectValue');
  runTest(testGetActionConstant_BreakEnd_ReturnsCorrectValue, 'testGetActionConstant_BreakEnd_ReturnsCorrectValue');
  
  // getAppConfig関数のテスト
  runTest(testGetAppConfig_StandardWorkHours_ReturnsCorrectValue, 'testGetAppConfig_StandardWorkHours_ReturnsCorrectValue');
  runTest(testGetAppConfig_MaxWorkHoursPerDay_ReturnsCorrectValue, 'testGetAppConfig_MaxWorkHoursPerDay_ReturnsCorrectValue');
  runTest(testGetAppConfig_BreakTimeAutoDeduct_ReturnsCorrectValue, 'testGetAppConfig_BreakTimeAutoDeduct_ReturnsCorrectValue');
  
  console.log('=== Constants.gs テスト実行完了 ===');
} 