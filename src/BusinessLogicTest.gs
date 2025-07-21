/**
 * BusinessLogic.gs のテストケース
 * TDD実装: Red-Green-Refactorサイクル
 * 
 * テスト対象関数:
 * - isHoliday(date): 休日判定
 * - calcWorkTime(startTime, endTime, breakMinutes): 勤務時間計算
 * - getEmployee(email): 従業員情報取得
 */

// === 休日判定関数のテスト ===

/**
 * isHoliday関数: 平日の判定テスト
 */
function testIsHoliday_Weekday_ReturnsFalse() {
  // 2025-07-10 (木曜日) - 平日
  var weekday = new Date(2025, 6, 10); // 月は0ベース
  var result = isHoliday(weekday);
  
  assertFalse(result, '平日は休日ではない');
}

/**
 * isHoliday関数: 土曜日の判定テスト
 */
function testIsHoliday_Saturday_ReturnsTrue() {
  // 2025-07-12 (土曜日)
  var saturday = new Date(2025, 6, 12);
  var result = isHoliday(saturday);
  
  assertTrue(result, '土曜日は休日');
}

/**
 * isHoliday関数: 日曜日の判定テスト
 */
function testIsHoliday_Sunday_ReturnsTrue() {
  // 2025-07-13 (日曜日)
  var sunday = new Date(2025, 6, 13);
  var result = isHoliday(sunday);
  
  assertTrue(result, '日曜日は休日');
}

/**
 * isHoliday関数: 祝日の判定テスト
 */
function testIsHoliday_NationalHoliday_ReturnsTrue() {
  // 2025-01-01 (元日) - Master_Holidayに登録済み想定
  var newYear = new Date(2025, 0, 1);
  var result = isHoliday(newYear);
  
  assertTrue(result, '元日は祝日');
}

/**
 * isHoliday関数: null値のテスト
 */
function testIsHoliday_NullDate_ThrowsError() {
  try {
    isHoliday(null);
    assert(false, 'null値で例外が発生し、catchされるべき');
  } catch (error) {
    assertTrue(error.message.includes('Invalid date'), 'null値入力時は"Invalid date"エラーが返るべき');
  }
}

// === 勤務時間計算関数のテスト ===

/**
 * calcWorkTime関数: 通常勤務の計算テスト
 */
function testCalcWorkTime_StandardWork_ReturnsCorrectTime() {
  // 9:00-18:00、休憩60分 = 8時間
  var result = calcWorkTime('09:00', '18:00', 60);
  
  assertEquals(480, result, '標準勤務(8時間)の計算'); // 8時間 = 480分
}

/**
 * calcWorkTime関数: 残業ありの計算テスト
 */
function testCalcWorkTime_Overtime_ReturnsCorrectTime() {
  // 9:00-20:00、休憩60分 = 10時間
  var result = calcWorkTime('09:00', '20:00', 60);
  
  assertEquals(600, result, '残業込み(10時間)の計算'); // 10時間 = 600分
}

/**
 * calcWorkTime関数: 短時間勤務の計算テスト
 */
function testCalcWorkTime_PartTime_ReturnsCorrectTime() {
  // 10:00-15:00、休憩0分 = 5時間
  var result = calcWorkTime('10:00', '15:00', 0);
  
  assertEquals(300, result, '短時間勤務(5時間)の計算'); // 5時間 = 300分
}

/**
 * calcWorkTime関数: 日跨ぎ勤務の計算テスト
 */
function testCalcWorkTime_OverMidnight_ReturnsCorrectTime() {
  // 22:00-06:00、休憩60分 = 7時間
  var result = calcWorkTime('22:00', '06:00', 60);
  
  assertEquals(420, result, '日跨ぎ勤務(7時間)の計算'); // 7時間 = 420分
}

/**
 * calcWorkTime関数: 不正な時刻フォーマットのテスト
 */
function testCalcWorkTime_InvalidTimeFormat_ThrowsError() {
  try {
    calcWorkTime('25:00', '18:00', 60);
    assert(false, '不正な時刻でエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.includes('Invalid time'), '不正な時刻で適切なエラーメッセージ');
  }
}

/**
 * calcWorkTime関数: 負の休憩時間のテスト
 */
function testCalcWorkTime_NegativeBreak_ThrowsError() {
  try {
    calcWorkTime('09:00', '18:00', -30);
    assert(false, '負の休憩時間でエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.includes('Invalid break'), '負の休憩時間で適切なエラーメッセージ');
  }
}

// === 従業員情報取得関数のテスト ===

/**
 * getEmployee関数: 有効なメールアドレスのテスト
 */
function testGetEmployee_ValidEmail_ReturnsEmployeeInfo() {
  // テスト用の有効なメールアドレス
  var email = 'yamada@example.com';
  var result = getEmployee(email);
  
  assertNotNull(result, '有効なメールアドレスで従業員情報を取得');
  assertNotNull(result.employeeId, '従業員IDが取得される');
  assertNotNull(result.name, '従業員名が取得される');
  assertEquals(email, result.gmail, 'メールアドレスが一致');
}

/**
 * getEmployee関数: 無効なメールアドレスのテスト
 */
function testGetEmployee_InvalidEmail_ReturnsNull() {
  var result = getEmployee('notfound@example.com');
  
  assertNull(result, '無効なメールアドレスではnullを返す');
}

/**
 * getEmployee関数: 空文字列のテスト
 */
function testGetEmployee_EmptyEmail_ThrowsError() {
  try {
    getEmployee('');
    assert(false, '空文字列でエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.includes('Empty email'), '空文字列で適切なエラーメッセージ');
  }
}

/**
 * getEmployee関数: null値のテスト
 */
function testGetEmployee_NullEmail_ThrowsError() {
  try {
    getEmployee(null);
    assert(false, 'null値でエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.includes('Empty email'), 'null値で適切なエラーメッセージ');
  }
}

/**
 * getEmployee関数: 不正なメールフォーマットのテスト
 */
function testGetEmployee_InvalidEmailFormat_ThrowsError() {
  try {
    getEmployee('invalid-email');
    assert(false, '不正なメールフォーマットでエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.includes('Invalid email'), '不正なメールフォーマットで適切なエラーメッセージ');
  }
}

// === テスト実行関数 ===

/**
 * BusinessLogicテストの実行
 */
function runBusinessLogicTests() {
  console.log('=== BusinessLogic.gs テスト実行開始 ===');
  
  // isHoliday関数のテスト
  runTest(testIsHoliday_Weekday_ReturnsFalse, 'testIsHoliday_Weekday_ReturnsFalse');
  runTest(testIsHoliday_Saturday_ReturnsTrue, 'testIsHoliday_Saturday_ReturnsTrue');
  runTest(testIsHoliday_Sunday_ReturnsTrue, 'testIsHoliday_Sunday_ReturnsTrue');
  runTest(testIsHoliday_NationalHoliday_ReturnsTrue, 'testIsHoliday_NationalHoliday_ReturnsTrue');
  runTest(testIsHoliday_NullDate_ThrowsError, 'testIsHoliday_NullDate_ThrowsError');
  
  // calcWorkTime関数のテスト
  runTest(testCalcWorkTime_StandardWork_ReturnsCorrectTime, 'testCalcWorkTime_StandardWork_ReturnsCorrectTime');
  runTest(testCalcWorkTime_Overtime_ReturnsCorrectTime, 'testCalcWorkTime_Overtime_ReturnsCorrectTime');
  runTest(testCalcWorkTime_PartTime_ReturnsCorrectTime, 'testCalcWorkTime_PartTime_ReturnsCorrectTime');
  runTest(testCalcWorkTime_OverMidnight_ReturnsCorrectTime, 'testCalcWorkTime_OverMidnight_ReturnsCorrectTime');
  runTest(testCalcWorkTime_InvalidTimeFormat_ThrowsError, 'testCalcWorkTime_InvalidTimeFormat_ThrowsError');
  runTest(testCalcWorkTime_NegativeBreak_ThrowsError, 'testCalcWorkTime_NegativeBreak_ThrowsError');
  
  // getEmployee関数のテスト
  runTest(testGetEmployee_ValidEmail_ReturnsEmployeeInfo, 'testGetEmployee_ValidEmail_ReturnsEmployeeInfo');
  runTest(testGetEmployee_InvalidEmail_ReturnsNull, 'testGetEmployee_InvalidEmail_ReturnsNull');
  runTest(testGetEmployee_EmptyEmail_ThrowsError, 'testGetEmployee_EmptyEmail_ThrowsError');
  runTest(testGetEmployee_NullEmail_ThrowsError, 'testGetEmployee_NullEmail_ThrowsError');
  runTest(testGetEmployee_InvalidEmailFormat_ThrowsError, 'testGetEmployee_InvalidEmailFormat_ThrowsError');
  
  console.log('=== BusinessLogic.gs テスト実行完了 ===');
} 