/**
 * CommonUtils.gs のテストケース
 * TDD実装の各関数をテスト
 */

/**
 * formatDate関数のテスト
 */
function testFormatDate_ValidDate_ReturnsFormattedString() {
  // Red: 有効な日付を正しいフォーマットで返すことを期待
  var testDate = new Date(2025, 6, 10); // 2025年7月10日
  var result = formatDate(testDate);
  assertEquals('2025-07-10', result, 'formatDate は正しいフォーマットを返すべき');
}

function testFormatDate_InvalidDate_ThrowsError() {
  // Red: 無効な日付でエラーが発生することを期待
  try {
    formatDate(null);
    assert(false, '無効な日付でエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.indexOf('Invalid date parameter') !== -1, 'エラーメッセージが正しいこと');
  }
}

function testFormatDate_NonDateObject_ThrowsError() {
  // Red: Date以外のオブジェクトでエラーが発生することを期待
  try {
    formatDate('2025-07-10');
    assert(false, '文字列でエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.indexOf('Invalid date parameter') !== -1, 'エラーメッセージが正しいこと');
  }
}

/**
 * formatTime関数のテスト
 */
function testFormatTime_ValidDateTime_ReturnsFormattedTime() {
  // Red: 有効な日時から時刻を正しいフォーマットで返すことを期待
  var testDate = new Date(2025, 6, 10, 14, 30, 0); // 2025年7月10日 14:30:00
  var result = formatTime(testDate);
  assertEquals('14:30', result, 'formatTime は正しい時刻フォーマットを返すべき');
}

function testFormatTime_MidnightTime_ReturnsCorrectFormat() {
  // Red: 真夜中（00:00）が正しくフォーマットされることを期待
  var testDate = new Date(2025, 6, 10, 0, 0, 0);
  var result = formatTime(testDate);
  assertEquals('00:00', result, 'formatTime は真夜中を正しくフォーマットすべき');
}

/**
 * formatDateTime関数のテスト
 */
function testFormatDateTime_ValidDateTime_ReturnsFormattedDateTime() {
  // Red: 有効な日時を正しいフォーマットで返すことを期待
  var testDate = new Date(2025, 6, 10, 14, 30, 0);
  var result = formatDateTime(testDate);
  assertEquals('2025-07-10 14:30', result, 'formatDateTime は正しい日時フォーマットを返すべき');
}

/**
 * parseDate関数のテスト
 */
function testParseDate_ValidDateString_ReturnsDateObject() {
  // Red: 正しい日付文字列からDateオブジェクトを返すことを期待
  var result = parseDate('2025-07-10');
  assertTrue(result instanceof Date, 'parseDate はDateオブジェクトを返すべき');
  assertEquals(2025, result.getFullYear(), '年が正しいこと');
  assertEquals(6, result.getMonth(), '月が正しいこと（0から始まる）');
  assertEquals(10, result.getDate(), '日が正しいこと');
}

function testParseDate_InvalidFormat_ThrowsError() {
  // Red: 無効なフォーマットでエラーが発生することを期待
  try {
    parseDate('2025/07/10');
    assert(false, '無効なフォーマットでエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.indexOf('YYYY-MM-DD format') !== -1, 'エラーメッセージが正しいこと');
  }
}

function testParseDate_InvalidDate_ThrowsError() {
  // Red: 存在しない日付でエラーが発生することを期待
  try {
    parseDate('2025-02-30'); // 2月30日は存在しない
    assert(false, '存在しない日付でエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.indexOf('Invalid date') !== -1, 'エラーメッセージが正しいこと');
  }
}

/**
 * timeStringToMinutes関数のテスト
 */
function testTimeStringToMinutes_ValidTime_ReturnsMinutes() {
  // Red: 正しい時刻文字列から分数を返すことを期待
  var result = timeStringToMinutes('14:30');
  assertEquals(870, result, '14:30は870分であるべき'); // 14*60 + 30 = 870
}

function testTimeStringToMinutes_Midnight_ReturnsZero() {
  // Red: 真夜中（00:00）で0を返すことを期待
  var result = timeStringToMinutes('00:00');
  assertEquals(0, result, '00:00は0分であるべき');
}

function testTimeStringToMinutes_InvalidFormat_ThrowsError() {
  // Red: 無効なフォーマットでエラーが発生することを期待
  try {
    timeStringToMinutes('14-30');
    assert(false, '無効な時刻フォーマットでエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.indexOf('HH:MM format') !== -1, 'エラーメッセージが正しいこと');
  }
}

/**
 * minutesToTimeString関数のテスト
 */
function testMinutesToTimeString_ValidMinutes_ReturnsTimeString() {
  // Red: 有効な分数から時刻文字列を返すことを期待
  var result = minutesToTimeString(870);
  assertEquals('14:30', result, '870分は14:30であるべき');
}

function testMinutesToTimeString_Zero_ReturnsMidnight() {
  // Red: 0分で00:00を返すことを期待
  var result = minutesToTimeString(0);
  assertEquals('00:00', result, '0分は00:00であるべき');
}

function testMinutesToTimeString_NegativeMinutes_ThrowsError() {
  // Red: 負の値でエラーが発生することを期待
  try {
    minutesToTimeString(-10);
    assert(false, '負の分数でエラーが発生すべき');
  } catch (error) {
    assertTrue(error.message.indexOf('non-negative number') !== -1, 'エラーメッセージが正しいこと');
  }
}

/**
 * isSameDate関数のテスト
 */
function testIsSameDate_SameDates_ReturnsTrue() {
  // Red: 同じ日付でtrueを返すことを期待
  var date1 = new Date(2025, 6, 10, 10, 0, 0);
  var date2 = new Date(2025, 6, 10, 15, 30, 0);
  var result = isSameDate(date1, date2);
  assertTrue(result, '同じ日付でtrueを返すべき');
}

function testIsSameDate_DifferentDates_ReturnsFalse() {
  // Red: 異なる日付でfalseを返すことを期待
  var date1 = new Date(2025, 6, 10);
  var date2 = new Date(2025, 6, 11);
  var result = isSameDate(date1, date2);
  assertFalse(result, '異なる日付でfalseを返すべき');
}

function testIsSameDate_NullDates_ReturnsFalse() {
  // Red: null値でfalseを返すことを期待
  var result = isSameDate(null, new Date());
  assertFalse(result, 'null値でfalseを返すべき');
}

/**
 * isEmpty関数のテスト
 */
function testIsEmpty_EmptyString_ReturnsTrue() {
  // Red: 空文字列でtrueを返すことを期待
  var result = isEmpty('');
  assertTrue(result, '空文字列でtrueを返すべき');
}

function testIsEmpty_WhitespaceString_ReturnsTrue() {
  // Red: 空白文字のみでtrueを返すことを期待
  var result = isEmpty('   ');
  assertTrue(result, '空白文字でtrueを返すべき');
}

function testIsEmpty_ValidString_ReturnsFalse() {
  // Red: 有効な文字列でfalseを返すことを期待
  var result = isEmpty('test');
  assertFalse(result, '有効な文字列でfalseを返すべき');
}

function testIsEmpty_Null_ReturnsTrue() {
  // Red: null値でtrueを返すことを期待
  var result = isEmpty(null);
  assertTrue(result, 'null値でtrueを返すべき');
}

/**
 * isValidEmail関数のテスト
 */
function testIsValidEmail_ValidEmail_ReturnsTrue() {
  // Red: 正しいメールアドレスでtrueを返すことを期待
  var result = isValidEmail('test@example.com');
  assertTrue(result, '正しいメールアドレスでtrueを返すべき');
}

function testIsValidEmail_InvalidEmail_ReturnsFalse() {
  // Red: 無効なメールアドレスでfalseを返すことを期待
  var result = isValidEmail('invalid-email');
  assertFalse(result, '無効なメールアドレスでfalseを返すべき');
}

function testIsValidEmail_EmptyString_ReturnsFalse() {
  // Red: 空文字列でfalseを返すことを期待
  var result = isValidEmail('');
  assertFalse(result, '空文字列でfalseを返すべき');
}

/**
 * CommonUtils テスト実行関数
 */
function runCommonUtilsTests() {
  console.log('=== CommonUtils.gs テスト実行開始 ===');
  
  // formatDate関数のテスト
  runTest(testFormatDate_ValidDate_ReturnsFormattedString, 'testFormatDate_ValidDate_ReturnsFormattedString');
  runTest(testFormatDate_InvalidDate_ThrowsError, 'testFormatDate_InvalidDate_ThrowsError');
  runTest(testFormatDate_NonDateObject_ThrowsError, 'testFormatDate_NonDateObject_ThrowsError');
  
  // formatTime関数のテスト
  runTest(testFormatTime_ValidDateTime_ReturnsFormattedTime, 'testFormatTime_ValidDateTime_ReturnsFormattedTime');
  runTest(testFormatTime_MidnightTime_ReturnsCorrectFormat, 'testFormatTime_MidnightTime_ReturnsCorrectFormat');
  
  // formatDateTime関数のテスト
  runTest(testFormatDateTime_ValidDateTime_ReturnsFormattedDateTime, 'testFormatDateTime_ValidDateTime_ReturnsFormattedDateTime');
  
  // parseDate関数のテスト
  runTest(testParseDate_ValidDateString_ReturnsDateObject, 'testParseDate_ValidDateString_ReturnsDateObject');
  runTest(testParseDate_InvalidFormat_ThrowsError, 'testParseDate_InvalidFormat_ThrowsError');
  runTest(testParseDate_InvalidDate_ThrowsError, 'testParseDate_InvalidDate_ThrowsError');
  
  // timeStringToMinutes関数のテスト
  runTest(testTimeStringToMinutes_ValidTime_ReturnsMinutes, 'testTimeStringToMinutes_ValidTime_ReturnsMinutes');
  runTest(testTimeStringToMinutes_Midnight_ReturnsZero, 'testTimeStringToMinutes_Midnight_ReturnsZero');
  runTest(testTimeStringToMinutes_InvalidFormat_ThrowsError, 'testTimeStringToMinutes_InvalidFormat_ThrowsError');
  
  // minutesToTimeString関数のテスト
  runTest(testMinutesToTimeString_ValidMinutes_ReturnsTimeString, 'testMinutesToTimeString_ValidMinutes_ReturnsTimeString');
  runTest(testMinutesToTimeString_Zero_ReturnsMidnight, 'testMinutesToTimeString_Zero_ReturnsMidnight');
  runTest(testMinutesToTimeString_NegativeMinutes_ThrowsError, 'testMinutesToTimeString_NegativeMinutes_ThrowsError');
  
  // isSameDate関数のテスト
  runTest(testIsSameDate_SameDates_ReturnsTrue, 'testIsSameDate_SameDates_ReturnsTrue');
  runTest(testIsSameDate_DifferentDates_ReturnsFalse, 'testIsSameDate_DifferentDates_ReturnsFalse');
  runTest(testIsSameDate_NullDates_ReturnsFalse, 'testIsSameDate_NullDates_ReturnsFalse');
  
  // isEmpty関数のテスト
  runTest(testIsEmpty_EmptyString_ReturnsTrue, 'testIsEmpty_EmptyString_ReturnsTrue');
  runTest(testIsEmpty_WhitespaceString_ReturnsTrue, 'testIsEmpty_WhitespaceString_ReturnsTrue');
  runTest(testIsEmpty_ValidString_ReturnsFalse, 'testIsEmpty_ValidString_ReturnsFalse');
  runTest(testIsEmpty_Null_ReturnsTrue, 'testIsEmpty_Null_ReturnsTrue');
  
  // isValidEmail関数のテスト
  runTest(testIsValidEmail_ValidEmail_ReturnsTrue, 'testIsValidEmail_ValidEmail_ReturnsTrue');
  runTest(testIsValidEmail_InvalidEmail_ReturnsFalse, 'testIsValidEmail_InvalidEmail_ReturnsFalse');
  runTest(testIsValidEmail_EmptyString_ReturnsFalse, 'testIsValidEmail_EmptyString_ReturnsFalse');
  
  console.log('=== CommonUtils.gs テスト実行完了 ===');
} 