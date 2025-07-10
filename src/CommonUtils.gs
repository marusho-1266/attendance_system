/**
 * 全体で使用される共通処理ユーティリティ関数
 * TDD実装による段階的な機能追加
 */

/**
 * 日付をYYYY-MM-DD形式にフォーマット
 * @param {Date} date - フォーマットする日付
 * @return {string} YYYY-MM-DD形式の文字列
 */
function formatDate(date) {
  if (!date || !(date instanceof Date)) {
    throw new Error('Invalid date parameter');
  }
  
  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);
  
  return year + '-' + month + '-' + day;
}

/**
 * 時刻をHH:MM形式にフォーマット
 * @param {Date} date - フォーマットする日付時刻
 * @return {string} HH:MM形式の文字列
 */
function formatTime(date) {
  if (!date || !(date instanceof Date)) {
    throw new Error('Invalid date parameter');
  }
  
  var hours = ('0' + date.getHours()).slice(-2);
  var minutes = ('0' + date.getMinutes()).slice(-2);
  
  return hours + ':' + minutes;
}

/**
 * 日付時刻をYYYY-MM-DD HH:MM形式にフォーマット
 * @param {Date} date - フォーマットする日付時刻
 * @return {string} YYYY-MM-DD HH:MM形式の文字列
 */
function formatDateTime(date) {
  if (!date || !(date instanceof Date)) {
    throw new Error('Invalid date parameter');
  }
  
  return formatDate(date) + ' ' + formatTime(date);
}

/**
 * 文字列から日付オブジェクトを作成
 * @param {string} dateString - YYYY-MM-DD形式の日付文字列
 * @return {Date} 日付オブジェクト
 */
function parseDate(dateString) {
  if (!dateString || typeof dateString !== 'string') {
    throw new Error('Invalid date string parameter');
  }
  
  // YYYY-MM-DD形式の検証
  var datePattern = /^\d{4}-\d{2}-\d{2}$/;
  if (!datePattern.test(dateString)) {
    throw new Error('Date string must be in YYYY-MM-DD format');
  }
  
  var parts = dateString.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1; // Dateオブジェクトは0から始まる
  var day = parseInt(parts[2], 10);
  
  var date = new Date(year, month, day);
  
  // 日付の妥当性確認
  if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) {
    throw new Error('Invalid date: ' + dateString);
  }
  
  return date;
}

/**
 * 時間文字列から分数を計算
 * @param {string} timeString - HH:MM形式の時間文字列
 * @return {number} 分数
 */
function timeStringToMinutes(timeString) {
  if (!timeString || typeof timeString !== 'string') {
    throw new Error('Invalid time string parameter');
  }
  
  // HH:MM形式の検証
  var timePattern = /^\d{1,2}:\d{2}$/;
  if (!timePattern.test(timeString)) {
    throw new Error('Time string must be in HH:MM format');
  }
  
  var parts = timeString.split(':');
  var hours = parseInt(parts[0], 10);
  var minutes = parseInt(parts[1], 10);
  
  // 時間と分の範囲確認
  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
    throw new Error('Invalid time: ' + timeString);
  }
  
  return hours * 60 + minutes;
}

/**
 * 分数から時間文字列を作成
 * @param {number} minutes - 分数
 * @return {string} HH:MM形式の時間文字列
 */
function minutesToTimeString(minutes) {
  if (typeof minutes !== 'number' || minutes < 0) {
    throw new Error('Minutes must be a non-negative number');
  }
  
  var hours = Math.floor(minutes / 60);
  var mins = minutes % 60;
  
  return ('0' + hours).slice(-2) + ':' + ('0' + mins).slice(-2);
}

/**
 * 現在の日時を取得（テスト可能性のため関数化）
 * @return {Date} 現在の日時
 */
function getCurrentDateTime() {
  return new Date();
}

/**
 * 今日の日付を取得（テスト可能性のため関数化）
 * @return {Date} 今日の日付（時刻は00:00:00）
 */
function getToday() {
  var now = getCurrentDateTime();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

/**
 * 2つの日付が同じ日かどうかを判定
 * @param {Date} date1 - 比較する日付1
 * @param {Date} date2 - 比較する日付2
 * @return {boolean} 同じ日の場合true
 */
function isSameDate(date1, date2) {
  if (!date1 || !date2 || !(date1 instanceof Date) || !(date2 instanceof Date)) {
    return false;
  }
  
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

/**
 * 文字列が空または未定義かどうかを判定
 * @param {string} str - 判定する文字列
 * @return {boolean} 空または未定義の場合true
 */
function isEmpty(str) {
  return !str || str.trim() === '';
}

/**
 * メールアドレスの形式を検証
 * @param {string} email - 検証するメールアドレス
 * @return {boolean} 正しい形式の場合true
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }
  
  // 基本的なメールアドレス形式の検証
  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
} 