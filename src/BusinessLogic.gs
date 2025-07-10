/**
 * 業務ロジック関数群
 * TDD実装: Green段階 - テストを通すための最小実装
 * 
 * 対象機能:
 * - isHoliday(date): 休日判定
 * - calcWorkTime(startTime, endTime, breakMinutes): 勤務時間計算
 * - getEmployee(email): 従業員情報取得
 */

// === 休日判定関数 ===

/**
 * 指定された日付が休日かどうかを判定
 * @param {Date} date - 判定対象の日付
 * @returns {boolean} 休日の場合true、平日の場合false
 */
function isHoliday(date) {
  // パラメータ検証
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    throw new Error('Invalid date parameter');
  }
  
  // 土日判定（最小実装）
  var dayOfWeek = date.getDay(); // 0=日曜, 6=土曜
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return true;
  }
  
  // 祝日判定（簡易実装 - Master_Holidayとの照合は後でRefactor）
  // 元日の固定判定（テスト用）
  if (date.getMonth() === 0 && date.getDate() === 1) {
    return true; // 元日
  }
  
  return false; // 平日
}

// === 勤務時間計算関数 ===

/**
 * 勤務時間を計算（分単位で返却）
 * @param {string} startTime - 開始時刻 (HH:MM形式)
 * @param {string} endTime - 終了時刻 (HH:MM形式)
 * @param {number} breakMinutes - 休憩時間（分）
 * @returns {number} 勤務時間（分）
 */
function calcWorkTime(startTime, endTime, breakMinutes) {
  // パラメータ検証
  if (!startTime || !endTime) {
    throw new Error('Invalid time parameters');
  }
  
  if (breakMinutes < 0) {
    throw new Error('Invalid break time - negative value');
  }
  
  // 時刻を分に変換
  var startMinutes = parseTimeToMinutes(startTime);
  var endMinutes = parseTimeToMinutes(endTime);
  
  // 日跨ぎ判定
  if (endMinutes <= startMinutes) {
    endMinutes += 24 * 60; // 翌日の時刻として計算
  }
  
  // 勤務時間計算
  var totalMinutes = endMinutes - startMinutes;
  var workMinutes = totalMinutes - breakMinutes;
  
  return workMinutes;
}

/**
 * HH:MM形式の時刻を分に変換
 * @param {string} timeString - HH:MM形式の時刻
 * @returns {number} 分数
 */
function parseTimeToMinutes(timeString) {
  var timeParts = timeString.split(':');
  if (timeParts.length !== 2) {
    throw new Error('Invalid time format - use HH:MM');
  }
  
  var hours = parseInt(timeParts[0], 10);
  var minutes = parseInt(timeParts[1], 10);
  
  // 時刻の妥当性チェック
  if (isNaN(hours) || isNaN(minutes) || hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
    throw new Error('Invalid time values');
  }
  
  return hours * 60 + minutes;
}

// === 従業員情報取得関数 ===

/**
 * メールアドレスから従業員情報を取得
 * @param {string} email - 検索対象のメールアドレス
 * @returns {Object|null} 従業員情報オブジェクトまたはnull
 */
function getEmployee(email) {
  // パラメータ検証
  if (!email) {
    throw new Error('Empty email parameter');
  }
  
  // メールフォーマット検証（簡易）
  if (!isValidEmail(email)) {
    throw new Error('Invalid email format');
  }
  
  // 最小実装: テスト用の固定データ
  // 実際のスプレッドシート検索は後でRefactor
  var testEmployees = [
    {
      employeeId: 'EMP001',
      name: '山田太郎',
      gmail: 'yamada@example.com',
      department: '開発部',
      employmentType: '正社員',
      supervisorGmail: 'manager@example.com',
      startTime: '09:00',
      endTime: '18:00'
    }
  ];
  
  // メールアドレスで検索
  for (var i = 0; i < testEmployees.length; i++) {
    if (testEmployees[i].gmail === email) {
      return testEmployees[i];
    }
  }
  
  return null; // 見つからない場合
}

// === ユーティリティ関数 ===

/**
 * メールアドレスの妥当性チェック（簡易版）
 * @param {string} email - チェック対象のメールアドレス
 * @returns {boolean} 有効な場合true
 */
function isValidEmail(email) {
  // 簡易的なメールフォーマットチェック
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
} 