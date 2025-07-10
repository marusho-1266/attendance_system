/**
 * アプリケーション全体で使用される定数定義
 * TDD実装での段階的な定数追加
 */

// === スプレッドシート列インデックス定数 ===

/**
 * Master_Employee シートの列定義
 */
var EMPLOYEE_COLUMNS = {
  EMPLOYEE_ID: 0,    // A列: 社員ID
  NAME: 1,           // B列: 氏名
  GMAIL: 2,          // C列: Gmail
  DEPARTMENT: 3,     // D列: 所属
  EMPLOYMENT_TYPE: 4, // E列: 雇用区分
  SUPERVISOR_GMAIL: 5, // F列: 上司Gmail
  START_TIME: 6,     // G列: 基準始業
  END_TIME: 7        // H列: 基準終業
};

/**
 * Master_Holiday シートの列定義
 */
var HOLIDAY_COLUMNS = {
  DATE: 0,           // A列: 日付
  NAME: 1,           // B列: 名称
  IS_LEGAL: 2,       // C列: 法定休日?
  IS_COMPANY: 3      // D列: 会社休日?
};

/**
 * Log_Raw シートの列定義
 */
var LOG_RAW_COLUMNS = {
  TIMESTAMP: 0,      // A列: タイムスタンプ
  EMPLOYEE_ID: 1,    // B列: 社員ID
  NAME: 2,           // C列: 氏名
  ACTION: 3,         // D列: Action
  IP_ADDRESS: 4,     // E列: 端末IP
  REMARKS: 5         // F列: 備考
};

/**
 * Daily_Summary シートの列定義
 */
var DAILY_SUMMARY_COLUMNS = {
  DATE: 0,           // A列: 日付
  EMPLOYEE_ID: 1,    // B列: 社員ID
  START_TIME: 2,     // C列: 出勤
  END_TIME: 3,       // D列: 退勤
  BREAK_TIME: 4,     // E列: 休憩
  WORK_TIME: 5,      // F列: 実働
  OVERTIME: 6,       // G列: 残業
  LATE_EARLY: 7,     // H列: 遅刻/早退
  APPROVAL_STATUS: 8 // I列: 承認状態
};

/**
 * Monthly_Summary シートの列定義
 */
var MONTHLY_SUMMARY_COLUMNS = {
  YEAR_MONTH: 0,     // A列: 年月
  EMPLOYEE_ID: 1,    // B列: 社員ID
  WORK_DAYS: 2,      // C列: 勤務日数
  TOTAL_WORK: 3,     // D列: 総労働
  OVERTIME: 4,       // E列: 残業
  PAID_LEAVE: 5,     // F列: 有給
  REMARKS: 6         // G列: 備考
};

/**
 * Request_Responses シートの列定義
 */
var REQUEST_RESPONSE_COLUMNS = {
  TIMESTAMP: 0,      // A列: タイムスタンプ
  EMPLOYEE_ID: 1,    // B列: 社員ID
  TYPE: 2,           // C列: 種別(修正/残業/休暇)
  DETAILS: 3,        // D列: 詳細
  REQUESTED_VALUE: 4, // E列: 希望値
  APPROVER: 5,       // F列: 承認者
  STATUS: 6          // G列: Status
};

// === アクション定数 ===

/**
 * 打刻アクション定数
 */
var ACTIONS = {
  CLOCK_IN: 'IN',           // 出勤
  CLOCK_OUT: 'OUT',         // 退勤
  BREAK_START: 'BRK_IN',    // 休憩開始
  BREAK_END: 'BRK_OUT'      // 休憩終了
};

// === 申請ステータス定数 ===

/**
 * 申請承認ステータス
 */
var APPROVAL_STATUS = {
  PENDING: 'Pending',       // 承認待ち
  APPROVED: 'Approved',     // 承認済み
  REJECTED: 'Rejected'      // 却下
};

// === シート名定数 ===

/**
 * スプレッドシート内のシート名
 */
var SHEET_NAMES = {
  MASTER_EMPLOYEE: 'Master_Employee',
  MASTER_HOLIDAY: 'Master_Holiday',
  LOG_RAW: 'Log_Raw',
  DAILY_SUMMARY: 'Daily_Summary',
  MONTHLY_SUMMARY: 'Monthly_Summary',
  REQUEST_RESPONSES: 'Request_Responses'
};

// === アプリケーション設定定数 ===

/**
 * システム設定
 */
var APP_CONFIG = {
  MAX_WORK_HOURS_PER_DAY: 24,     // 1日の最大労働時間
  STANDARD_WORK_HOURS: 8,         // 標準労働時間
  BREAK_TIME_AUTO_DEDUCT: 45,     // 自動控除する休憩時間（分）
  OVERTIME_THRESHOLD: 8           // 残業判定の閾値（時間）
};

/**
 * 列インデックス取得関数（TDD対象）
 */
function getColumnIndex(sheetType, columnName) {
  switch (sheetType) {
    case 'EMPLOYEE':
      return EMPLOYEE_COLUMNS[columnName];
    case 'HOLIDAY':
      return HOLIDAY_COLUMNS[columnName];
    case 'LOG_RAW':
      return LOG_RAW_COLUMNS[columnName];
    case 'DAILY_SUMMARY':
      return DAILY_SUMMARY_COLUMNS[columnName];
    case 'MONTHLY_SUMMARY':
      return MONTHLY_SUMMARY_COLUMNS[columnName];
    case 'REQUEST_RESPONSES':
      return REQUEST_RESPONSE_COLUMNS[columnName];
    default:
      throw new Error('Unknown sheet type: ' + sheetType);
  }
}

/**
 * シート名取得関数（TDD対象）
 */
function getSheetName(sheetType) {
  if (!SHEET_NAMES.hasOwnProperty(sheetType)) {
    throw new Error('Unknown sheet type: ' + sheetType);
  }
  return SHEET_NAMES[sheetType];
}
/**
 * アクション定数取得関数（TDD対象）
 */
function getActionConstant(actionType) {
  if (!ACTIONS.hasOwnProperty(actionType)) {
    throw new Error('Unknown action type: ' + actionType);
  }
  return ACTIONS[actionType];
}

/**
function getAppConfig(configKey) {
  if (!APP_CONFIG.hasOwnProperty(configKey)) {
    throw new Error('Unknown config key: ' + configKey);
  }
  return APP_CONFIG[configKey];
}  return APP_CONFIG[configKey];
} 