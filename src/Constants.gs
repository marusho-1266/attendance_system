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

/**
 * System_Config シートの列定義
 */
var SYSTEM_CONFIG_COLUMNS = {
  CONFIG_KEY: 0,     // A列: 設定キー
  CONFIG_VALUE: 1,   // B列: 設定値
  DESCRIPTION: 2,    // C列: 説明
  IS_ACTIVE: 3       // D列: 有効フラグ
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

/**
 * 権限アクション定数（認証機能用）
 */
var PERMISSION_ACTIONS = {
  CLOCK_IN: 'CLOCK_IN',         // 打刻権限
  CLOCK_OUT: 'CLOCK_OUT',       // 退勤権限
  BREAK_START: 'BREAK_START',   // 休憩開始権限
  BREAK_END: 'BREAK_END',       // 休憩終了権限
  ADMIN_ACCESS: 'ADMIN_ACCESS', // 管理者権限
  VIEW_REPORTS: 'VIEW_REPORTS', // レポート閲覧権限
  EDIT_MASTER: 'EDIT_MASTER'    // マスタ編集権限
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
  REQUEST_RESPONSES: 'Request_Responses',
  SYSTEM_CONFIG: 'System_Config'
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
    case 'SYSTEM_CONFIG':
      return SYSTEM_CONFIG_COLUMNS[columnName];
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
 * メール送信モード設定
 * 相互排他的な設定で、本番環境での誤用を防止
 */
var EMAIL_MODE_CONFIG = {
  // メール送信モード（'MOCK' | 'ACTUAL'）
  // MOCK: テスト環境用（実際のメールは送信されない）
  // ACTUAL: 本番環境用（実際のメールが送信される）
  EMAIL_SEND_MODE: 'MOCK',
  
  // モックメール情報の保存先（テスト環境のみ使用）
  EMAIL_MOCK_STORAGE: {},
  
  // メール送信クォータ設定
  EMAIL_DAILY_QUOTA: 100,      // 1日の最大送信数
  EMAIL_QUOTA_RESET_HOUR: 0    // クォータリセット時刻（0-23時）
};

/**
 * モックメール情報をクリア
 */
function clearMockEmailData() {
  EMAIL_MODE_CONFIG.EMAIL_MOCK_STORAGE = {};
}

/**
 * モックメール情報を取得
 */
function getMockEmailData() {
  return EMAIL_MODE_CONFIG.EMAIL_MOCK_STORAGE;
}

/**
 * メール送信モード取得関数
 */
function getEmailMode() {
  return EMAIL_MODE_CONFIG.EMAIL_SEND_MODE;
}

/**
 * メール送信モード設定関数
 */
function setEmailMode(mode) {
  if (mode !== 'MOCK' && mode !== 'ACTUAL') {
    throw new Error('Invalid email mode: ' + mode + '. Must be "MOCK" or "ACTUAL"');
  }
  EMAIL_MODE_CONFIG.EMAIL_SEND_MODE = mode;
}

/**
 * メール送信モードがモックかどうかを判定
 */
function isEmailMockMode() {
  return EMAIL_MODE_CONFIG.EMAIL_SEND_MODE === 'MOCK';
}

/**
 * メール送信モードが実際の送信かどうかを判定
 */
function isEmailActualMode() {
  return EMAIL_MODE_CONFIG.EMAIL_SEND_MODE === 'ACTUAL';
}

/**
 * メール設定取得関数
 */
function getEmailConfig(configKey) {
  if (!EMAIL_MODE_CONFIG.hasOwnProperty(configKey)) {
    throw new Error('Unknown email config key: ' + configKey);
  }
  return EMAIL_MODE_CONFIG[configKey];
}

/**
 * メールテンプレート定数
 */
var EMAIL_TEMPLATES = {
  // 未退勤者メールテンプレート
  UNFINISHED_CLOCKOUT: {
    SUBJECT_PREFIX: '【重要】未退勤通知',
    SYSTEM_NAME: '出勤管理システム',
    AUTO_SEND_NOTICE: '※このメールは自動送信です。',
    FORM_INSTRUCTION: '退勤打刻を忘れた場合は、修正申請フォームより報告してください。'
  },
  
  // 月次レポートメールテンプレート
  MONTHLY_REPORT: {
    SUBJECT_PREFIX: '【月次レポート】',
    SUBJECT_SUFFIX: ' 出勤管理',
    SYSTEM_NAME: '出勤管理システム',
    REPORT_TITLE: '出勤管理システム 月次レポート'
  },
  
  // 共通テンプレート
  COMMON: {
    SYSTEM_SIGNATURE: '出勤管理システム',
    AUTO_SEND_NOTICE: '※このメールは自動送信です。返信は不要です。',
    LINE_BREAK: '\n',
    DOUBLE_LINE_BREAK: '\n\n'
  }
};

/**
 * メールテンプレート取得関数
 */
function getEmailTemplate(templateType, templateKey) {
  if (!EMAIL_TEMPLATES.hasOwnProperty(templateType)) {
    throw new Error('Unknown email template type: ' + templateType);
  }
  
  var template = EMAIL_TEMPLATES[templateType];
  if (templateKey && !template.hasOwnProperty(templateKey)) {
    throw new Error('Unknown template key: ' + templateKey + ' in template type: ' + templateType);
  }
  
  return templateKey ? template[templateKey] : template;
}

/**
 * パラメータ付きメールテンプレート定義
 */
var EMAIL_TEMPLATE_DEFS = {
  UNFINISHED_CLOCKOUT: {
    subject: '【重要】未退勤通知 ({date})',
    body: '{systemName}より自動送信\n\n{name} 様\n\n{date} の退勤打刻が記録されていません。\n出勤時刻: {clockInTime}\n現在時刻: {currentTime}\n\n{formInstruction}\n\n{autoSendNotice}\n{systemSignature}'
  },
  MONTHLY_REPORT: {
    subject: '【月次レポート】{month} 出勤管理',
    body: '{reportTitle}\n\n対象月: {month}\n総従業員数: {totalEmployees}名\n総稼働日数: {totalWorkDays}日\n平均労働時間: {averageWorkHours}時間\n総残業時間: {overtimeHours}時間\n\nCSVファイル: {csvFileName}\n\n詳細については添付のCSVファイルをご確認ください。\n\n{systemSignature}'
  }
};

/**
 * テンプレート文字列をパラメータで置換する共通関数
 */
function buildEmailTemplate(type, params) {
  var def = EMAIL_TEMPLATE_DEFS[type];
  if (!def) throw new Error('Unknown email template type: ' + type);
  var subject = def.subject;
  var body = def.body;
  Object.keys(params).forEach(function(key) {
    var re = new RegExp('{' + key + '}', 'g');
    subject = subject.replace(re, params[key]);
    body = body.replace(re, params[key]);
  });
  return { subject: subject, body: body };
}

/**
 * アプリケーション設定取得関数（TDD対象）
 */
function getAppConfig(configKey) {
  if (!APP_CONFIG.hasOwnProperty(configKey)) {
    throw new Error('Unknown config key: ' + configKey);
  }
  return APP_CONFIG[configKey];
}

/**
 * 権限アクション定数取得関数（TDD対象）
 */
function getPermissionAction(actionType) {
  if (!PERMISSION_ACTIONS.hasOwnProperty(actionType)) {
    throw new Error('Unknown permission action type: ' + actionType);
  }
  return PERMISSION_ACTIONS[actionType];
}