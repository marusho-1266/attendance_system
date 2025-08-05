/**
 * 出勤管理システム - 設定ファイル
 * 基本的な定数とグローバル設定を定義
 */

// ===== システム基本設定 =====
const SYSTEM_CONFIG = {
  // システム名
  SYSTEM_NAME: '出勤管理システム',
  
  // バージョン
  VERSION: '1.0.0',
  
  // タイムゾーン
  TIMEZONE: 'Asia/Tokyo',
  
  // 管理者メールアドレス（実際の運用時に変更）
  ADMIN_EMAIL: 'admin@example.com',
  
  // 組織名（実際の運用時に変更）
  ORGANIZATION_NAME: 'サンプル組織',
  
  // 時刻修正申請フォームURL（実際の運用時に設定）
  TIME_ADJUSTMENT_FORM_URL: 'https://forms.google.com/your-form-url'
};

// ===== スプレッドシート設定 =====
const SHEET_CONFIG = {
  // シート名定数
  SHEETS: {
    MASTER_EMPLOYEE: 'Master_Employee',
    MASTER_HOLIDAY: 'Master_Holiday',
    LOG_RAW: 'Log_Raw',
    DAILY_SUMMARY: 'Daily_Summary',
    MONTHLY_SUMMARY: 'Monthly_Summary',
    REQUEST_RESPONSES: 'Request_Responses'
  },
  
  // スプレッドシートID（実際の運用時に設定）
  SPREADSHEET_ID: '',
  
  // バックアップフォルダID（実際の運用時に設定）
  BACKUP_FOLDER_ID: ''
};

// ===== 業務ルール設定 =====
const BUSINESS_RULES = {
  // 標準労働時間（分）
  STANDARD_WORK_MINUTES: 480, // 8時間
  
  // 休憩時間設定
  BREAK_RULES: {
    // 6時間以上勤務時の休憩時間（分）
    BREAK_MINUTES_FOR_6H: 45,
    // 休憩適用の最小勤務時間（分）
    MIN_WORK_FOR_BREAK: 360 // 6時間
  },
  
  // 残業警告閾値
  OVERTIME_WARNING_THRESHOLD: 4800, // 80時間（4週間、分単位）
  OVERTIME_WARNING_THRESHOLD_HOURS: 80, // 80時間（4週間、時間単位）
  
  // 深夜勤務時間帯
  NIGHT_WORK: {
    START_HOUR: 22, // 22:00
    END_HOUR: 5     // 5:00
  },
  
  // 打刻アクション
  TIME_ACTIONS: {
    CLOCK_IN: 'IN',
    CLOCK_OUT: 'OUT',
    BREAK_IN: 'BRK_IN',
    BREAK_OUT: 'BRK_OUT'
  },
  
  // 申請ステータス
  REQUEST_STATUS: {
    PENDING: 'Pending',
    APPROVED: 'Approved',
    REJECTED: 'Rejected'
  },
  
  // 申請種別
  REQUEST_TYPES: {
    TIME_CORRECTION: '修正',
    OVERTIME: '残業',
    VACATION: '休暇'
  }
};

// ===== 自動処理設定 =====
const AUTOMATION_CONFIG = {
  // 日次ジョブ実行時刻
  DAILY_JOB_HOUR: 2,
  DAILY_JOB_MINUTE: 0,
  
  // 週次ジョブ実行時刻（月曜日）
  WEEKLY_JOB_DAY: 2, // 月曜日（1=日曜日）
  WEEKLY_JOB_HOUR: 7,
  WEEKLY_JOB_MINUTE: 0,
  
  // 月次ジョブ実行時刻（1日）
  MONTHLY_JOB_DAY: 1,
  MONTHLY_JOB_HOUR: 2,
  MONTHLY_JOB_MINUTE: 30,
  
  // バッチ処理設定
  BATCH_SIZE: 100,
  MAX_EXECUTION_TIME: 270000 // 4.5分（ミリ秒）
};

// ===== メール設定 =====
const MAIL_CONFIG = {
  // メール送信制限（1日あたり）
  DAILY_MAIL_QUOTA: 100,
  
  // メールテンプレート
  TEMPLATES: {
    ERROR_ALERT: {
      SUBJECT: '[緊急] 出勤管理システムエラー',
      BODY_PREFIX: 'システムでエラーが発生しました。\n\n'
    },
    APPROVAL_REQUEST: {
      SUBJECT: '【承認依頼】出勤申請の確認をお願いします',
      BODY_PREFIX: '以下の申請について承認をお願いします。\n\n'
    },
    OVERTIME_WARNING: {
      SUBJECT: '【警告】残業時間超過のお知らせ',
      BODY_PREFIX: '残業時間が規定値を超過しています。\n\n'
    },
    MISSING_CLOCKOUT: {
      SUBJECT: '【通知】退勤打刻漏れのお知らせ',
      BODY_PREFIX: '以下の従業員の退勤打刻が確認できません。\n\n'
    }
  }
};

// ===== エラーハンドリング設定 =====
// ERROR_CONFIGはErrorHandler.gsで定義されています

// ===== セキュリティ設定 =====
const SECURITY_CONFIG = {
  // セッションタイムアウト（分）
  SESSION_TIMEOUT: 480, // 8時間
  
  // IPアドレス記録フラグ
  RECORD_IP_ADDRESS: true,
  
  // アクセスログ保持期間（日）
  ACCESS_LOG_RETENTION_DAYS: 90,
  
  // X-Frame-Options設定
  // DENY: iframe埋め込みを完全に拒否（最も安全）
  // SAMEORIGIN: 同じオリジンのみ許可
  // ALLOWALL: すべてのドメインを許可（危険）
  X_FRAME_OPTIONS_MODE: 'DENY'
};

/**
 * 設定値を取得する関数
 * @param {string} category - 設定カテゴリ
 * @param {string} key - 設定キー
 * @return {any} 設定値
 */
function getConfig(category, key) {
  const configs = {
    'SYSTEM': SYSTEM_CONFIG,
    'SHEET': SHEET_CONFIG,
    'BUSINESS': BUSINESS_RULES,
    'AUTOMATION': AUTOMATION_CONFIG,
    'MAIL': MAIL_CONFIG,
    'ERROR': typeof ERROR_CONFIG !== 'undefined' ? ERROR_CONFIG : null,
    'SECURITY': SECURITY_CONFIG
  };
  
  if (!configs[category]) {
    throw new Error(`設定カテゴリが見つかりません: ${category}`);
  }
  
  if (key && !configs[category][key]) {
    throw new Error(`設定キーが見つかりません: ${category}.${key}`);
  }
  
  return key ? configs[category][key] : configs[category];
}

/**
 * スプレッドシートIDを設定する関数
 * @param {string} spreadsheetId - スプレッドシートID
 */
function setSpreadsheetId(spreadsheetId) {
  SHEET_CONFIG.SPREADSHEET_ID = spreadsheetId;
}

/**
 * 管理者メールアドレスを設定する関数
 * @param {string} adminEmail - 管理者メールアドレス
 */
function setAdminEmail(adminEmail) {
  SYSTEM_CONFIG.ADMIN_EMAIL = adminEmail;
}

/**
 * 現在の設定を表示する関数（デバッグ用）
 */
function showCurrentConfig() {
  console.log('=== 現在の設定 ===');
  console.log('システム設定:', SYSTEM_CONFIG);
  console.log('シート設定:', SHEET_CONFIG);
  console.log('業務ルール:', BUSINESS_RULES);
  console.log('自動処理設定:', AUTOMATION_CONFIG);
  console.log('メール設定:', MAIL_CONFIG);
  console.log('エラー設定:', typeof ERROR_CONFIG !== 'undefined' ? ERROR_CONFIG : 'ErrorHandler.gsで定義');
  console.log('セキュリティ設定:', SECURITY_CONFIG);
}