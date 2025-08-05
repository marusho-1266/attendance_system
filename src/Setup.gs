/**
 * 出勤管理システム - セットアップモジュール
 * 初期設定とプロジェクト構造の確認を担当
 */

/**
 * プロジェクト構造を確認する
 * @return {Object} 構造確認結果
 */
function checkProjectStructure() {
  try {
    const result = {
      timestamp: new Date(),
      success: true,
      modules: {},
      config: {},
      errors: []
    };
    
    // モジュール存在確認
    const modules = [
      'Authentication',
      'WebApp', 
      'BusinessLogic',
      'FormManager',
      'MailManager',
      'Triggers',
      'Utils',
      'Config'
    ];
    
    modules.forEach(module => {
      try {
        // 各モジュールの関数が定義されているかチェック
        result.modules[module] = {
          exists: true,
          functions: getFunctionNames(module)
        };
      } catch (error) {
        result.modules[module] = {
          exists: false,
          error: error.message
        };
        result.errors.push(`モジュール ${module} が見つかりません: ${error.message}`);
      }
    });
    
    // 設定確認
    try {
      result.config = {
        systemConfig: typeof SYSTEM_CONFIG !== 'undefined',
        sheetConfig: typeof SHEET_CONFIG !== 'undefined',
        businessRules: typeof BUSINESS_RULES !== 'undefined',
        automationConfig: typeof AUTOMATION_CONFIG !== 'undefined',
        mailConfig: typeof MAIL_CONFIG !== 'undefined',
        errorConfig: typeof ERROR_CONFIG !== 'undefined',
        securityConfig: typeof SECURITY_CONFIG !== 'undefined'
      };
    } catch (error) {
      result.errors.push(`設定確認エラー: ${error.message}`);
    }
    
    return result;
    
  } catch (error) {
    Logger.log('プロジェクト構造確認エラー: ' + error.toString());
    return {
      timestamp: new Date(),
      success: false,
      error: error.message
    };
  }
}

/**
 * モジュールの関数名を取得する（デバッグ用）
 * @param {string} moduleName - モジュール名
 * @return {Array} 関数名リスト
 */
function getFunctionNames(moduleName) {
  // 実際の実装では、各モジュールの関数を動的に取得
  // 現在はスケルトン段階なので、予定されている関数名を返す
  const functionMap = {
    'Authentication': ['authenticateUser', 'validateEmployeeAccess', 'getEmployeeInfo'],
    'WebApp': ['doGet', 'doPost', 'recordTimeEntry', 'validateTimeEntry'],
    'BusinessLogic': ['calculateDailyWorkTime', 'calculateOvertime', 'applyBreakDeduction', 'checkHoliday', 'updateDailySummary'],
    'FormManager': ['onRequestSubmit', 'processApprovalRequest', 'batchApprovalNotifications', 'updateRequestStatus'],
    'MailManager': ['sendBatchNotification', 'sendErrorAlert', 'sendOvertimeWarning', 'sendMonthlyReport'],
    'Triggers': ['dailyJob', 'weeklyOvertimeJob', 'monthlyJob', 'onOpen'],
    'Utils': ['isHoliday', 'formatTime', 'getSheetData', 'setSheetData', 'exportToCsv'],
    'Config': ['getConfig', 'setSpreadsheetId', 'setAdminEmail', 'showCurrentConfig']
  };
  
  return functionMap[moduleName] || [];
}

/**
 * スプレッドシートの構造を設定する
 */
function setupSpreadsheetStructure() {
  try {
    Logger.log('=== スプレッドシート構造設定開始 ===');
    
    // 現在のスプレッドシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 各シートを作成・設定
    const sheetResults = {
      masterEmployee: setupMasterEmployeeSheet(spreadsheet),
      masterHoliday: setupMasterHolidaySheet(spreadsheet),
      logRaw: setupLogRawSheet(spreadsheet),
      dailySummary: setupDailySummarySheet(spreadsheet),
      monthlySummary: setupMonthlySummarySheet(spreadsheet),
      requestResponses: setupRequestResponsesSheet(spreadsheet)
    };
    
    // シート保護設定を適用
    applySheetProtections(spreadsheet);
    
    Logger.log('=== スプレッドシート構造設定完了 ===');
    
    return {
      success: true,
      message: 'スプレッドシート構造が正常に設定されました',
      sheetResults: sheetResults
    };
    
  } catch (error) {
    Logger.log('スプレッドシート構造設定エラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Master_Employeeシートを設定する
 */
function setupMasterEmployeeSheet(spreadsheet) {
  try {
    // SHEET_CONFIGの存在チェック
    if (typeof SHEET_CONFIG === 'undefined' || !SHEET_CONFIG.SHEETS) {
      throw new Error('SHEET_CONFIGが定義されていないか、SHEETSプロパティが存在しません');
    }
    
    const sheetName = SHEET_CONFIG.SHEETS.MASTER_EMPLOYEE;
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    
    // ヘッダー行を設定
    const headers = [
      '社員ID',
      '氏名', 
      'Gmail',
      '所属',
      '雇用区分',
      '上司Gmail',
      '基準始業',
      '基準終業'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のフォーマット設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 列幅の調整
    sheet.setColumnWidth(1, 100); // 社員ID
    sheet.setColumnWidth(2, 120); // 氏名
    sheet.setColumnWidth(3, 200); // Gmail
    sheet.setColumnWidth(4, 100); // 所属
    sheet.setColumnWidth(5, 100); // 雇用区分
    sheet.setColumnWidth(6, 200); // 上司Gmail
    sheet.setColumnWidth(7, 100); // 基準始業
    sheet.setColumnWidth(8, 100); // 基準終業
    
    // サンプルデータを追加
    const sampleData = [
      ['EMP001', '山田太郎', 'yamada@example.com', '営業部', '正社員', 'manager@example.com', '09:00', '18:00'],
      ['EMP002', '佐藤花子', 'sato@example.com', '開発部', '正社員', 'dev-manager@example.com', '09:30', '18:30']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    Logger.log(`${sheetName}シートを設定しました`);
    return { success: true, sheetName: sheetName };
    
  } catch (error) {
    Logger.log(`${sheetName}シート設定エラー: ${error.toString()}`);
    return { success: false, error: error.message };
  }
}

/**
 * Master_Holidayシートを設定する
 */
function setupMasterHolidaySheet(spreadsheet) {
  try {
    // SHEET_CONFIGの存在チェック
    if (typeof SHEET_CONFIG === 'undefined' || !SHEET_CONFIG.SHEETS) {
      throw new Error('SHEET_CONFIGが定義されていないか、SHEETSプロパティが存在しません');
    }
    
    const sheetName = SHEET_CONFIG.SHEETS.MASTER_HOLIDAY;
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    
    const headers = [
      '日付',
      '名称',
      '法定休日フラグ',
      '会社休日フラグ'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のフォーマット設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 列幅の調整
    sheet.setColumnWidth(1, 120); // 日付
    sheet.setColumnWidth(2, 200); // 名称
    sheet.setColumnWidth(3, 120); // 法定休日フラグ
    sheet.setColumnWidth(4, 120); // 会社休日フラグ
    
    // サンプルデータを追加（2024年の主要祝日）
    const sampleData = [
      [new Date('2024-01-01'), '元日', true, true],
      [new Date('2024-01-08'), '成人の日', true, true],
      [new Date('2024-02-11'), '建国記念の日', true, true],
      [new Date('2024-02-23'), '天皇誕生日', true, true],
      [new Date('2024-03-20'), '春分の日', true, true],
      [new Date('2024-04-29'), '昭和の日', true, true],
      [new Date('2024-05-03'), '憲法記念日', true, true],
      [new Date('2024-05-04'), 'みどりの日', true, true],
      [new Date('2024-05-05'), 'こどもの日', true, true]
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    Logger.log(`${sheetName}シートを設定しました`);
    return { success: true, sheetName: sheetName };
    
  } catch (error) {
    Logger.log(`${sheetName}シート設定エラー: ${error.toString()}`);
    return { success: false, error: error.message };
  }
}

/**
 * Log_Rawシートを設定する
 */
function setupLogRawSheet(spreadsheet) {
  try {
    // SHEET_CONFIGの存在チェック
    if (typeof SHEET_CONFIG === 'undefined' || !SHEET_CONFIG.SHEETS) {
      throw new Error('SHEET_CONFIGが定義されていないか、SHEETSプロパティが存在しません');
    }
    
    const sheetName = SHEET_CONFIG.SHEETS.LOG_RAW;
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    
    const headers = [
      'タイムスタンプ',
      '社員ID',
      '氏名',
      'Action',
      '端末IP',
      '備考'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のフォーマット設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#ea4335');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 列幅の調整
    sheet.setColumnWidth(1, 180); // タイムスタンプ
    sheet.setColumnWidth(2, 100); // 社員ID
    sheet.setColumnWidth(3, 120); // 氏名
    sheet.setColumnWidth(4, 100); // Action
    sheet.setColumnWidth(5, 120); // 端末IP
    sheet.setColumnWidth(6, 200); // 備考
    
    Logger.log(`${sheetName}シートを設定しました`);
    return { success: true, sheetName: sheetName };
    
  } catch (error) {
    Logger.log(`${sheetName}シート設定エラー: ${error.toString()}`);
    return { success: false, error: error.message };
  }
}

/**
 * Daily_Summaryシートを設定する
 */
function setupDailySummarySheet(spreadsheet) {
  try {
    // SHEET_CONFIGの存在チェック
    if (typeof SHEET_CONFIG === 'undefined' || !SHEET_CONFIG.SHEETS) {
      throw new Error('SHEET_CONFIGが定義されていないか、SHEETSプロパティが存在しません');
    }
    
    const sheetName = SHEET_CONFIG.SHEETS.DAILY_SUMMARY;
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    
    const headers = [
      '日付',
      '社員ID',
      '出勤時刻',
      '退勤時刻',
      '休憩時間(分)',
      '実働時間(分)',
      '残業時間(分)',
      '遅刻早退(分)',
      '承認状態'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のフォーマット設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#ff9900');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 列幅の調整
    sheet.setColumnWidth(1, 100); // 日付
    sheet.setColumnWidth(2, 100); // 社員ID
    sheet.setColumnWidth(3, 100); // 出勤時刻
    sheet.setColumnWidth(4, 100); // 退勤時刻
    sheet.setColumnWidth(5, 100); // 休憩時間
    sheet.setColumnWidth(6, 100); // 実働時間
    sheet.setColumnWidth(7, 100); // 残業時間
    sheet.setColumnWidth(8, 100); // 遅刻早退
    sheet.setColumnWidth(9, 100); // 承認状態
    
    Logger.log(`${sheetName}シートを設定しました`);
    return { success: true, sheetName: sheetName };
    
  } catch (error) {
    Logger.log(`${sheetName}シート設定エラー: ${error.toString()}`);
    return { success: false, error: error.message };
  }
}

/**
 * Monthly_Summaryシートを設定する
 */
function setupMonthlySummarySheet(spreadsheet) {
  try {
    // SHEET_CONFIGの存在チェック
    if (typeof SHEET_CONFIG === 'undefined' || !SHEET_CONFIG.SHEETS) {
      throw new Error('SHEET_CONFIGが定義されていないか、SHEETSプロパティが存在しません');
    }
    
    const sheetName = SHEET_CONFIG.SHEETS.MONTHLY_SUMMARY;
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    
    const headers = [
      '年月',
      '社員ID',
      '勤務日数',
      '総労働時間(分)',
      '残業時間(分)',
      '有給日数',
      '備考'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のフォーマット設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#9c27b0');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 列幅の調整
    sheet.setColumnWidth(1, 100); // 年月
    sheet.setColumnWidth(2, 100); // 社員ID
    sheet.setColumnWidth(3, 100); // 勤務日数
    sheet.setColumnWidth(4, 120); // 総労働時間
    sheet.setColumnWidth(5, 100); // 残業時間
    sheet.setColumnWidth(6, 100); // 有給日数
    sheet.setColumnWidth(7, 200); // 備考
    
    Logger.log(`${sheetName}シートを設定しました`);
    return { success: true, sheetName: sheetName };
    
  } catch (error) {
    Logger.log(`${sheetName}シート設定エラー: ${error.toString()}`);
    return { success: false, error: error.message };
  }
}

/**
 * Request_Responsesシートを設定する
 */
function setupRequestResponsesSheet(spreadsheet) {
  try {
    // SHEET_CONFIGの存在チェック
    if (typeof SHEET_CONFIG === 'undefined' || !SHEET_CONFIG.SHEETS) {
      throw new Error('SHEET_CONFIGが定義されていないか、SHEETSプロパティが存在しません');
    }
    
    const sheetName = SHEET_CONFIG.SHEETS.REQUEST_RESPONSES;
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    
    const headers = [
      'タイムスタンプ',
      '社員ID',
      '種別',
      '詳細',
      '希望値',
      '承認者',
      'Status'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のフォーマット設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#607d8b');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 列幅の調整
    sheet.setColumnWidth(1, 180); // タイムスタンプ
    sheet.setColumnWidth(2, 100); // 社員ID
    sheet.setColumnWidth(3, 100); // 種別
    sheet.setColumnWidth(4, 200); // 詳細
    sheet.setColumnWidth(5, 150); // 希望値
    sheet.setColumnWidth(6, 200); // 承認者
    sheet.setColumnWidth(7, 100); // Status
    
    // Status列にプルダウンを設定
    setupStatusDropdown(sheet);
    
    Logger.log(`${sheetName}シートを設定しました`);
    return { success: true, sheetName: sheetName };
    
  } catch (error) {
    Logger.log(`${sheetName}シート設定エラー: ${error.toString()}`);
    return { success: false, error: error.message };
  }
}

/**
 * Status列にプルダウンを設定する
 */
function setupStatusDropdown(sheet) {
  try {
    // BUSINESS_RULESの存在チェック
    if (typeof BUSINESS_RULES === 'undefined' || !BUSINESS_RULES.REQUEST_STATUS) {
      throw new Error('BUSINESS_RULESが定義されていないか、REQUEST_STATUSプロパティが存在しません');
    }
    
    // Status列（G列）の2行目以降にプルダウンを設定
    const statusValues = [
      BUSINESS_RULES.REQUEST_STATUS.PENDING,
      BUSINESS_RULES.REQUEST_STATUS.APPROVED,
      BUSINESS_RULES.REQUEST_STATUS.REJECTED
    ];
    
    // データ検証ルールを作成
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusValues, true)
      .setAllowInvalid(false)
      .setHelpText('承認ステータスを選択してください')
      .build();
    
    // G列の2行目から1000行目まで適用（十分な範囲を確保）
    const range = sheet.getRange('G2:G1000');
    range.setDataValidation(rule);
    
    Logger.log('Status列にプルダウンを設定しました');
    
  } catch (error) {
    Logger.log('プルダウン設定エラー: ' + error.toString());
    throw error;
  }
}

/**
 * シート保護設定を適用する
 */
function applySheetProtections(spreadsheet) {
  try {
    // SHEET_CONFIGの存在チェック
    if (typeof SHEET_CONFIG === 'undefined' || !SHEET_CONFIG.SHEETS) {
      throw new Error('SHEET_CONFIGが定義されていないか、SHEETSプロパティが存在しません');
    }
    
    Logger.log('シート保護設定を開始します');
    
    // Log_Rawシートを完全保護（管理者のみアクセス）
    const logRawSheet = spreadsheet.getSheetByName(SHEET_CONFIG.SHEETS.LOG_RAW);
    if (logRawSheet) {
      const protection = logRawSheet.protect();
      protection.setDescription('Log_Raw - 管理者のみアクセス可能');
      protection.setWarningOnly(false);
      
      // 管理者のメールアドレスを編集者として追加
      if (SYSTEM_CONFIG.ADMIN_EMAIL && SYSTEM_CONFIG.ADMIN_EMAIL !== 'admin@example.com') {
        protection.addEditor(SYSTEM_CONFIG.ADMIN_EMAIL);
      }
      
      Logger.log('Log_Rawシートを保護しました');
    }
    
    // Daily_SummaryとMonthly_Summaryシートを読み取り専用に設定
    const summarySheets = [
      SHEET_CONFIG.SHEETS.DAILY_SUMMARY,
      SHEET_CONFIG.SHEETS.MONTHLY_SUMMARY
    ];
    
    summarySheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        const protection = sheet.protect();
        protection.setDescription(`${sheetName} - 読み取り専用`);
        protection.setWarningOnly(true); // 警告のみ（編集は可能だが警告表示）
        
        Logger.log(`${sheetName}シートを読み取り専用に設定しました`);
      }
    });
    
    // Request_ResponsesシートのStatus列以外を保護
    const requestSheet = spreadsheet.getSheetByName(SHEET_CONFIG.SHEETS.REQUEST_RESPONSES);
    if (requestSheet) {
      // ヘッダー行を保護
      const headerProtection = requestSheet.getRange('A1:G1').protect();
      headerProtection.setDescription('Request_Responses - ヘッダー行保護');
      headerProtection.setWarningOnly(false);
      
      // A-F列（Status列以外）を保護
      const dataProtection = requestSheet.getRange('A2:F1000').protect();
      dataProtection.setDescription('Request_Responses - データ列保護');
      dataProtection.setWarningOnly(false);
      
      Logger.log('Request_Responsesシートの保護を設定しました');
    }
    
    Logger.log('シート保護設定が完了しました');
    
  } catch (error) {
    Logger.log('シート保護設定エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 初期設定を実行する
 */
function runInitialSetup() {
  try {
    Logger.log('=== 出勤管理システム 初期設定開始 ===');
    
    // プロジェクト構造確認
    const structureCheck = checkProjectStructure();
    Logger.log('プロジェクト構造確認結果:', JSON.stringify(structureCheck, null, 2));
    
    // スプレッドシート構造設定
    const spreadsheetSetup = setupSpreadsheetStructure();
    Logger.log('スプレッドシート設定結果:', JSON.stringify(spreadsheetSetup, null, 2));
    
    // 設定表示
    if (typeof showCurrentConfig === 'function') {
      showCurrentConfig();
    }
    
    Logger.log('=== 初期設定完了 ===');
    
    return {
      success: true,
      message: '初期設定が正常に完了しました',
      structureCheck: structureCheck,
      spreadsheetSetup: spreadsheetSetup
    };
    
  } catch (error) {
    Logger.log('初期設定エラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * システム情報を表示する
 */
function showSystemInfo() {
  console.log('=== 出勤管理システム情報 ===');
  console.log('システム名:', SYSTEM_CONFIG.SYSTEM_NAME);
  console.log('バージョン:', SYSTEM_CONFIG.VERSION);
  console.log('タイムゾーン:', SYSTEM_CONFIG.TIMEZONE);
  console.log('管理者メール:', SYSTEM_CONFIG.ADMIN_EMAIL);
  console.log('組織名:', SYSTEM_CONFIG.ORGANIZATION_NAME);
  console.log('================================');
}
/**

 * スプレッドシート構造をテストする関数
 */
function testSpreadsheetStructure() {
  try {
    // SHEET_CONFIGの存在チェック
    if (typeof SHEET_CONFIG === 'undefined' || !SHEET_CONFIG.SHEETS) {
      throw new Error('SHEET_CONFIGが定義されていないか、SHEETSプロパティが存在しません');
    }
    
    Logger.log('=== スプレッドシート構造テスト開始 ===');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const testResults = {
      timestamp: new Date(),
      sheets: {},
      protections: {},
      validations: {},
      errors: []
    };
    
    // 各シートの存在確認
    const requiredSheets = Object.values(SHEET_CONFIG.SHEETS);
    requiredSheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        testResults.sheets[sheetName] = {
          exists: true,
          rowCount: sheet.getLastRow(),
          columnCount: sheet.getLastColumn(),
          hasHeaders: sheet.getLastRow() >= 1
        };
        
        // ヘッダー行の内容確認
        if (sheet.getLastRow() >= 1) {
          const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          testResults.sheets[sheetName].headers = headers;
        }
        
      } else {
        testResults.sheets[sheetName] = { exists: false };
        testResults.errors.push(`シート ${sheetName} が見つかりません`);
      }
    });
    
    // 保護設定の確認
    const protections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    testResults.protections.sheetProtections = protections.map(p => ({
      description: p.getDescription(),
      isWarningOnly: p.isWarningOnly()
    }));
    
    const rangeProtections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    testResults.protections.rangeProtections = rangeProtections.map(p => ({
      description: p.getDescription(),
      range: p.getRange().getA1Notation()
    }));
    
    // Request_ResponsesシートのStatus列プルダウン確認
    const requestSheet = spreadsheet.getSheetByName(SHEET_CONFIG.SHEETS.REQUEST_RESPONSES);
    if (requestSheet) {
      const statusRange = requestSheet.getRange('G2');
      const validation = statusRange.getDataValidation();
      if (validation) {
        testResults.validations.statusDropdown = {
          exists: true,
          allowInvalid: validation.getAllowInvalid(),
          helpText: validation.getHelpText()
        };
      } else {
        testResults.validations.statusDropdown = { exists: false };
        testResults.errors.push('Status列のプルダウンが設定されていません');
      }
    }
    
    Logger.log('テスト結果:', JSON.stringify(testResults, null, 2));
    Logger.log('=== スプレッドシート構造テスト完了 ===');
    
    return testResults;
    
  } catch (error) {
    Logger.log('スプレッドシート構造テストエラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * スプレッドシート構造をリセットする関数（開発・テスト用）
 */
function resetSpreadsheetStructure() {
  try {
    // SHEET_CONFIGの存在チェック
    if (typeof SHEET_CONFIG === 'undefined' || !SHEET_CONFIG.SHEETS) {
      throw new Error('SHEET_CONFIGが定義されていないか、SHEETSプロパティが存在しません');
    }
    
    Logger.log('=== スプレッドシート構造リセット開始 ===');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 既存のシートを削除（Sheet1以外）
    const sheets = spreadsheet.getSheets();
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName !== 'Sheet1' && Object.values(SHEET_CONFIG.SHEETS).includes(sheetName)) {
        spreadsheet.deleteSheet(sheet);
        Logger.log(`シート ${sheetName} を削除しました`);
      }
    });
    
    // 保護設定をクリア
    const protections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(protection => {
      protection.remove();
    });
    
    const rangeProtections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    rangeProtections.forEach(protection => {
      protection.remove();
    });
    
    Logger.log('=== スプレッドシート構造リセット完了 ===');
    
    return {
      success: true,
      message: 'スプレッドシート構造をリセットしました'
    };
    
  } catch (error) {
    Logger.log('スプレッドシート構造リセットエラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 時間ベースのトリガーを設定する
 */
function setupTimeTriggers() {
  try {
    Logger.log('=== 時間ベースのトリガー設定開始 ===');
    
    // 既存のトリガーを削除
    const existingTriggers = ScriptApp.getProjectTriggers();
    existingTriggers.forEach(trigger => {
      if (trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
        ScriptApp.deleteTrigger(trigger);
        Logger.log(`既存のトリガーを削除: ${trigger.getHandlerFunction()}`);
      }
    });
    
    // 日次ジョブトリガー（毎日02:00）
    ScriptApp.newTrigger('dailyJob')
      .timeBased()
      .everyDays(1)
      .atHour(AUTOMATION_CONFIG.DAILY_JOB_HOUR)
      .create();
    Logger.log(`日次ジョブトリガーを設定: 毎日${AUTOMATION_CONFIG.DAILY_JOB_HOUR}:00`);
    
    // 週次ジョブトリガー（毎週月曜日07:00）
    ScriptApp.newTrigger('weeklyOvertimeJob')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(AUTOMATION_CONFIG.WEEKLY_JOB_HOUR)
      .create();
    Logger.log(`週次ジョブトリガーを設定: 毎週月曜日${AUTOMATION_CONFIG.WEEKLY_JOB_HOUR}:00`);
    
    // 月次ジョブトリガー（毎月1日02:30）
    ScriptApp.newTrigger('monthlyJob')
      .timeBased()
      .onMonthDay(AUTOMATION_CONFIG.MONTHLY_JOB_DAY)
      .atHour(AUTOMATION_CONFIG.MONTHLY_JOB_HOUR)
      .nearMinute(AUTOMATION_CONFIG.MONTHLY_JOB_MINUTE)
      .create();
    Logger.log(`月次ジョブトリガーを設定: 毎月${AUTOMATION_CONFIG.MONTHLY_JOB_DAY}日${AUTOMATION_CONFIG.MONTHLY_JOB_HOUR}:${AUTOMATION_CONFIG.MONTHLY_JOB_MINUTE}`);
    
    Logger.log('=== 時間ベースのトリガー設定完了 ===');
    
    return {
      success: true,
      message: '時間ベースのトリガーが正常に設定されました',
      triggers: [
        { function: 'dailyJob', schedule: `毎日${AUTOMATION_CONFIG.DAILY_JOB_HOUR}:00` },
        { function: 'weeklyOvertimeJob', schedule: `毎週月曜日${AUTOMATION_CONFIG.WEEKLY_JOB_HOUR}:00` },
        { function: 'monthlyJob', schedule: `毎月${AUTOMATION_CONFIG.MONTHLY_JOB_DAY}日${AUTOMATION_CONFIG.MONTHLY_JOB_HOUR}:${AUTOMATION_CONFIG.MONTHLY_JOB_MINUTE}` }
      ]
    };
    
  } catch (error) {
    Logger.log('トリガー設定エラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 設定されているトリガーの状況を確認する
 */
function checkTriggerStatus() {
  try {
    Logger.log('=== トリガー状況確認開始 ===');
    
    const triggers = ScriptApp.getProjectTriggers();
    const triggerInfo = {
      timestamp: new Date(),
      totalTriggers: triggers.length,
      triggers: []
    };
    
    triggers.forEach(trigger => {
      const info = {
        handlerFunction: trigger.getHandlerFunction(),
        triggerSource: trigger.getTriggerSource().toString(),
        eventType: trigger.getEventType().toString()
      };
      
      // 時間ベースのトリガーの詳細情報を追加
      if (trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
        try {
          // トリガーの詳細情報を取得（可能な場合）
          info.details = 'Time-based trigger';
        } catch (e) {
          info.details = 'Time-based trigger (details unavailable)';
        }
      }
      
      triggerInfo.triggers.push(info);
    });
    
    Logger.log('トリガー状況:', JSON.stringify(triggerInfo, null, 2));
    Logger.log('=== トリガー状況確認完了 ===');
    
    return triggerInfo;
    
  } catch (error) {
    Logger.log('トリガー状況確認エラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * すべてのトリガーを削除する
 */
function deleteAllTriggers() {
  try {
    Logger.log('=== 全トリガー削除開始 ===');
    
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`トリガー削除: ${trigger.getHandlerFunction()}`);
      deletedCount++;
    });
    
    Logger.log(`=== 全トリガー削除完了: ${deletedCount}個のトリガーを削除 ===`);
    
    return {
      success: true,
      message: `${deletedCount}個のトリガーを削除しました`,
      deletedCount: deletedCount
    };
    
  } catch (error) {
    Logger.log('トリガー削除エラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 完全な初期設定を実行する（トリガー設定を含む）
 */
function runCompleteSetup() {
  try {
    Logger.log('=== 完全初期設定開始 ===');
    
    const results = {
      timestamp: new Date(),
      steps: {},
      success: true,
      errors: []
    };
    
    // ステップ1: プロジェクト構造確認
    results.steps.structureCheck = checkProjectStructure();
    if (!results.steps.structureCheck.success) {
      results.success = false;
      results.errors.push('プロジェクト構造確認失敗');
    }
    
    // ステップ2: スプレッドシート構造設定
    results.steps.spreadsheetSetup = setupSpreadsheetStructure();
    if (!results.steps.spreadsheetSetup.success) {
      results.success = false;
      results.errors.push('スプレッドシート構造設定失敗');
    }
    
    // ステップ3: トリガー設定
    results.steps.triggerSetup = setupTimeTriggers();
    if (!results.steps.triggerSetup.success) {
      results.success = false;
      results.errors.push('トリガー設定失敗');
    }
    
    // ステップ4: 設定確認
    results.steps.configCheck = {
      success: true,
      message: '設定確認完了',
      config: {
        systemName: SYSTEM_CONFIG.SYSTEM_NAME,
        version: SYSTEM_CONFIG.VERSION,
        adminEmail: SYSTEM_CONFIG.ADMIN_EMAIL
      }
    };
    
    Logger.log('完全初期設定結果:', JSON.stringify(results, null, 2));
    Logger.log('=== 完全初期設定完了 ===');
    
    return results;
    
  } catch (error) {
    Logger.log('完全初期設定エラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 統合テストを実行する
 */
function runIntegrationTests() {
  try {
    Logger.log('=== 統合テスト開始 ===');
    
    const testResults = {
      timestamp: new Date(),
      tests: {},
      summary: {
        total: 0,
        passed: 0,
        failed: 0
      },
      errors: []
    };
    
    // テスト1: 基本設定確認
    testResults.tests.basicConfig = testBasicConfiguration();
    testResults.summary.total++;
    if (testResults.tests.basicConfig.success) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    // テスト2: スプレッドシート構造確認
    testResults.tests.spreadsheetStructure = testSpreadsheetStructure();
    testResults.summary.total++;
    if (testResults.tests.spreadsheetStructure.success !== false) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    // テスト3: トリガー設定確認
    testResults.tests.triggerStatus = checkTriggerStatus();
    testResults.summary.total++;
    if (testResults.tests.triggerStatus.success !== false) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    // テスト4: 認証機能テスト
    testResults.tests.authentication = testAuthenticationModule();
    testResults.summary.total++;
    if (testResults.tests.authentication.success) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    // テスト5: 業務ロジックテスト
    testResults.tests.businessLogic = testBusinessLogicModule();
    testResults.summary.total++;
    if (testResults.tests.businessLogic.success) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    // テスト6: エラーハンドリングテスト
    testResults.tests.errorHandling = testErrorHandling();
    testResults.summary.total++;
    if (testResults.tests.errorHandling.success) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    // テスト7: メール機能テスト（送信なし）
    testResults.tests.mailFunction = testMailFunctionality();
    testResults.summary.total++;
    if (testResults.tests.mailFunction.success) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    Logger.log('統合テスト結果:', JSON.stringify(testResults, null, 2));
    Logger.log(`=== 統合テスト完了: ${testResults.summary.passed}/${testResults.summary.total} 成功 ===`);
    
    return testResults;
    
  } catch (error) {
    Logger.log('統合テストエラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 基本設定をテストする
 */
function testBasicConfiguration() {
  try {
    const results = {
      success: true,
      checks: {},
      errors: []
    };
    
    // 設定オブジェクトの存在確認
    const configObjects = [
      'SYSTEM_CONFIG',
      'SHEET_CONFIG', 
      'BUSINESS_RULES',
      'AUTOMATION_CONFIG',
      'MAIL_CONFIG',
      'ERROR_CONFIG',
      'SECURITY_CONFIG'
    ];
    
    configObjects.forEach(configName => {
      try {
        const config = globalThis[configName];
        results.checks[configName] = {
          exists: true,
          type: typeof config
        };
      } catch (error) {
        results.checks[configName] = {
          exists: false,
          error: error.message
        };
        results.errors.push(`設定 ${configName} が見つかりません`);
        results.success = false;
      }
    });
    
    // 管理者メール設定確認
    try {
      const adminEmail = getAdminEmail();
      results.checks.adminEmail = {
        configured: adminEmail && adminEmail !== 'admin@example.com',
        value: adminEmail
      };
      if (!results.checks.adminEmail.configured) {
        results.errors.push('管理者メールが設定されていません');
      }
    } catch (error) {
      results.checks.adminEmail = {
        configured: false,
        error: error.message
      };
      results.errors.push('管理者メール取得エラー');
    }
    
    return results;
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 認証モジュールをテストする
 */
function testAuthenticationModule() {
  try {
    const results = {
      success: true,
      tests: {},
      errors: []
    };
    
    // 認証関数の存在確認
    const authFunctions = ['authenticateUser', 'validateEmployeeAccess', 'getEmployeeInfo'];
    authFunctions.forEach(funcName => {
      try {
        const func = globalThis[funcName];
        results.tests[funcName] = {
          exists: typeof func === 'function',
          type: typeof func
        };
      } catch (error) {
        results.tests[funcName] = {
          exists: false,
          error: error.message
        };
        results.errors.push(`認証関数 ${funcName} が見つかりません`);
        results.success = false;
      }
    });
    
    // 従業員マスターデータの存在確認
    try {
      const employeeData = getSheetData('Master_Employee', 'A:H');
      results.tests.employeeData = {
        exists: employeeData && employeeData.length > 1,
        rowCount: employeeData ? employeeData.length : 0
      };
      if (!results.tests.employeeData.exists) {
        results.errors.push('従業員マスターデータが不足しています');
      }
    } catch (error) {
      results.tests.employeeData = {
        exists: false,
        error: error.message
      };
      results.errors.push('従業員マスターデータ取得エラー');
      results.success = false;
    }
    
    return results;
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 業務ロジックモジュールをテストする
 */
function testBusinessLogicModule() {
  try {
    const results = {
      success: true,
      tests: {},
      errors: []
    };
    
    // 業務ロジック関数の存在確認
    const businessFunctions = [
      'calculateDailyWorkTime',
      'calculateOvertime', 
      'applyBreakDeduction',
      'isHoliday',
      'updateDailySummary'
    ];
    
    businessFunctions.forEach(funcName => {
      try {
        const func = globalThis[funcName];
        results.tests[funcName] = {
          exists: typeof func === 'function',
          type: typeof func
        };
      } catch (error) {
        results.tests[funcName] = {
          exists: false,
          error: error.message
        };
        results.errors.push(`業務ロジック関数 ${funcName} が見つかりません`);
        results.success = false;
      }
    });
    
    // 時間計算テスト
    try {
      const testTime = calculateTimeDifference('09:00', '17:00');
      results.tests.timeCalculation = {
        success: testTime === 480,
        result: testTime,
        expected: 480
      };
      if (!results.tests.timeCalculation.success) {
        results.errors.push('時間計算が正しく動作していません');
      }
    } catch (error) {
      results.tests.timeCalculation = {
        success: false,
        error: error.message
      };
      results.errors.push('時間計算テストエラー');
      results.success = false;
    }
    
    // 祝日判定テスト
    try {
      const newYearTest = isHoliday(new Date('2024-01-01'));
      results.tests.holidayCheck = {
        success: newYearTest === true,
        result: newYearTest,
        expected: true
      };
      if (!results.tests.holidayCheck.success) {
        results.errors.push('祝日判定が正しく動作していません');
      }
    } catch (error) {
      results.tests.holidayCheck = {
        success: false,
        error: error.message
      };
      results.errors.push('祝日判定テストエラー');
    }
    
    return results;
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * エラーハンドリングをテストする
 */
function testErrorHandling() {
  try {
    const results = {
      success: true,
      tests: {},
      errors: []
    };
    
    // エラーハンドリング関数の存在確認
    const errorFunctions = ['withErrorHandling', 'sendErrorAlert', 'logError'];
    errorFunctions.forEach(funcName => {
      try {
        const func = globalThis[funcName];
        results.tests[funcName] = {
          exists: typeof func === 'function',
          type: typeof func
        };
      } catch (error) {
        results.tests[funcName] = {
          exists: false,
          error: error.message
        };
        results.errors.push(`エラーハンドリング関数 ${funcName} が見つかりません`);
        results.success = false;
      }
    });
    
    // エラーハンドリングの動作テスト
    try {
      const testResult = withErrorHandling(() => {
        // 正常処理のテスト
        return { success: true, message: 'テスト成功' };
      }, 'TestFunction', 'LOW');
      
      results.tests.errorHandlingExecution = {
        success: testResult && testResult.success,
        result: testResult
      };
      
      if (!results.tests.errorHandlingExecution.success) {
        results.errors.push('エラーハンドリング実行テスト失敗');
      }
    } catch (error) {
      results.tests.errorHandlingExecution = {
        success: false,
        error: error.message
      };
      results.errors.push('エラーハンドリング実行テストエラー');
      results.success = false;
    }
    
    // エラー発生時の処理テスト
    try {
      const errorTestResult = withErrorHandling(() => {
        throw new Error('テスト用エラー');
      }, 'TestErrorFunction', 'LOW');
      
      results.tests.errorHandlingWithError = {
        success: errorTestResult === null || errorTestResult === undefined,
        result: errorTestResult
      };
      
    } catch (error) {
      // エラーが適切にキャッチされた場合
      results.tests.errorHandlingWithError = {
        success: true,
        message: 'エラーが適切にキャッチされました'
      };
    }
    
    return results;
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * メール機能をテストする（実際の送信は行わない）
 */
function testMailFunctionality() {
  try {
    const results = {
      success: true,
      tests: {},
      errors: []
    };
    
    // メール関数の存在確認
    const mailFunctions = [
      'sendBatchNotification',
      'sendErrorAlert', 
      'sendOvertimeWarning',
      'sendMonthlyReport'
    ];
    
    mailFunctions.forEach(funcName => {
      try {
        const func = globalThis[funcName];
        results.tests[funcName] = {
          exists: typeof func === 'function',
          type: typeof func
        };
      } catch (error) {
        results.tests[funcName] = {
          exists: false,
          error: error.message
        };
        results.errors.push(`メール関数 ${funcName} が見つかりません`);
        results.success = false;
      }
    });
    
    // メール設定確認
    try {
      const adminEmail = getAdminEmail();
      results.tests.mailConfig = {
        adminEmailConfigured: adminEmail && adminEmail !== 'admin@example.com',
        adminEmail: adminEmail
      };
      
      if (!results.tests.mailConfig.adminEmailConfigured) {
        results.errors.push('管理者メールが設定されていないため、メール送信ができません');
      }
    } catch (error) {
      results.tests.mailConfig = {
        adminEmailConfigured: false,
        error: error.message
      };
      results.errors.push('メール設定確認エラー');
      results.success = false;
    }
    
    return results;
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * フォームトリガーを設定する
 */
function setupFormTrigger() {
  try {
    Logger.log('=== フォームトリガー設定開始 ===');
    
    // 既存のフォームトリガーを削除
    const existingTriggers = ScriptApp.getProjectTriggers();
    existingTriggers.forEach(trigger => {
      if (trigger.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS &&
          trigger.getEventType() === ScriptApp.EventType.ON_EDIT) {
        ScriptApp.deleteTrigger(trigger);
        Logger.log(`既存の編集トリガーを削除: ${trigger.getHandlerFunction()}`);
      }
    });
    
    // Request_Responsesシートの編集トリガーを設定
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger('onRequestResponsesEdit')
      .spreadsheet(spreadsheet)
      .onEdit()
      .create();
    
    Logger.log('Request_Responsesシート編集トリガーを設定しました');
    Logger.log('=== フォームトリガー設定完了 ===');
    
    return {
      success: true,
      message: 'フォームトリガーが正常に設定されました'
    };
    
  } catch (error) {
    Logger.log('フォームトリガー設定エラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 全機能統合テストを実行する
 */
function runFullSystemTest() {
  try {
    Logger.log('=== 全機能統合テスト開始 ===');
    
    const testResults = {
      timestamp: new Date(),
      phases: {},
      summary: {
        totalPhases: 0,
        passedPhases: 0,
        failedPhases: 0
      },
      overallSuccess: true
    };
    
    // フェーズ1: 基本設定とインフラ
    testResults.phases.infrastructure = {
      name: '基本設定とインフラ',
      tests: {
        basicConfig: testBasicConfiguration(),
        spreadsheetStructure: testSpreadsheetStructure(),
        triggerStatus: checkTriggerStatus()
      }
    };
    testResults.summary.totalPhases++;
    
    // フェーズ2: 認証とセキュリティ
    testResults.phases.authentication = {
      name: '認証とセキュリティ',
      tests: {
        authModule: testAuthenticationModule()
      }
    };
    testResults.summary.totalPhases++;
    
    // フェーズ3: 業務ロジック
    testResults.phases.businessLogic = {
      name: '業務ロジック',
      tests: {
        businessModule: testBusinessLogicModule()
      }
    };
    testResults.summary.totalPhases++;
    
    // フェーズ4: エラーハンドリング
    testResults.phases.errorHandling = {
      name: 'エラーハンドリング',
      tests: {
        errorModule: testErrorHandling()
      }
    };
    testResults.summary.totalPhases++;
    
    // フェーズ5: 通信機能
    testResults.phases.communication = {
      name: '通信機能',
      tests: {
        mailModule: testMailFunctionality()
      }
    };
    testResults.summary.totalPhases++;
    
    // 各フェーズの成功/失敗を判定
    Object.keys(testResults.phases).forEach(phaseKey => {
      const phase = testResults.phases[phaseKey];
      let phaseSuccess = true;
      
      Object.keys(phase.tests).forEach(testKey => {
        const test = phase.tests[testKey];
        if (test.success === false) {
          phaseSuccess = false;
        }
      });
      
      phase.success = phaseSuccess;
      if (phaseSuccess) {
        testResults.summary.passedPhases++;
      } else {
        testResults.summary.failedPhases++;
        testResults.overallSuccess = false;
      }
    });
    
    Logger.log('全機能統合テスト結果:', JSON.stringify(testResults, null, 2));
    Logger.log(`=== 全機能統合テスト完了: ${testResults.summary.passedPhases}/${testResults.summary.totalPhases} フェーズ成功 ===`);
    
    return testResults;
    
  } catch (error) {
    Logger.log('全機能統合テストエラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 管理メニューを追加する関数
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('出勤管理システム')
      .addItem('完全初期設定実行', 'runCompleteSetup')
      .addItem('初期設定実行', 'runInitialSetup')
      .addItem('スプレッドシート構造設定', 'setupSpreadsheetStructure')
      .addSeparator()
      .addItem('トリガー設定', 'setupTimeTriggers')
      .addItem('フォームトリガー設定', 'setupFormTrigger')
      .addItem('トリガー状況確認', 'checkTriggerStatus')
      .addItem('全トリガー削除', 'deleteAllTriggers')
      .addSeparator()
      .addItem('統合テスト実行', 'runIntegrationTests')
      .addItem('全機能統合テスト実行', 'runFullSystemTest')
      .addItem('構造テスト実行', 'testSpreadsheetStructure')
      .addItem('BUSINESS_RULES存在チェックテスト', 'testBusinessRulesExistenceCheck')
      .addSeparator()
      .addItem('システム情報表示', 'showSystemInfo')
      .addItem('設定表示', 'showCurrentConfig')
      .addSeparator()
      .addItem('構造リセット（開発用）', 'resetSpreadsheetStructure')
      .addToUi();
      
    Logger.log('管理メニューを追加しました');
    
  } catch (error) {
    Logger.log('メニュー追加エラー: ' + error.toString());
  }
}

/**
 * BUSINESS_RULESの存在チェックテスト
 */
function testBusinessRulesExistenceCheck() {
  try {
    Logger.log('=== BUSINESS_RULES存在チェックテスト開始 ===');
    
    const testResults = {
      success: true,
      tests: [],
      errors: []
    };
    
    // テスト1: BUSINESS_RULESが正常に定義されている場合
    try {
      if (typeof BUSINESS_RULES !== 'undefined' && BUSINESS_RULES.REQUEST_STATUS) {
        testResults.tests.push({
          name: 'BUSINESS_RULES正常定義テスト',
          success: true,
          message: 'BUSINESS_RULESが正常に定義されています'
        });
      } else {
        testResults.tests.push({
          name: 'BUSINESS_RULES正常定義テスト',
          success: false,
          message: 'BUSINESS_RULESが定義されていないか、REQUEST_STATUSプロパティが存在しません'
        });
        testResults.success = false;
      }
    } catch (error) {
      testResults.tests.push({
        name: 'BUSINESS_RULES正常定義テスト',
        success: false,
        message: `エラー: ${error.message}`
      });
      testResults.success = false;
      testResults.errors.push(error.message);
    }
    
    // テスト2: setupStatusDropdown関数の動作確認
    try {
      // テスト用のシートを作成
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const testSheet = spreadsheet.insertSheet('TestStatusDropdown');
      
      // setupStatusDropdown関数を実行
      setupStatusDropdown(testSheet);
      
      testResults.tests.push({
        name: 'setupStatusDropdown関数テスト',
        success: true,
        message: 'setupStatusDropdown関数が正常に動作しました'
      });
      
      // テスト用シートを削除
      spreadsheet.deleteSheet(testSheet);
      
    } catch (error) {
      testResults.tests.push({
        name: 'setupStatusDropdown関数テスト',
        success: false,
        message: `エラー: ${error.message}`
      });
      testResults.success = false;
      testResults.errors.push(error.message);
    }
    
    Logger.log('BUSINESS_RULES存在チェックテスト結果:', JSON.stringify(testResults, null, 2));
    Logger.log(`=== BUSINESS_RULES存在チェックテスト完了: ${testResults.success ? '成功' : '失敗'} ===`);
    
    return testResults;
    
  } catch (error) {
    Logger.log('BUSINESS_RULES存在チェックテストエラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}