// ========== セキュリティ関連テストヘルパー関数 ==========

/**
 * システム権限の確認
 */
function testSystemPermissions() {
  try {
    const checks = {
      spreadsheetAccess: testSpreadsheetAccess(),
      gmailAccess: testGmailAccess(),
      driveAccess: testDriveAccess(),
      scriptAccess: testScriptAccess()
    };
    
    const allPassed = Object.values(checks).every(check => check.success);
    
    return {
      success: allPassed,
      checks: checks
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * アクセス制御の確認
 */
function testAccessControls() {
  try {
    // 基本的なアクセス制御機能の存在確認
    const accessFunctions = ['authenticateUser', 'validateEmployeeAccess'];
    let allExist = true;
    
    accessFunctions.forEach(funcName => {
      try {
        const func = eval(funcName);
        if (typeof func !== 'function') {
          allExist = false;
        }
      } catch (error) {
        allExist = false;
      }
    });
    
    return {
      success: allExist,
      accessControlFunctionsExist: allExist
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 監査ログ設定の確認
 */
function testAuditLoggingSetup() {
  try {
    // Logger機能の基本的な確認
    const loggerAvailable = typeof Logger !== 'undefined' && typeof Logger.log === 'function';
    
    return {
      success: loggerAvailable,
      loggerAvailable: loggerAvailable
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * スプレッドシートアクセスのテスト
 */
function testSpreadsheetAccess() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    return {
      success: sheets.length > 0,
      sheetCount: sheets.length
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Gmailアクセスのテスト
 */
function testGmailAccess() {
  try {
    // Gmail APIの基本的な利用可能性を確認
    const gmailAvailable = typeof GmailApp !== 'undefined';
    
    return {
      success: gmailAvailable,
      gmailAvailable: gmailAvailable
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Driveアクセスのテスト
 */
function testDriveAccess() {
  try {
    // Drive APIの基本的な利用可能性を確認
    const driveAvailable = typeof DriveApp !== 'undefined';
    
    return {
      success: driveAvailable,
      driveAvailable: driveAvailable
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * スクリプトアクセスのテスト
 */
function testScriptAccess() {
  try {
    // Script APIの基本的な利用可能性を確認
    const scriptAvailable = typeof ScriptApp !== 'undefined';
    
    return {
      success: scriptAvailable,
      scriptAvailable: scriptAvailable
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
} 