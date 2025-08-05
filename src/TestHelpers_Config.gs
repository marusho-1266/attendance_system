// ========== 設定関連テストヘルパー関数 ==========

/**
 * 設定オブジェクトの存在確認
 */
function testConfigObjectsExistence() {
  try {
    const configObjects = [
      'SYSTEM_CONFIG',
      'SHEET_CONFIG',
      'BUSINESS_RULES',
      'AUTOMATION_CONFIG',
      'MAIL_CONFIG',
      'ERROR_CONFIG',
      'SECURITY_CONFIG'
    ];
    
    const results = {};
    let allExist = true;
    
    configObjects.forEach(configName => {
      try {
        // 安全な関数存在チェック - eval()の代わりにグローバルスコープから直接アクセス
        const config = this[configName] || (typeof window !== 'undefined' ? window[configName] : null);
        results[configName] = {
          exists: config !== null && config !== undefined,
          type: typeof config
        };
      } catch (error) {
        results[configName] = {
          exists: false,
          error: error.message
        };
        allExist = false;
      }
    });
    
    return {
      success: allExist,
      results: results
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * スプレッドシート構造の整合性確認
 */
function testSpreadsheetStructureIntegrity() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = Object.values(SHEET_CONFIG.SHEETS);
    const existingSheets = spreadsheet.getSheets().map(s => s.getName());
    
    const missingSheets = requiredSheets.filter(sheet => !existingSheets.includes(sheet));
    const extraSheets = existingSheets.filter(sheet => !requiredSheets.includes(sheet) && sheet !== 'Sheet1');
    
    return {
      success: missingSheets.length === 0,
      requiredSheets: requiredSheets,
      existingSheets: existingSheets,
      missingSheets: missingSheets,
      extraSheets: extraSheets
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 管理者設定の確認
 */
function testAdminConfiguration() {
  try {
    const adminEmail = getAdminEmail();
    const isConfigured = adminEmail && adminEmail !== 'admin@example.com';
    
    return {
      success: isConfigured,
      adminEmail: adminEmail,
      isConfigured: isConfigured
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 時間トリガーの設定確認
 */
function testTimeTriggerConfiguration(triggers) {
  try {
    const timeBasedTriggers = triggers.filter(t => t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK);
    
    return {
      success: timeBasedTriggers.length >= 3,
      timeBasedTriggerCount: timeBasedTriggers.length,
      triggerFunctions: timeBasedTriggers.map(t => t.getHandlerFunction())
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * シート保護設定の確認
 */
function testSheetProtections() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const protections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    
    // Log_Rawシートが保護されているかを確認
    const logRawProtected = protections.some(p => {
      const range = p.getRange();
      return range && range.getSheet().getName() === 'Log_Raw';
    });
    
    return {
      success: protections.length > 0,
      protectionCount: protections.length,
      logRawProtected: logRawProtected
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * データ検証ルールの確認
 */
function testDataValidationRules() {
  try {
    const sheet = getSheet('Request_Responses');
    const statusRange = sheet.getRange('G2');
    const validation = statusRange.getDataValidation();
    
    return {
      success: validation !== null,
      hasStatusValidation: validation !== null
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
} 