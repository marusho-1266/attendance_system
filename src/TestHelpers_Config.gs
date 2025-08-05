// ========== 設定関連テストヘルパー関数 ==========

/**
 * 設定オブジェクトの存在確認
 */
function testConfigObjectsExistence() {
  try {
    const results = {};
    let allExist = true;
    
    // 明示的な設定オブジェクト参照による安全な存在確認
    const configChecks = [
      {
        name: 'SYSTEM_CONFIG',
        config: typeof SYSTEM_CONFIG !== 'undefined' ? SYSTEM_CONFIG : null
      },
      {
        name: 'SHEET_CONFIG',
        config: typeof SHEET_CONFIG !== 'undefined' ? SHEET_CONFIG : null
      },
      {
        name: 'BUSINESS_RULES',
        config: typeof BUSINESS_RULES !== 'undefined' ? BUSINESS_RULES : null
      },
      {
        name: 'AUTOMATION_CONFIG',
        config: typeof AUTOMATION_CONFIG !== 'undefined' ? AUTOMATION_CONFIG : null
      },
      {
        name: 'MAIL_CONFIG',
        config: typeof MAIL_CONFIG !== 'undefined' ? MAIL_CONFIG : null
      },
      {
        name: 'ERROR_CONFIG',
        config: typeof ERROR_CONFIG !== 'undefined' ? ERROR_CONFIG : null
      },
      {
        name: 'SECURITY_CONFIG',
        config: typeof SECURITY_CONFIG !== 'undefined' ? SECURITY_CONFIG : null
      }
    ];
    
    configChecks.forEach(check => {
      try {
        results[check.name] = {
          exists: check.config !== null && check.config !== undefined,
          type: typeof check.config
        };
        
        if (!results[check.name].exists) {
          allExist = false;
        }
      } catch (error) {
        results[check.name] = {
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
    // triggersパラメータの妥当性チェック
    if (!triggers) {
      return {
        success: false,
        error: 'triggersパラメータがnullまたはundefinedです'
      };
    }
    
    if (!Array.isArray(triggers)) {
      return {
        success: false,
        error: `triggersパラメータが配列ではありません。型: ${typeof triggers}`
      };
    }
    
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