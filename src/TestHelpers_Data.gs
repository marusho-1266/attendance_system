// ========== データ検証関連テストヘルパー関数 ==========

/**
 * 従業員データの一貫性確認
 */
function testEmployeeDataConsistency() {
  try {
    const employeeData = getSheetData('Master_Employee', 'A:H');
    
    if (!employeeData || employeeData.length <= 1) {
      return {
        success: false,
        error: '従業員データが不足しています'
      };
    }
    
    const issues = [];
    
    // ヘッダー行をスキップして検証
    for (let i = 1; i < employeeData.length; i++) {
      const row = employeeData[i];
      const employeeId = row[0];
      const name = row[1];
      const email = row[2];
      
      if (!employeeId) {
        issues.push(`行${i + 1}: 社員IDが空です`);
      }
      if (!name) {
        issues.push(`行${i + 1}: 氏名が空です`);
      }
      if (!email || !email.includes('@')) {
        issues.push(`行${i + 1}: 無効なメールアドレスです`);
      }
    }
    
    return {
      success: issues.length === 0,
      employeeCount: employeeData.length - 1,
      issues: issues
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 祝日データの妥当性確認
 */
function testHolidayDataValidity() {
  try {
    const holidayData = getSheetData('Master_Holiday', 'A:D');
    
    if (!holidayData || holidayData.length <= 1) {
      return {
        success: true, // 祝日データは必須ではない
        holidayCount: 0
      };
    }
    
    const issues = [];
    
    // ヘッダー行をスキップして検証
    for (let i = 1; i < holidayData.length; i++) {
      const row = holidayData[i];
      const date = row[0];
      const name = row[1];
      
      if (!date || !(date instanceof Date)) {
        issues.push(`行${i + 1}: 無効な日付です`);
      }
      if (!name) {
        issues.push(`行${i + 1}: 祝日名が空です`);
      }
    }
    
    return {
      success: issues.length === 0,
      holidayCount: holidayData.length - 1,
      issues: issues
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Request_Responsesシートの整合性確認
 */
function testRequestResponsesIntegrity() {
  try {
    const sheet = getSheet('Request_Responses');
    
    // Status列のデータ検証ルールを確認
    const statusRange = sheet.getRange('G2');
    const validation = statusRange.getDataValidation();
    
    const hasValidation = validation !== null;
    
    return {
      success: hasValidation,
      hasValidation: hasValidation,
      validationRule: hasValidation ? validation.getHelpText() : null
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * サマリーデータの精度確認
 */
function testSummaryDataAccuracy() {
  try {
    const dailySummaryData = getSheetData('Daily_Summary', 'A:I');
    const monthlySummaryData = getSheetData('Monthly_Summary', 'A:G');
    
    return {
      success: true,
      dailySummaryRows: dailySummaryData ? dailySummaryData.length - 1 : 0,
      monthlySummaryRows: monthlySummaryData ? monthlySummaryData.length - 1 : 0
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
} 