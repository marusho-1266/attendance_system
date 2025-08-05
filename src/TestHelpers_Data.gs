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
 * Status列の動的検出とデータ検証ルールの確認
 */
function testRequestResponsesIntegrity() {
  try {
    const sheet = getSheet('Request_Responses');
    
    // ヘッダー行を取得してStatus列の位置を動的に検出
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusColumnIndex = headerRow.findIndex(header => header === 'Status') + 1;
    
    if (statusColumnIndex === 0) {
      return {
        success: false,
        error: 'Status列が見つかりません',
        headerRow: headerRow
      };
    }
    
    // Status列の適用範囲を動的に決定
    const lastRow = sheet.getLastRow();
    const statusRange = sheet.getRange(2, statusColumnIndex, lastRow - 1, 1);
    
    // 範囲全体のデータ検証ルールを確認
    const validation = statusRange.getDataValidation();
    const hasValidation = validation !== null;
    
    // 範囲内の各セルの検証ルールも確認
    const validationDetails = [];
    if (hasValidation) {
      for (let row = 2; row <= lastRow; row++) {
        const cellValidation = sheet.getRange(row, statusColumnIndex).getDataValidation();
        if (cellValidation) {
          validationDetails.push({
            row: row,
            hasValidation: true,
            helpText: cellValidation.getHelpText(),
            criteriaType: cellValidation.getCriteriaType(),
            criteriaValues: cellValidation.getCriteriaValues()
          });
        } else {
          validationDetails.push({
            row: row,
            hasValidation: false
          });
        }
      }
    }
    
    return {
      success: hasValidation,
      hasValidation: hasValidation,
      statusColumnIndex: statusColumnIndex,
      statusColumnLetter: sheet.getRange(1, statusColumnIndex).getA1Notation().replace('1', ''),
      validationRange: statusRange.getA1Notation(),
      totalRowsChecked: lastRow - 1,
      validationRule: hasValidation ? validation.getHelpText() : null,
      validationDetails: validationDetails,
      headerRow: headerRow
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
 * 日次サマリーデータと月次サマリーデータの整合性を検証
 */
function testSummaryDataAccuracy() {
  try {
    const dailySummaryData = getSheetData('Daily_Summary', 'A:I');
    const monthlySummaryData = getSheetData('Monthly_Summary', 'A:G');
    
    if (!dailySummaryData || dailySummaryData.length <= 1) {
      return {
        success: true,
        message: '日次サマリーデータが存在しません（正常）',
        dailySummaryRows: 0,
        monthlySummaryRows: monthlySummaryData ? monthlySummaryData.length - 1 : 0
      };
    }
    
    if (!monthlySummaryData || monthlySummaryData.length <= 1) {
      return {
        success: true,
        message: '月次サマリーデータが存在しません（正常）',
        dailySummaryRows: dailySummaryData.length - 1,
        monthlySummaryRows: 0
      };
    }
    
    const issues = [];
    const accuracyChecks = [];
    
    // ヘッダー行をスキップして検証
    const dailyData = dailySummaryData.slice(1);
    const monthlyData = monthlySummaryData.slice(1);
    
    // 月次データごとに日次データとの整合性をチェック
    for (let i = 0; i < monthlyData.length; i++) {
      const monthlyRow = monthlyData[i];
      const yearMonth = monthlyRow[0]; // 年月 (YYYY-MM)
      const employeeId = monthlyRow[1]; // 社員ID
      const monthlyWorkDays = Number(monthlyRow[2]) || 0; // 勤務日数
      const monthlyTotalWorkMinutes = Number(monthlyRow[3]) || 0; // 総労働時間
      const monthlyOvertimeMinutes = Number(monthlyRow[4]) || 0; // 残業時間
      
      // 該当年月・従業員の日次データを抽出
      const dailyEntries = dailyData.filter(row => {
        const date = new Date(row[0]);
        const dateYearMonth = Utilities.formatDate(date, 'JST', 'yyyy-MM');
        const empId = row[1];
        return dateYearMonth === yearMonth && empId === employeeId;
      });
      
      // 日次データから集計値を計算
      let calculatedWorkDays = 0;
      let calculatedTotalWorkMinutes = 0;
      let calculatedOvertimeMinutes = 0;
      
      dailyEntries.forEach(row => {
        const workMinutes = Number(row[5]) || 0; // 実働時間
        const overtimeMinutes = Number(row[6]) || 0; // 残業時間
        
        if (workMinutes > 0) {
          calculatedWorkDays++;
          calculatedTotalWorkMinutes += workMinutes;
          calculatedOvertimeMinutes += overtimeMinutes;
        }
      });
      
      // 整合性チェック
      const workDaysDiff = Math.abs(monthlyWorkDays - calculatedWorkDays);
      const totalWorkDiff = Math.abs(monthlyTotalWorkMinutes - calculatedTotalWorkMinutes);
      const overtimeDiff = Math.abs(monthlyOvertimeMinutes - calculatedOvertimeMinutes);
      
      const checkResult = {
        yearMonth: yearMonth,
        employeeId: employeeId,
        monthlyWorkDays: monthlyWorkDays,
        calculatedWorkDays: calculatedWorkDays,
        workDaysDiff: workDaysDiff,
        monthlyTotalWorkMinutes: monthlyTotalWorkMinutes,
        calculatedTotalWorkMinutes: calculatedTotalWorkMinutes,
        totalWorkDiff: totalWorkDiff,
        monthlyOvertimeMinutes: monthlyOvertimeMinutes,
        calculatedOvertimeMinutes: calculatedOvertimeMinutes,
        overtimeDiff: overtimeDiff,
        hasIssues: false
      };
      
      // 許容誤差を設定（1分以内の差は許容）
      const tolerance = 1;
      
      if (workDaysDiff > 0) {
        issues.push(`${yearMonth} ${employeeId}: 勤務日数の不一致 (月次: ${monthlyWorkDays}, 日次集計: ${calculatedWorkDays})`);
        checkResult.hasIssues = true;
      }
      
      if (totalWorkDiff > tolerance) {
        issues.push(`${yearMonth} ${employeeId}: 総労働時間の不一致 (月次: ${monthlyTotalWorkMinutes}分, 日次集計: ${calculatedTotalWorkMinutes}分, 差: ${totalWorkDiff}分)`);
        checkResult.hasIssues = true;
      }
      
      if (overtimeDiff > tolerance) {
        issues.push(`${yearMonth} ${employeeId}: 残業時間の不一致 (月次: ${monthlyOvertimeMinutes}分, 日次集計: ${calculatedOvertimeMinutes}分, 差: ${overtimeDiff}分)`);
        checkResult.hasIssues = true;
      }
      
      accuracyChecks.push(checkResult);
    }
    
    // 日次データに存在するが月次データに存在しない組み合わせをチェック
    const monthlyKeys = new Set();
    monthlyData.forEach(row => {
      const yearMonth = row[0];
      const employeeId = row[1];
      monthlyKeys.add(`${yearMonth}_${employeeId}`);
    });
    
    const dailyKeys = new Set();
    dailyData.forEach(row => {
      const date = new Date(row[0]);
      const yearMonth = Utilities.formatDate(date, 'JST', 'yyyy-MM');
      const employeeId = row[1];
      dailyKeys.add(`${yearMonth}_${employeeId}`);
    });
    
    // 日次データのみに存在する組み合わせを検出
    const missingMonthlyEntries = [];
    dailyKeys.forEach(key => {
      if (!monthlyKeys.has(key)) {
        const [yearMonth, employeeId] = key.split('_');
        missingMonthlyEntries.push(`${yearMonth} ${employeeId}`);
      }
    });
    
    if (missingMonthlyEntries.length > 0) {
      issues.push(`月次サマリーに存在しない日次データ: ${missingMonthlyEntries.join(', ')}`);
    }
    
    return {
      success: issues.length === 0,
      message: issues.length === 0 ? 
        `${accuracyChecks.length}件の月次データの整合性確認が完了しました` : 
        `${issues.length}件の整合性問題が見つかりました`,
      dailySummaryRows: dailyData.length,
      monthlySummaryRows: monthlyData.length,
      accuracyChecks: accuracyChecks,
      issues: issues,
      missingMonthlyEntries: missingMonthlyEntries
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      dailySummaryRows: 0,
      monthlySummaryRows: 0
    };
  }
} 

/**
 * testSummaryDataAccuracy関数のテスト実行
 */
function testSummaryDataAccuracyTest() {
  try {
    Logger.log('=== testSummaryDataAccuracy関数テスト開始 ===');
    
    const result = testSummaryDataAccuracy();
    
    Logger.log('テスト結果:');
    Logger.log(`成功: ${result.success}`);
    Logger.log(`メッセージ: ${result.message}`);
    Logger.log(`日次サマリー行数: ${result.dailySummaryRows}`);
    Logger.log(`月次サマリー行数: ${result.monthlySummaryRows}`);
    
    if (result.accuracyChecks) {
      Logger.log(`精度チェック件数: ${result.accuracyChecks.length}`);
      result.accuracyChecks.forEach((check, index) => {
        Logger.log(`チェック${index + 1}: ${check.yearMonth} ${check.employeeId} - 問題: ${check.hasIssues ? 'あり' : 'なし'}`);
      });
    }
    
    if (result.issues && result.issues.length > 0) {
      Logger.log('発見された問題:');
      result.issues.forEach((issue, index) => {
        Logger.log(`${index + 1}. ${issue}`);
      });
    }
    
    if (result.missingMonthlyEntries && result.missingMonthlyEntries.length > 0) {
      Logger.log('月次サマリーに存在しない日次データ:');
      result.missingMonthlyEntries.forEach((entry, index) => {
        Logger.log(`${index + 1}. ${entry}`);
      });
    }
    
    Logger.log('=== testSummaryDataAccuracy関数テスト完了 ===');
    
    return {
      success: true,
      testResult: result,
      message: 'テストが正常に完了しました'
    };
    
  } catch (error) {
    Logger.log(`テスト実行エラー: ${error.message}`);
    return {
      success: false,
      error: error.message,
      message: 'テスト実行中にエラーが発生しました'
    };
  }
} 

/**
 * testRequestResponsesIntegrity関数のテスト実行
 */
function testRequestResponsesIntegrityTest() {
  try {
    Logger.log('=== testRequestResponsesIntegrity関数テスト開始 ===');
    
    const result = testRequestResponsesIntegrity();
    
    Logger.log('テスト結果:');
    Logger.log(`成功: ${result.success}`);
    Logger.log(`検証ルールあり: ${result.hasValidation}`);
    
    if (result.error) {
      Logger.log(`エラー: ${result.error}`);
    }
    
    if (result.statusColumnIndex) {
      Logger.log(`Status列インデックス: ${result.statusColumnIndex}`);
      Logger.log(`Status列文字: ${result.statusColumnLetter}`);
      Logger.log(`検証範囲: ${result.validationRange}`);
      Logger.log(`チェック行数: ${result.totalRowsChecked}`);
    }
    
    if (result.validationRule) {
      Logger.log(`検証ルール: ${result.validationRule}`);
    }
    
    if (result.headerRow) {
      Logger.log(`ヘッダー行: ${result.headerRow.join(', ')}`);
    }
    
    if (result.validationDetails && result.validationDetails.length > 0) {
      Logger.log(`詳細検証結果: ${result.validationDetails.length}行をチェック`);
      const validRows = result.validationDetails.filter(detail => detail.hasValidation).length;
      const invalidRows = result.validationDetails.filter(detail => !detail.hasValidation).length;
      Logger.log(`検証ルールあり: ${validRows}行, 検証ルールなし: ${invalidRows}行`);
      
      // 最初の5行の詳細を表示
      result.validationDetails.slice(0, 5).forEach((detail, index) => {
        Logger.log(`行${detail.row}: 検証ルール${detail.hasValidation ? 'あり' : 'なし'}`);
        if (detail.hasValidation && detail.helpText) {
          Logger.log(`  ヘルプテキスト: ${detail.helpText}`);
        }
      });
    }
    
    Logger.log('=== testRequestResponsesIntegrity関数テスト完了 ===');
    
    return {
      success: true,
      testResult: result,
      message: 'テストが正常に完了しました'
    };
    
  } catch (error) {
    Logger.log(`テスト実行エラー: ${error.message}`);
    return {
      success: false,
      error: error.message,
      message: 'テスト実行中にエラーが発生しました'
    };
  }
} 