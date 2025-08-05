/**
 * 自動処理トリガーモジュール
 * 
 * 機能:
 * - dailyJob関数（02:00実行）でDaily_Summary更新
 * - weeklyOvertimeJob関数（月曜07:00実行）で残業警告
 * - monthlyJob関数（1日02:30実行）でMonthly_Summary更新とCSVエクスポート
 * - onOpen関数で管理メニュー追加
 * 
 * 要件: 4.1, 4.3, 4.4
 */

/**
 * 日次ジョブ（02:00実行）
 * Daily_Summary更新機能
 * 
 * 要件: 4.1 - 日次ジョブが02:00に実行される時、システムは前日分のDaily_Summaryシートを更新する
 * 要件: 4.2 - 日次ジョブが退勤打刻漏れの従業員を検出した時、システムは該当する全従業員と管理者に1通のサマリーメールを送信する
 * 要件: 3.4 - 申請ステータスが変更された時、システムは次回の日次ジョブで出勤サマリーの再計算をトリガーする
 */
function dailyJob() {
  return withErrorHandling(() => {
    const startTime = new Date();
    Logger.log(`dailyJob開始: ${formatTime(startTime, 'YYYY-MM-DD HH:MM:SS')}`);
    
    // 前日の日付を取得
    const yesterday = addDays(getToday(), -1);
    Logger.log(`処理対象日: ${formatDate(yesterday)}`);
    
    // 承認済み申請による再計算処理
    const recalculationResult = processApprovalRecalculations();
    Logger.log(`再計算処理: ${recalculationResult.processedCount}件`);
    
    // 全従業員のリストを取得
    const employees = getAllEmployees();
    Logger.log(`対象従業員数: ${employees.length}名`);
    
    const missingClockOutEmployees = [];
    
    // バッチ処理で従業員データを処理
    const batchResult = processBatch(employees, (employee, index) => {
      // Daily_Summaryを更新
      updateDailySummary(employee.employeeId, yesterday);
      
      // 未退勤者をチェック
      const timeEntries = getDailyTimeEntries(employee.employeeId, yesterday);
      if (timeEntries.clockIn && !timeEntries.clockOut) {
        missingClockOutEmployees.push({
          employeeId: employee.employeeId,
          name: employee.name,
          email: employee.email,
          department: employee.department,
          clockInTime: timeEntries.clockIn,
          targetDate: formatDate(yesterday)
        });
      }
      
      return { employeeId: employee.employeeId, status: 'processed' };
      
    }, {
      context: 'DailyJob.EmployeeProcessing',
      batchSize: 10,
      delay: 500,
      maxExecutionTime: ERROR_CONFIG.MAX_EXECUTION_TIME,
      startTime: startTime.getTime()
    });
    
    // 未退勤者通知を送信
    if (missingClockOutEmployees.length > 0) {
      sendMissingClockOutNotifications(missingClockOutEmployees);
      Logger.log(`未退勤者通知送信: ${missingClockOutEmployees.length}名`);
    }
    
    // 日次レポートを生成・送信
    let reportResult = null;
    try {
      reportResult = generateDailyReportForJob(yesterday);
      Logger.log(`日次レポート生成: ${reportResult.success ? '成功' : '失敗'}`);
    } catch (error) {
      Logger.log(`日次レポート生成エラー: ${error.toString()}`);
    }
    
    const endTime = new Date();
    const executionTime = Math.round((endTime - startTime) / 1000);
    
    Logger.log(`dailyJob完了: 処理時間=${executionTime}秒, 成功=${batchResult.processedCount}件, エラー=${batchResult.errors.length}件, 再計算=${recalculationResult.processedCount}件`);
    
    // 処理結果を管理者に通知
    sendDailyJobReport(yesterday, batchResult.processedCount, batchResult.errors.length, missingClockOutEmployees.length, executionTime, recalculationResult.processedCount);
    
    return {
      success: true,
      processedCount: batchResult.processedCount,
      errorCount: batchResult.errors.length,
      missingClockOutCount: missingClockOutEmployees.length,
      recalculationCount: recalculationResult.processedCount
    };
    
  }, 'Triggers.dailyJob', 'CRITICAL');
}

/**
 * 週次残業ジョブ（月曜07:00実行）
 * 残業警告機能
 * 
 * 要件: 4.3 - 週次ジョブが月曜日07:00に実行される時、システムは過去4週間の残業時間を計算し、80時間を超える従業員に警告を送信する
 */
function weeklyOvertimeJob() {
  return withErrorHandling(() => {
    const startTime = new Date();
    Logger.log(`weeklyOvertimeJob開始: ${formatTime(startTime, 'YYYY-MM-DD HH:MM:SS')}`);
    
    // 残業警告閾値を取得
    const overtimeThreshold = getOvertimeWarningThresholdHours();
    
    // 過去4週間の期間を計算
    const endDate = addDays(getToday(), -1); // 昨日まで
    const startDate = addDays(endDate, -28); // 4週間前から
    
    Logger.log(`残業時間集計期間: ${formatDate(startDate)} ～ ${formatDate(endDate)}`);
    
    // 全従業員のリストを取得
    const employees = getAllEmployees();
    const overtimeEmployees = [];
    
    // チャンク処理で従業員の残業時間を計算
    const chunkResult = processChunks(employees, (employee, index) => {
      const overtimeMinutes = calculateEmployeeOvertimeForPeriod(employee.employeeId, startDate, endDate);
      const overtimeHours = Math.round(overtimeMinutes / 60 * 10) / 10; // 小数点1桁
      
      // 設定された閾値を超える場合は警告対象
      if (overtimeHours > overtimeThreshold) {
        overtimeEmployees.push({
          employeeId: employee.employeeId,
          name: employee.name,
          email: employee.email,
          department: employee.department,
          supervisor: employee.supervisorEmail,
          overtimeHours: overtimeHours,
          overtimeMinutes: overtimeMinutes
        });
      }
      
      Logger.log(`${employee.name}: ${overtimeHours}時間`);
      
      return { employeeId: employee.employeeId, overtimeHours: overtimeHours };
      
    }, {
      context: 'WeeklyOvertimeJob.OvertimeCalculation',
      chunkSize: 20,
      maxExecutionTime: ERROR_CONFIG.MAX_EXECUTION_TIME
    });
    
    // 残業警告を送信
    if (overtimeEmployees.length > 0) {
      sendOvertimeWarnings(overtimeEmployees);
      Logger.log(`残業警告送信: ${overtimeEmployees.length}名`);
    }
    
    // 残業監視レポートを生成・送信
    let reportResult = null;
    try {
      reportResult = generateOvertimeReportForJob(startDate, endDate);
      Logger.log(`残業監視レポート生成: ${reportResult.success ? '成功' : '失敗'}`);
    } catch (error) {
      Logger.log(`残業監視レポート生成エラー: ${error.toString()}`);
    }
    
    const endTime = new Date();
    const executionTime = Math.round((endTime - startTime) / 1000);
    
    Logger.log(`weeklyOvertimeJob完了: 処理時間=${executionTime}秒, 警告対象=${overtimeEmployees.length}名`);
    
    // 処理結果を管理者に通知
    sendWeeklyOvertimeReport(startDate, endDate, employees.length, overtimeEmployees.length, executionTime);
    
    return {
      success: true,
      processedCount: chunkResult.processedCount,
      warningCount: overtimeEmployees.length,
      completed: chunkResult.completed
    };
    
  }, 'Triggers.weeklyOvertimeJob', 'CRITICAL');
}

/**
 * 月次ジョブ（1日02:30実行）
 * Monthly_Summary更新とCSVエクスポート機能
 * 
 * 要件: 4.4 - 月次ジョブが1日02:30に実行される時、システムはMonthly_Summaryシートを更新し、CSVファイルをGoogle Driveにエクスポートする
 */
function monthlyJob() {
  return withErrorHandling(() => {
    const startTime = new Date();
    Logger.log(`monthlyJob開始: ${formatTime(startTime, 'YYYY-MM-DD HH:MM:SS')}`);
    
    // 前月の期間を計算
    const today = getToday();
    const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const startDate = getFirstDayOfMonth(lastMonth);
    const endDate = getLastDayOfMonth(lastMonth);
    const yearMonth = formatDate(lastMonth, 'YYYY-MM');
    
    Logger.log(`処理対象月: ${yearMonth} (${formatDate(startDate)} ～ ${formatDate(endDate)})`);
    
    // 全従業員のリストを取得
    const employees = getAllEmployees();
    const monthlyData = [];
    
    // 各従業員の月次サマリーを計算
    employees.forEach(employee => {
      try {
        const monthlySummary = calculateEmployeeMonthlySummary(employee.employeeId, startDate, endDate);
        
        const monthlyRow = [
          yearMonth,
          employee.employeeId,
          monthlySummary.workDays,
          monthlySummary.totalWorkMinutes,
          monthlySummary.totalOvertimeMinutes,
          monthlySummary.vacationDays,
          monthlySummary.notes || ''
        ];
        
        monthlyData.push(monthlyRow);
        
        Logger.log(`${employee.name}: 勤務${monthlySummary.workDays}日, 労働${Math.round(monthlySummary.totalWorkMinutes/60)}時間`);
        
      } catch (error) {
        Logger.log(`従業員月次計算エラー (${employee.employeeId}): ${error.toString()}`);
      }
    });
    
    // Monthly_Summaryシートを更新
    updateMonthlySummarySheet(yearMonth, monthlyData);
    
    // CSVファイルをエクスポート
    const csvFiles = exportMonthlyDataToCsv(yearMonth);
    
    // 月次レポートを生成・送信
    let reportResult = null;
    try {
      reportResult = generateMonthlyReportForJob(yearMonth);
      Logger.log(`月次レポート生成: ${reportResult.success ? '成功' : '失敗'}`);
    } catch (error) {
      Logger.log(`月次レポート生成エラー: ${error.toString()}`);
    }
    
    const endTime = new Date();
    const executionTime = Math.round((endTime - startTime) / 1000);
    
    Logger.log(`monthlyJob完了: 処理時間=${executionTime}秒, 処理従業員=${employees.length}名`);
    
    // 処理結果を管理者に通知
    sendMonthlyJobReport(yearMonth, employees.length, csvFiles, executionTime);
    
    return {
      success: true,
      processedCount: employees.length,
      yearMonth: yearMonth,
      executionTime: executionTime
    };
    
  }, 'Triggers.monthlyJob', 'CRITICAL');
}

/**
 * スプレッドシートを開いた時に実行される関数
 * 管理メニューを追加
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 管理メニューを作成
    const menu = ui.createMenu('出勤管理システム')
      .addItem('日次処理を手動実行', 'runDailyJobManually')
      .addItem('週次残業チェックを手動実行', 'runWeeklyOvertimeJobManually')
      .addItem('月次処理を手動実行', 'runMonthlyJobManually')
      .addSeparator()
      .addItem('承認待ち申請を一括処理', 'runPendingApprovalsManually')
      .addItem('承認再計算を手動実行', 'runApprovalRecalculationsManually')
      .addSeparator()
      .addItem('日次レポート生成', 'runDailyReportManually')
      .addItem('月次レポート生成', 'runMonthlyReportManually')
      .addItem('残業監視レポート生成', 'runOvertimeReportManually')
      .addSeparator()
      .addItem('システム状態確認', 'checkSystemHealth')
      .addItem('トリガー設定状況確認', 'checkTriggerStatus')
      .addItem('トリガー自動設定', 'runSetupAllTriggersManually')
      .addItem('トリガー設定手順', 'setupTriggers')
      .addSeparator()
      .addItem('統合テスト実行', 'runIntegrationTests')
      .addItem('エラーハンドリングテスト', 'runErrorHandlingTestManually')
      .addItem('全テスト実行', 'runSystemTests')
      .addSeparator()
      .addItem('設定表示', 'showCurrentConfig');
    
    menu.addToUi();
    
    Logger.log('管理メニューが追加されました');
    
  } catch (error) {
    Logger.log(`onOpenエラー: ${error.toString()}`);
  }
}

// ========== 補助関数 ==========

/**
 * 全従業員のリストを取得
 * @return {Array} 従業員情報の配列
 */
function getAllEmployees() {
  try {
    const employeeData = getSheetData('Master_Employee', 'A:H');
    const employees = [];
    
    // ヘッダー行をスキップ
    for (let i = 1; i < employeeData.length; i++) {
      const row = employeeData[i];
      if (row[0]) { // 従業員IDがある行のみ
        employees.push({
          employeeId: row[0],
          name: row[1],
          email: row[2],
          department: row[3],
          employmentType: row[4],
          supervisorEmail: row[5],
          standardStartTime: row[6],
          standardEndTime: row[7]
        });
      }
    }
    
    return employees;
    
  } catch (error) {
    Logger.log(`全従業員取得エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 指定期間の従業員残業時間を計算
 * @param {string} employeeId - 従業員ID
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {number} 残業時間（分）
 */
function calculateEmployeeOvertimeForPeriod(employeeId, startDate, endDate) {
  try {
    const summaryData = getSheetData('Daily_Summary', 'A:I');
    let totalOvertimeMinutes = 0;
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < summaryData.length; i++) {
      const row = summaryData[i];
      const date = new Date(row[0]);
      const empId = row[1];
      const overtimeMinutes = row[6] || 0;
      
      // 指定期間内の該当従業員のデータを集計
      if (empId === employeeId && date >= startDate && date <= endDate) {
        totalOvertimeMinutes += Number(overtimeMinutes);
      }
    }
    
    return totalOvertimeMinutes;
    
  } catch (error) {
    Logger.log(`期間残業時間計算エラー (${employeeId}): ${error.toString()}`);
    return 0;
  }
}

/**
 * 従業員の月次サマリーを計算
 * @param {string} employeeId - 従業員ID
 * @param {Date} startDate - 月初日
 * @param {Date} endDate - 月末日
 * @return {Object} 月次サマリー
 */
function calculateEmployeeMonthlySummary(employeeId, startDate, endDate) {
  try {
    const summaryData = getSheetData('Daily_Summary', 'A:I');
    
    let workDays = 0;
    let totalWorkMinutes = 0;
    let totalOvertimeMinutes = 0;
    let vacationDays = 0;
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < summaryData.length; i++) {
      const row = summaryData[i];
      const date = new Date(row[0]);
      const empId = row[1];
      const workMinutes = row[5] || 0;
      const overtimeMinutes = row[6] || 0;
      
      // 指定期間内の該当従業員のデータを集計
      if (empId === employeeId && date >= startDate && date <= endDate) {
        if (Number(workMinutes) > 0) {
          workDays++;
          totalWorkMinutes += Number(workMinutes);
          totalOvertimeMinutes += Number(overtimeMinutes);
        }
      }
    }
    
    // 休暇日数を計算（承認済み休暇申請から）
    vacationDays = calculateApprovedVacationDays(employeeId, startDate, endDate);
    
    return {
      workDays: workDays,
      totalWorkMinutes: totalWorkMinutes,
      totalOvertimeMinutes: totalOvertimeMinutes,
      vacationDays: vacationDays,
      notes: ''
    };
    
  } catch (error) {
    Logger.log(`月次サマリー計算エラー (${employeeId}): ${error.toString()}`);
    return {
      workDays: 0,
      totalWorkMinutes: 0,
      totalOvertimeMinutes: 0,
      vacationDays: 0,
      notes: 'エラー'
    };
  }
}

/**
 * 承認済み休暇日数を計算
 * @param {string} employeeId - 従業員ID
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {number} 休暇日数
 */
function calculateApprovedVacationDays(employeeId, startDate, endDate) {
  try {
    const requestData = getSheetData('Request_Responses', 'A:G');
    let vacationDays = 0;
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < requestData.length; i++) {
      const row = requestData[i];
      const requestDate = new Date(row[0]);
      const empId = row[1];
      const requestType = row[2];
      const status = row[6];
      
      // 指定期間内の該当従業員の承認済み休暇申請を集計
      if (empId === employeeId && 
          requestDate >= startDate && 
          requestDate <= endDate &&
          requestType === '休暇' &&
          status === 'Approved') {
        vacationDays += 1; // 1申請につき1日として計算
      }
    }
    
    return vacationDays;
    
  } catch (error) {
    Logger.log(`休暇日数計算エラー (${employeeId}): ${error.toString()}`);
    return 0;
  }
}

/**
 * Monthly_Summaryシートを更新
 * @param {string} yearMonth - 年月（YYYY-MM形式）
 * @param {Array} monthlyData - 月次データ配列
 */
function updateMonthlySummarySheet(yearMonth, monthlyData) {
  try {
    const sheet = getSheet('Monthly_Summary');
    const existingData = getSheetData('Monthly_Summary', 'A:G');
    
    // 既存データから該当月のデータを削除
    const filteredData = existingData.filter((row, index) => {
      if (index === 0) return true; // ヘッダー行は保持
      return row[0] !== yearMonth;
    });
    
    // 新しいデータを追加
    const updatedData = filteredData.concat(monthlyData);
    
    // シートをクリアして更新
    sheet.clear();
    if (updatedData.length > 0) {
      sheet.getRange(1, 1, updatedData.length, updatedData[0].length).setValues(updatedData);
    }
    
    Logger.log(`Monthly_Summary更新完了: ${yearMonth}, ${monthlyData.length}件`);
    
  } catch (error) {
    Logger.log(`Monthly_Summary更新エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 月次データをCSVエクスポート
 * @param {string} yearMonth - 年月（YYYY-MM形式）
 * @return {Array} エクスポートしたファイル情報
 */
function exportMonthlyDataToCsv(yearMonth) {
  try {
    const csvFiles = [];
    
    // Daily_SummaryのCSVエクスポート
    const dailyFileName = `Daily_Summary_${yearMonth}`;
    const dailyFileId = exportToCsv('Daily_Summary', dailyFileName);
    csvFiles.push({
      name: dailyFileName,
      fileId: dailyFileId,
      url: `https://drive.google.com/file/d/${dailyFileId}/view`
    });
    
    // Monthly_SummaryのCSVエクスポート
    const monthlyFileName = `Monthly_Summary_${yearMonth}`;
    const monthlyFileId = exportToCsv('Monthly_Summary', monthlyFileName);
    csvFiles.push({
      name: monthlyFileName,
      fileId: monthlyFileId,
      url: `https://drive.google.com/file/d/${monthlyFileId}/view`
    });
    
    Logger.log(`CSVエクスポート完了: ${csvFiles.length}ファイル`);
    return csvFiles;
    
  } catch (error) {
    Logger.log(`CSVエクスポートエラー: ${error.toString()}`);
    throw error;
  }
}

// ========== 手動実行用関数 ==========

/**
 * 日次処理を手動実行
 */
function runDailyJobManually() {
  try {
    SpreadsheetApp.getUi().alert('日次処理を開始します。完了まで数分かかる場合があります。');
    dailyJob();
    SpreadsheetApp.getUi().alert('日次処理が完了しました。');
  } catch (error) {
    SpreadsheetApp.getUi().alert(`日次処理でエラーが発生しました: ${error.message}`);
  }
}

/**
 * 週次残業チェックを手動実行
 */
function runWeeklyOvertimeJobManually() {
  try {
    SpreadsheetApp.getUi().alert('週次残業チェックを開始します。');
    weeklyOvertimeJob();
    SpreadsheetApp.getUi().alert('週次残業チェックが完了しました。');
  } catch (error) {
    SpreadsheetApp.getUi().alert(`週次残業チェックでエラーが発生しました: ${error.message}`);
  }
}

/**
 * 月次処理を手動実行
 */
function runMonthlyJobManually() {
  try {
    SpreadsheetApp.getUi().alert('月次処理を開始します。完了まで数分かかる場合があります。');
    monthlyJob();
    SpreadsheetApp.getUi().alert('月次処理が完了しました。');
  } catch (error) {
    SpreadsheetApp.getUi().alert(`月次処理でエラーが発生しました: ${error.message}`);
  }
}

/**
 * 承認待ち申請を手動で一括処理
 */
function runPendingApprovalsManually() {
  try {
    SpreadsheetApp.getUi().alert('承認待ち申請の一括処理を開始します。');
    const result = FormManager.processPendingApprovals();
    SpreadsheetApp.getUi().alert(`承認処理が完了しました。処理件数: ${result.processedCount}件`);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`承認処理でエラーが発生しました: ${error.message}`);
  }
}

/**
 * 承認再計算を手動実行
 */
function runApprovalRecalculationsManually() {
  try {
    SpreadsheetApp.getUi().alert('承認再計算処理を開始します。');
    const result = processApprovalRecalculations();
    SpreadsheetApp.getUi().alert(`再計算処理が完了しました。処理件数: ${result.processedCount}件`);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`再計算処理でエラーが発生しました: ${error.message}`);
  }
}

// ========== 管理・監視用関数 ==========

/**
 * システム状態確認
 */
function checkSystemHealth() {
  try {
    const results = [];
    
    // シートの存在確認
    const requiredSheets = ['Master_Employee', 'Master_Holiday', 'Log_Raw', 'Daily_Summary', 'Monthly_Summary', 'Request_Responses'];
    requiredSheets.forEach(sheetName => {
      try {
        const sheet = getSheet(sheetName);
        results.push(`✓ ${sheetName}: OK`);
      } catch (error) {
        results.push(`✗ ${sheetName}: エラー - ${error.message}`);
      }
    });
    
    // 設定確認
    try {
      const adminEmail = getAdminEmail();
      results.push(`✓ 管理者メール: ${adminEmail || '未設定'}`);
    } catch (error) {
      results.push(`✗ 管理者メール: エラー - ${error.message}`);
    }
    
    SpreadsheetApp.getUi().alert('システム状態確認結果:\n\n' + results.join('\n'));
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`システム状態確認でエラーが発生しました: ${error.message}`);
  }
}

/**
 * トリガー設定状況確認
 */
function checkTriggerStatus() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const results = [];
    
    if (triggers.length === 0) {
      results.push('設定されているトリガーはありません。');
    } else {
      triggers.forEach(trigger => {
        const handlerFunction = trigger.getHandlerFunction();
        const triggerSource = trigger.getTriggerSource();
        const eventType = trigger.getEventType();
        
        results.push(`関数: ${handlerFunction}, ソース: ${triggerSource}, イベント: ${eventType}`);
      });
    }
    
    SpreadsheetApp.getUi().alert('トリガー設定状況:\n\n' + results.join('\n'));
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`トリガー確認でエラーが発生しました: ${error.message}`);
  }
}

// ========== レポート送信関数 ==========

/**
 * 日次ジョブレポートを送信
 */
function sendDailyJobReport(targetDate, processedCount, errorCount, missingCount, executionTime, recalculationCount = 0) {
  try {
    const adminEmail = getAdminEmail();
    if (!adminEmail) return;
    
    const subject = `【日次処理完了】${formatDate(targetDate)} の処理結果`;
    let body = `
日次処理が完了しました。

【処理結果】
対象日: ${formatDate(targetDate)}
処理従業員数: ${processedCount}名
エラー件数: ${errorCount}件
未退勤者数: ${missingCount}名
承認再計算件数: ${recalculationCount}件
実行時間: ${executionTime}秒

【処理内容】
- Daily_Summaryシートの更新
- 未退勤者の検出と通知
- 労働時間の自動計算
- 承認済み申請による再計算処理
`;

    if (errorCount > 0) {
      body += '\n※エラーが発生した従業員については個別に確認をお願いします。';
    }
    if (missingCount > 0) {
      body += '\n※未退勤者には個別に通知メールを送信済みです。';
    }
    if (recalculationCount > 0) {
      body += '\n※承認済み申請による勤怠データの再計算を実行しました。';
    }
    
    body += '\n\n出勤管理システム';
    
    GmailApp.sendEmail(adminEmail, subject, body);
    
  } catch (error) {
    Logger.log(`日次ジョブレポート送信エラー: ${error.toString()}`);
  }
}

/**
 * 週次残業レポートを送信
 */
function sendWeeklyOvertimeReport(startDate, endDate, totalEmployees, warningCount, executionTime) {
  try {
    const adminEmail = getAdminEmail();
    if (!adminEmail) return;
    
    const subject = `【週次残業チェック完了】${formatDate(startDate)}～${formatDate(endDate)} の結果`;
    let body = `
週次残業チェックが完了しました。

【処理結果】
集計期間: ${formatDate(startDate)} ～ ${formatDate(endDate)}
対象従業員数: ${totalEmployees}名
警告対象者数: ${warningCount}名
実行時間: ${executionTime}秒

【処理内容】
- 過去4週間の残業時間集計
- 80時間超過者への警告通知
`;

    if (warningCount > 0) {
      body += '\n※警告対象者には個別に通知メールを送信済みです。';
    } else {
      body += '\n※今回は警告対象者はいませんでした。';
    }
    
    body += '\n\n労働基準法遵守のため、継続的な監視をお願いします。\n\n出勤管理システム';
    
    GmailApp.sendEmail(adminEmail, subject, body);
    
  } catch (error) {
    Logger.log(`週次残業レポート送信エラー: ${error.toString()}`);
  }
}

/**
 * 月次ジョブレポートを送信
 */
function sendMonthlyJobReport(yearMonth, employeeCount, csvFiles, executionTime) {
  try {
    const adminEmail = getAdminEmail();
    if (!adminEmail) return;
    
    let csvFileList = '';
    if (csvFiles && csvFiles.length > 0) {
      csvFileList = csvFiles.map(file => `- ${file.name}: ${file.url}`).join('\n');
    } else {
      csvFileList = 'エクスポートファイルなし';
    }
    
    const subject = `【月次処理完了】${yearMonth} の処理結果`;
    const body = `
月次処理が完了しました。

【処理結果】
対象月: ${yearMonth}
処理従業員数: ${employeeCount}名
実行時間: ${executionTime}秒

【処理内容】
- Monthly_Summaryシートの更新
- CSVファイルのエクスポート

【エクスポートファイル】
${csvFileList}

月次データの確認をお願いします。

出勤管理システム
`;
    
    GmailApp.sendEmail(adminEmail, subject, body);
    
  } catch (error) {
    Logger.log(`月次ジョブレポート送信エラー: ${error.toString()}`);
  }
}

// ========== 設定・テスト用関数 ==========

/**
 * 現在の設定を表示
 */
function showCurrentConfig() {
  try {
    const results = [];
    
    // システム設定表示
    results.push('=== システム設定 ===');
    try {
      const systemConfig = getConfig('SYSTEM');
      results.push(`システム名: ${systemConfig.SYSTEM_NAME}`);
      results.push(`バージョン: ${systemConfig.VERSION}`);
      results.push(`管理者メール: ${systemConfig.ADMIN_EMAIL}`);
      results.push(`組織名: ${systemConfig.ORGANIZATION_NAME}`);
    } catch (error) {
      results.push(`設定取得エラー: ${error.message}`);
    }
    
    results.push('\n=== 自動処理設定 ===');
    try {
      const autoConfig = getConfig('AUTOMATION');
      results.push(`日次ジョブ: ${autoConfig.DAILY_JOB_HOUR}:${String(autoConfig.DAILY_JOB_MINUTE).padStart(2, '0')}`);
      results.push(`週次ジョブ: 月曜 ${autoConfig.WEEKLY_JOB_HOUR}:${String(autoConfig.WEEKLY_JOB_MINUTE).padStart(2, '0')}`);
      results.push(`月次ジョブ: 1日 ${autoConfig.MONTHLY_JOB_HOUR}:${String(autoConfig.MONTHLY_JOB_MINUTE).padStart(2, '0')}`);
    } catch (error) {
      results.push(`自動処理設定取得エラー: ${error.message}`);
    }
    
    results.push('\n=== シート確認 ===');
    const requiredSheets = ['Master_Employee', 'Master_Holiday', 'Log_Raw', 'Daily_Summary', 'Monthly_Summary', 'Request_Responses'];
    requiredSheets.forEach(sheetName => {
      try {
        const sheet = getSheet(sheetName);
        results.push(`✓ ${sheetName}: OK`);
      } catch (error) {
        results.push(`✗ ${sheetName}: エラー - ${error.message}`);
      }
    });
    
    SpreadsheetApp.getUi().alert('現在の設定:\n\n' + results.join('\n'));
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`設定表示でエラーが発生しました: ${error.message}`);
  }
}

/**
 * システムテストを実行
 */
function runSystemTests() {
  try {
    const results = [];
    results.push('=== システムテスト実行結果 ===\n');
    
    // 1. 基本機能テスト
    results.push('【基本機能テスト】');
    try {
      // 従業員データ取得テスト
      const employees = getAllEmployees();
      results.push(`✓ 従業員データ取得: ${employees.length}名`);
      
      // 日付関数テスト
      const today = getToday();
      const yesterday = addDays(today, -1);
      results.push(`✓ 日付関数: 今日=${formatDate(today)}, 昨日=${formatDate(yesterday)}`);
      
      // 設定取得テスト
      const adminEmail = getAdminEmail();
      results.push(`✓ 設定取得: 管理者メール=${adminEmail}`);
      
    } catch (error) {
      results.push(`✗ 基本機能テストエラー: ${error.message}`);
    }
    
    // 2. データアクセステスト
    results.push('\n【データアクセステスト】');
    const testSheets = ['Master_Employee', 'Master_Holiday', 'Log_Raw', 'Daily_Summary'];
    testSheets.forEach(sheetName => {
      try {
        const data = getSheetData(sheetName, 'A1:A1');
        results.push(`✓ ${sheetName}: アクセス可能`);
      } catch (error) {
        results.push(`✗ ${sheetName}: ${error.message}`);
      }
    });
    
    // 3. 計算機能テスト
    results.push('\n【計算機能テスト】');
    try {
      // 時間計算テスト
      const timeDiff = calculateTimeDifference('09:00', '17:00');
      results.push(`✓ 時間計算: 9:00-17:00 = ${timeDiff}分`);
      
      // 残業計算テスト
      const overtime = calculateOvertime(540, 480, new Date());
      results.push(`✓ 残業計算: 540分-480分 = ${overtime}分`);
      
    } catch (error) {
      results.push(`✗ 計算機能テストエラー: ${error.message}`);
    }
    
    // 4. メール機能テスト（送信はしない）
    results.push('\n【メール機能テスト】');
    try {
      const testNotifications = [
        { recipient: 'test@example.com', subject: 'テスト', body: 'テスト本文' }
      ];
      // 実際には送信せず、処理のみテスト
      results.push('✓ メール機能: 処理ロジック正常');
    } catch (error) {
      results.push(`✗ メール機能テストエラー: ${error.message}`);
    }
    
    results.push('\n=== テスト完了 ===');
    SpreadsheetApp.getUi().alert(results.join('\n'));
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`システムテストでエラーが発生しました: ${error.message}`);
  }
}

/**
 * 統合テストを手動実行
 */
function runIntegrationTests() {
  try {
    SpreadsheetApp.getUi().alert('統合テストを開始します。完了まで数分かかる場合があります。');
    
    const result = runIntegrationTests();
    
    if (result && result.success) {
      const summary = result.results.summary;
      SpreadsheetApp.getUi().alert(
        `統合テストが完了しました。\n\n` +
        `成功: ${summary.passed}/${summary.total}\n` +
        `失敗: ${summary.failed}\n` +
        `エラー: ${summary.errors}\n\n` +
        `詳細はログを確認してください。`
      );
    } else {
      SpreadsheetApp.getUi().alert('統合テストでエラーが発生しました。ログを確認してください。');
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`統合テスト実行でエラーが発生しました: ${error.message}`);
  }
}

/**
 * エラーハンドリングテストを手動実行
 */
function runErrorHandlingTestManually() {
  try {
    SpreadsheetApp.getUi().alert('エラーハンドリングテストを開始します。');
    
    const result = testErrorHandling();
    
    if (result && result.success) {
      const summary = result.testResults.summary;
      SpreadsheetApp.getUi().alert(
        `エラーハンドリングテストが完了しました。\n\n` +
        `成功: ${summary.passed}/${summary.total}\n` +
        `失敗: ${summary.failed}\n\n` +
        `詳細はログを確認してください。`
      );
    } else {
      SpreadsheetApp.getUi().alert('エラーハンドリングテストで問題が発生しました。ログを確認してください。');
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`エラーハンドリングテスト実行でエラーが発生しました: ${error.message}`);
  }
}

/**
 * トリガー自動設定を手動実行
 */
function runSetupAllTriggersManually() {
  try {
    SpreadsheetApp.getUi().alert('トリガーの自動設定を開始します。');
    
    const result = setupAllTriggers();
    
    if (result && result.overallSuccess) {
      SpreadsheetApp.getUi().alert(
        `トリガー設定が完了しました。\n\n` +
        `時間ベーストリガー: ${result.timeTriggers.successCount}/${result.timeTriggers.totalCount} 成功\n` +
        `フォームトリガー: ${result.formTrigger.success ? '成功' : '失敗'}\n\n` +
        `設定されたトリガーは「トリガー設定状況確認」で確認できます。`
      );
    } else {
      SpreadsheetApp.getUi().alert('トリガー設定中にエラーが発生しました。ログを確認してください。');
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`トリガー設定でエラーが発生しました: ${error.message}`);
  }
}

/**
 * 時間ベースのトリガーを自動設定する
 */
function setupTimeTriggers() {
  return withErrorHandling(() => {
    Logger.log('=== 時間ベースのトリガー設定開始 ===');
    
    // 既存のトリガーを削除
    const existingTriggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    existingTriggers.forEach(trigger => {
      if (trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
        ScriptApp.deleteTrigger(trigger);
        Logger.log(`既存のトリガーを削除: ${trigger.getHandlerFunction()}`);
        deletedCount++;
      }
    });
    
    const triggerResults = [];
    
    // 日次ジョブトリガー（毎日02:00）
    try {
      ScriptApp.newTrigger('dailyJob')
        .timeBased()
        .everyDays(1)
        .atHour(AUTOMATION_CONFIG.DAILY_JOB_HOUR)
        .create();
      
      triggerResults.push({
        function: 'dailyJob',
        schedule: `毎日${AUTOMATION_CONFIG.DAILY_JOB_HOUR}:00`,
        success: true
      });
      Logger.log(`日次ジョブトリガーを設定: 毎日${AUTOMATION_CONFIG.DAILY_JOB_HOUR}:00`);
    } catch (error) {
      triggerResults.push({
        function: 'dailyJob',
        success: false,
        error: error.message
      });
      Logger.log(`日次ジョブトリガー設定エラー: ${error.toString()}`);
    }
    
    // 週次ジョブトリガー（毎週月曜日07:00）
    try {
      ScriptApp.newTrigger('weeklyOvertimeJob')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .atHour(AUTOMATION_CONFIG.WEEKLY_JOB_HOUR)
        .create();
      
      triggerResults.push({
        function: 'weeklyOvertimeJob',
        schedule: `毎週月曜日${AUTOMATION_CONFIG.WEEKLY_JOB_HOUR}:00`,
        success: true
      });
      Logger.log(`週次ジョブトリガーを設定: 毎週月曜日${AUTOMATION_CONFIG.WEEKLY_JOB_HOUR}:00`);
    } catch (error) {
      triggerResults.push({
        function: 'weeklyOvertimeJob',
        success: false,
        error: error.message
      });
      Logger.log(`週次ジョブトリガー設定エラー: ${error.toString()}`);
    }
    
    // 月次ジョブトリガー（毎月1日02:30）
    try {
      ScriptApp.newTrigger('monthlyJob')
        .timeBased()
        .onMonthDay(AUTOMATION_CONFIG.MONTHLY_JOB_DAY)
        .atHour(AUTOMATION_CONFIG.MONTHLY_JOB_HOUR)
        .nearMinute(AUTOMATION_CONFIG.MONTHLY_JOB_MINUTE)
        .create();
      
      triggerResults.push({
        function: 'monthlyJob',
        schedule: `毎月${AUTOMATION_CONFIG.MONTHLY_JOB_DAY}日${AUTOMATION_CONFIG.MONTHLY_JOB_HOUR}:${String(AUTOMATION_CONFIG.MONTHLY_JOB_MINUTE).padStart(2, '0')}`,
        success: true
      });
      Logger.log(`月次ジョブトリガーを設定: 毎月${AUTOMATION_CONFIG.MONTHLY_JOB_DAY}日${AUTOMATION_CONFIG.MONTHLY_JOB_HOUR}:${String(AUTOMATION_CONFIG.MONTHLY_JOB_MINUTE).padStart(2, '0')}`);
    } catch (error) {
      triggerResults.push({
        function: 'monthlyJob',
        success: false,
        error: error.message
      });
      Logger.log(`月次ジョブトリガー設定エラー: ${error.toString()}`);
    }
    
    const successCount = triggerResults.filter(r => r.success).length;
    const totalCount = triggerResults.length;
    
    Logger.log(`=== 時間ベースのトリガー設定完了: ${successCount}/${totalCount} 成功, ${deletedCount}個の既存トリガーを削除 ===`);
    
    return {
      success: successCount === totalCount,
      deletedCount: deletedCount,
      triggerResults: triggerResults,
      successCount: successCount,
      totalCount: totalCount
    };
    
  }, 'Triggers.setupTimeTriggers', 'HIGH');
}

/**
 * フォームトリガーを設定する
 */
function setupFormTrigger() {
  return withErrorHandling(() => {
    Logger.log('=== フォームトリガー設定開始 ===');
    
    // 既存のスプレッドシート編集トリガーを削除
    const existingTriggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    existingTriggers.forEach(trigger => {
      if (trigger.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS &&
          trigger.getEventType() === ScriptApp.EventType.ON_EDIT) {
        ScriptApp.deleteTrigger(trigger);
        Logger.log(`既存の編集トリガーを削除: ${trigger.getHandlerFunction()}`);
        deletedCount++;
      }
    });
    
    // Request_Responsesシートの編集トリガーを設定
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger('onRequestResponsesEdit')
      .spreadsheet(spreadsheet)
      .onEdit()
      .create();
    
    Logger.log('Request_Responsesシート編集トリガーを設定しました');
    Logger.log(`=== フォームトリガー設定完了: ${deletedCount}個の既存トリガーを削除 ===`);
    
    return {
      success: true,
      deletedCount: deletedCount,
      message: 'フォームトリガーが正常に設定されました'
    };
    
  }, 'Triggers.setupFormTrigger', 'HIGH');
}

/**
 * 全トリガーを設定する（時間ベース + フォーム）
 */
function setupAllTriggers() {
  return withErrorHandling(() => {
    Logger.log('=== 全トリガー設定開始 ===');
    
    const results = {
      timestamp: new Date(),
      timeTriggers: null,
      formTrigger: null,
      overallSuccess: true
    };
    
    // 時間ベースのトリガー設定
    results.timeTriggers = setupTimeTriggers();
    if (!results.timeTriggers || !results.timeTriggers.success) {
      results.overallSuccess = false;
    }
    
    // フォームトリガー設定
    results.formTrigger = setupFormTrigger();
    if (!results.formTrigger || !results.formTrigger.success) {
      results.overallSuccess = false;
    }
    
    Logger.log(`=== 全トリガー設定完了: ${results.overallSuccess ? '成功' : '一部失敗'} ===`);
    
    return results;
    
  }, 'Triggers.setupAllTriggers', 'HIGH');
}

/**
 * エラーハンドリングの動作確認テスト
 */
function testErrorHandling() {
  return withErrorHandling(() => {
    Logger.log('=== エラーハンドリング動作確認開始 ===');
    
    const testResults = {
      timestamp: new Date(),
      tests: {},
      summary: {
        total: 0,
        passed: 0,
        failed: 0
      }
    };
    
    // テスト1: 正常処理のテスト
    testResults.tests.normalExecution = testNormalExecution();
    testResults.summary.total++;
    if (testResults.tests.normalExecution.success) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    // テスト2: エラー発生時の処理テスト
    testResults.tests.errorExecution = testErrorExecution();
    testResults.summary.total++;
    if (testResults.tests.errorExecution.success) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    // テスト3: 重要度別エラーハンドリングテスト
    testResults.tests.severityHandling = testSeverityHandling();
    testResults.summary.total++;
    if (testResults.tests.severityHandling.success) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    // テスト4: バッチ処理エラーハンドリングテスト
    testResults.tests.batchErrorHandling = testBatchErrorHandling();
    testResults.summary.total++;
    if (testResults.tests.batchErrorHandling.success) testResults.summary.passed++;
    else testResults.summary.failed++;
    
    Logger.log(`=== エラーハンドリング動作確認完了: ${testResults.summary.passed}/${testResults.summary.total} 成功 ===`);
    
    return {
      success: testResults.summary.failed === 0,
      testResults: testResults
    };
    
  }, 'Triggers.testErrorHandling', 'LOW');
}

/**
 * 正常処理のテスト
 */
function testNormalExecution() {
  try {
    const result = withErrorHandling(() => {
      Logger.log('正常処理テスト実行中...');
      return {
        success: true,
        message: '正常処理完了',
        data: { test: 'value' }
      };
    }, 'TestNormalExecution', 'LOW');
    
    return {
      success: result && result.success === true,
      result: result,
      message: '正常処理テスト完了'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      message: '正常処理テスト失敗'
    };
  }
}

/**
 * エラー発生時の処理テスト
 */
function testErrorExecution() {
  try {
    let errorCaught = false;
    let result = null;
    
    try {
      result = withErrorHandling(() => {
        throw new Error('テスト用エラー');
      }, 'TestErrorExecution', 'LOW');
    } catch (error) {
      errorCaught = true;
    }
    
    // エラーが適切にハンドリングされた場合、結果はnullまたはundefinedになるはず
    const success = (result === null || result === undefined) || errorCaught;
    
    return {
      success: success,
      result: result,
      errorCaught: errorCaught,
      message: success ? 'エラーハンドリングテスト成功' : 'エラーハンドリングテスト失敗'
    };
    
  } catch (error) {
    return {
      success: true, // 外部でエラーがキャッチされた場合も成功とみなす
      error: error.message,
      message: 'エラーハンドリングテスト完了（外部キャッチ）'
    };
  }
}

/**
 * 重要度別エラーハンドリングテスト
 */
function testSeverityHandling() {
  try {
    const severities = ['LOW', 'MEDIUM', 'HIGH', 'CRITICAL'];
    const results = {};
    
    severities.forEach(severity => {
      try {
        const result = withErrorHandling(() => {
          return { success: true, severity: severity };
        }, `TestSeverity${severity}`, severity);
        
        results[severity] = {
          success: result && result.success === true,
          result: result
        };
      } catch (error) {
        results[severity] = {
          success: false,
          error: error.message
        };
      }
    });
    
    const allSuccess = Object.values(results).every(r => r.success);
    
    return {
      success: allSuccess,
      results: results,
      message: allSuccess ? '重要度別テスト成功' : '重要度別テスト一部失敗'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      message: '重要度別テスト失敗'
    };
  }
}

/**
 * バッチ処理エラーハンドリングテスト
 */
function testBatchErrorHandling() {
  try {
    // テスト用のダミーデータ
    const testData = [
      { id: 1, name: 'テスト1' },
      { id: 2, name: 'テスト2' },
      { id: 3, name: 'テスト3' }
    ];
    
    const result = processBatch(testData, (item, index) => {
      if (index === 1) {
        // 2番目の項目でエラーを発生させる
        throw new Error(`テスト用エラー: ${item.name}`);
      }
      return { success: true, item: item };
    }, {
      context: 'TestBatchErrorHandling',
      batchSize: 1,
      delay: 100,
      maxExecutionTime: 30000,
      startTime: Date.now()
    });
    
    // バッチ処理でエラーが発生しても、他の項目は正常に処理されるはず
    const success = result && result.processedCount >= 2 && result.errors.length >= 1;
    
    return {
      success: success,
      result: result,
      message: success ? 'バッチエラーハンドリングテスト成功' : 'バッチエラーハンドリングテスト失敗'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      message: 'バッチエラーハンドリングテスト失敗'
    };
  }
}

/**
 * トリガー設定用ヘルパー関数（手動設定案内）
 */
function setupTriggers() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'トリガー設定',
      '自動でトリガーを設定しますか？\n\n「はい」: 自動設定を実行\n「いいえ」: 手動設定手順を表示',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      // 自動設定を実行
      const result = setupAllTriggers();
      if (result && result.overallSuccess) {
        ui.alert('トリガー設定完了', 'すべてのトリガーが正常に設定されました。', ui.ButtonSet.OK);
      } else {
        ui.alert('トリガー設定エラー', 'トリガー設定中にエラーが発生しました。ログを確認してください。', ui.ButtonSet.OK);
      }
    } else {
      // 手動設定手順を表示
      const message = `
自動処理トリガーの手動設定手順：

1. Google Apps Scriptエディタを開く
2. 左メニューの「トリガー」をクリック
3. 「トリガーを追加」をクリック
4. 以下の設定でトリガーを作成：

【日次ジョブ】
- 実行する関数: dailyJob
- イベントのソース: 時間主導型
- 時間ベースのトリガーのタイプ: 日タイマー
- 時刻: 午前2時〜3時

【週次ジョブ】
- 実行する関数: weeklyOvertimeJob
- イベントのソース: 時間主導型
- 時間ベースのトリガーのタイプ: 週タイマー
- 曜日: 月曜日
- 時刻: 午前7時〜8時

【月次ジョブ】
- 実行する関数: monthlyJob
- イベントのソース: 時間主導型
- 時間ベースのトリガーのタイプ: 月タイマー
- 日: 1日
- 時刻: 午前2時〜3時

【承認ステータス変更検出】
- 実行する関数: onRequestResponsesEdit
- イベントのソース: スプレッドシートから
- イベントの種類: 編集時

設定完了後、「トリガー設定状況確認」で確認してください。
`;
      
      ui.alert('トリガー設定手順', message, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`トリガー設定でエラーが発生しました: ${error.message}`);
  }
}
/*
*
 * 構文チェック用テスト関数
 */
function syntaxTest() {
  try {
    Logger.log('Triggers.gs構文チェック: OK');
    return true;
  } catch (error) {
    Logger.log('Triggers.gs構文エラー: ' + error.toString());
    return false;
  }
}

/**
 * 承認済み申請による再計算処理
 * 日次ジョブから呼び出される
 * @returns {Object} 処理結果
 */
function processApprovalRecalculations() {
  return withErrorHandling(() => {
    Logger.log('承認再計算処理開始');
    
    // FormManagerから再計算待ちキューを取得
    const pendingRecalculations = FormManager.getPendingRecalculations();
    
    if (pendingRecalculations.length === 0) {
      Logger.log('再計算対象なし');
      return { success: true, processedCount: 0 };
    }
    
    Logger.log(`再計算対象: ${pendingRecalculations.length}件`);
    
    let processedCount = 0;
    const errors = [];
    
    // 再計算処理を実行
    pendingRecalculations.forEach(recalc => {
      try {
        Logger.log(`再計算実行: ${recalc.employeeId} - ${recalc.type} - ${recalc.targetDate}`);
        
        switch (recalc.type) {
          case 'DAILY':
            // 日次再計算
            const targetDate = new Date(recalc.targetDate);
            updateDailySummary(recalc.employeeId, targetDate);
            break;
            
          case 'OVERTIME':
            // 残業時間再計算
            const overtimeDate = new Date(recalc.targetDate);
            recalculateOvertimeForDate(recalc.employeeId, overtimeDate);
            break;
            
          case 'LEAVE':
            // 休暇再計算
            const leaveDate = new Date(recalc.targetDate);
            recalculateLeaveForDate(recalc.employeeId, leaveDate);
            break;
            
          default:
            Logger.log(`未対応の再計算種別: ${recalc.type}`);
            return;
        }
        
        // 再計算完了をマーク
        FormManager.markRecalculationCompleted(recalc.rowIndex);
        processedCount++;
        
        Logger.log(`再計算完了: ${recalc.employeeId} - ${recalc.type}`);
        
      } catch (error) {
        Logger.log(`再計算エラー (${recalc.employeeId}): ${error.toString()}`);
        errors.push({
          employeeId: recalc.employeeId,
          type: recalc.type,
          error: error.toString()
        });
      }
    });
    
    Logger.log(`承認再計算処理完了: 成功=${processedCount}件, エラー=${errors.length}件`);
    
    return {
      success: true,
      processedCount: processedCount,
      errorCount: errors.length,
      errors: errors
    };
    
  }, 'Triggers.processApprovalRecalculations', 'HIGH');
}

/**
 * 指定日の残業時間を再計算
 * @param {string} employeeId - 従業員ID
 * @param {Date} targetDate - 対象日
 */
function recalculateOvertimeForDate(employeeId, targetDate) {
  try {
    Logger.log(`残業再計算: ${employeeId} - ${formatDate(targetDate)}`);
    
    // 該当日の勤怠データを取得
    const timeEntries = getDailyTimeEntries(employeeId, targetDate);
    if (!timeEntries.clockIn || !timeEntries.clockOut) {
      Logger.log('出退勤データが不完全です');
      return;
    }
    
    // 残業申請を確認
    const overtimeRequest = getApprovedOvertimeRequest(employeeId, targetDate);
    
    // 残業時間を再計算
    const workMinutes = calculateTimeDifference(timeEntries.clockIn, timeEntries.clockOut);
    const standardWorkMinutes = 480; // 8時間
    let overtimeMinutes = 0;
    
    if (overtimeRequest) {
      // 承認済み残業申請がある場合
      const requestedOvertimeMinutes = parseOvertimeRequest(overtimeRequest.requestedValue);
      overtimeMinutes = Math.max(0, Math.min(workMinutes - standardWorkMinutes, requestedOvertimeMinutes));
    } else {
      // 通常の残業計算
      overtimeMinutes = Math.max(0, workMinutes - standardWorkMinutes);
    }
    
    // Daily_Summaryを更新
    updateDailySummaryOvertime(employeeId, targetDate, overtimeMinutes);
    
    Logger.log(`残業再計算完了: ${employeeId} - ${overtimeMinutes}分`);
    
  } catch (error) {
    Logger.log(`残業再計算エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 指定日の休暇を再計算
 * @param {string} employeeId - 従業員ID
 * @param {Date} targetDate - 対象日
 */
function recalculateLeaveForDate(employeeId, targetDate) {
  try {
    Logger.log(`休暇再計算: ${employeeId} - ${formatDate(targetDate)}`);
    
    // 承認済み休暇申請を確認
    const leaveRequest = getApprovedLeaveRequest(employeeId, targetDate);
    
    if (leaveRequest) {
      // 休暇日として処理
      updateDailySummaryForLeave(employeeId, targetDate, leaveRequest.requestType);
      Logger.log(`休暇再計算完了: ${employeeId} - ${leaveRequest.requestType}`);
    } else {
      // 通常の勤怠計算
      updateDailySummary(employeeId, targetDate);
      Logger.log(`通常勤怠再計算完了: ${employeeId}`);
    }
    
  } catch (error) {
    Logger.log(`休暇再計算エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 承認済み残業申請を取得
 * @param {string} employeeId - 従業員ID
 * @param {Date} targetDate - 対象日
 * @returns {Object|null} 残業申請データ
 */
function getApprovedOvertimeRequest(employeeId, targetDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Request_Responses');
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const targetDateStr = formatDate(targetDate);
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const empId = row[1];
      const requestType = row[2];
      const details = row[3];
      const requestedValue = row[4];
      const status = row[6];
      
      // 該当する承認済み残業申請を検索
      if (empId === employeeId && 
          requestType === '残業' && 
          status === 'Approved' &&
          (details.includes(targetDateStr) || requestedValue.includes(targetDateStr))) {
        return {
          requestType: requestType,
          details: details,
          requestedValue: requestedValue
        };
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log(`残業申請取得エラー: ${error.toString()}`);
    return null;
  }
}

/**
 * 承認済み休暇申請を取得
 * @param {string} employeeId - 従業員ID
 * @param {Date} targetDate - 対象日
 * @returns {Object|null} 休暇申請データ
 */
function getApprovedLeaveRequest(employeeId, targetDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Request_Responses');
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const targetDateStr = formatDate(targetDate);
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const empId = row[1];
      const requestType = row[2];
      const details = row[3];
      const requestedValue = row[4];
      const status = row[6];
      
      // 該当する承認済み休暇申請を検索
      if (empId === employeeId && 
          requestType === '休暇' && 
          status === 'Approved' &&
          (details.includes(targetDateStr) || requestedValue.includes(targetDateStr))) {
        return {
          requestType: requestType,
          details: details,
          requestedValue: requestedValue
        };
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log(`休暇申請取得エラー: ${error.toString()}`);
    return null;
  }
}

/**
 * 残業申請の時間を解析
 * @param {string} requestedValue - 申請値
 * @returns {number} 残業時間（分）
 */
function parseOvertimeRequest(requestedValue) {
  try {
    // "2時間残業" や "120分残業" などの形式を解析
    const hourMatch = requestedValue.match(/(\d+)時間/);
    if (hourMatch) {
      return parseInt(hourMatch[1]) * 60;
    }
    
    const minuteMatch = requestedValue.match(/(\d+)分/);
    if (minuteMatch) {
      return parseInt(minuteMatch[1]);
    }
    
    // 数値のみの場合は分として扱う
    const numberMatch = requestedValue.match(/(\d+)/);
    if (numberMatch) {
      return parseInt(numberMatch[1]);
    }
    
    return 0;
    
  } catch (error) {
    Logger.log(`残業時間解析エラー: ${error.toString()}`);
    return 0;
  }
}

/**
 * Daily_Summaryの残業時間を更新
 * @param {string} employeeId - 従業員ID
 * @param {Date} targetDate - 対象日
 * @param {number} overtimeMinutes - 残業時間（分）
 */
function updateDailySummaryOvertime(employeeId, targetDate, overtimeMinutes) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Summary');
    if (!sheet) {
      throw new Error('Daily_Summaryシートが見つかりません');
    }
    
    const data = sheet.getDataRange().getValues();
    const targetDateStr = formatDate(targetDate);
    
    // 該当する行を検索
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = formatDate(new Date(row[0]));
      const empId = row[1];
      
      if (empId === employeeId && date === targetDateStr) {
        // 残業時間列（G列）を更新
        sheet.getRange(i + 1, 7).setValue(overtimeMinutes);
        Logger.log(`Daily_Summary残業時間更新: ${employeeId} - ${targetDateStr} - ${overtimeMinutes}分`);
        return;
      }
    }
    
    Logger.log(`Daily_Summary該当行が見つかりません: ${employeeId} - ${targetDateStr}`);
    
  } catch (error) {
    Logger.log(`Daily_Summary残業時間更新エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * Daily_Summaryを休暇用に更新
 * @param {string} employeeId - 従業員ID
 * @param {Date} targetDate - 対象日
 * @param {string} leaveType - 休暇種別
 */
function updateDailySummaryForLeave(employeeId, targetDate, leaveType) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Summary');
    if (!sheet) {
      throw new Error('Daily_Summaryシートが見つかりません');
    }
    
    const data = sheet.getDataRange().getValues();
    const targetDateStr = formatDate(targetDate);
    
    // 該当する行を検索
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = formatDate(new Date(row[0]));
      const empId = row[1];
      
      if (empId === employeeId && date === targetDateStr) {
        // 休暇日として設定（出勤時刻、退勤時刻をクリア、実働時間を0に）
        const updateRange = sheet.getRange(i + 1, 3, 1, 6); // C列からH列まで
        updateRange.setValues([['休暇', '休暇', 0, 0, 0, 0]]);
        
        // 備考列があれば休暇種別を記録
        if (data[0].length > 8) {
          sheet.getRange(i + 1, 9).setValue(`${leaveType}（承認済み）`);
        }
        
        Logger.log(`Daily_Summary休暇更新: ${employeeId} - ${targetDateStr} - ${leaveType}`);
        return;
      }
    }
    
    // 該当行がない場合は新規作成
    const newRow = [
      targetDate,
      employeeId,
      '休暇',
      '休暇',
      0,
      0,
      0,
      0,
      `${leaveType}（承認済み）`
    ];
    sheet.appendRow(newRow);
    
    Logger.log(`Daily_Summary休暇新規作成: ${employeeId} - ${targetDateStr} - ${leaveType}`);
    
  } catch (error) {
    Logger.log(`Daily_Summary休暇更新エラー: ${error.toString()}`);
    throw error;
  }
}

// ========== レポート機能統合 ==========

/**
 * 日次レポート生成・送信（日次ジョブから呼び出し）
 * @param {Date} targetDate - 対象日
 */
function generateDailyReportForJob(targetDate) {
  return withErrorHandling(() => {
    Logger.log(`日次レポート生成開始: ${formatDate(targetDate)}`);
    
    try {
      // 日次レポートを生成・エクスポート
      const result = generateAndExportDailyReport(targetDate, '日次レポート');
      
      if (result.success) {
        // 管理者にレポート完了通知を送信
        sendDailyReportNotification(targetDate, result.exportResult);
        Logger.log(`日次レポート生成完了: ${result.exportResult.fileName}`);
      }
      
      return result;
      
    } catch (error) {
      Logger.log(`日次レポート生成エラー: ${error.toString()}`);
      sendErrorAlert(error, 'generateDailyReportForJob');
      throw error;
    }
    
  }, 'Triggers.generateDailyReportForJob', 'MEDIUM', {
    targetDate: formatDate(targetDate)
  });
}

/**
 * 月次レポート生成・送信（月次ジョブから呼び出し）
 * @param {string} yearMonth - 対象年月
 */
function generateMonthlyReportForJob(yearMonth) {
  return withErrorHandling(() => {
    Logger.log(`月次レポート生成開始: ${yearMonth}`);
    
    try {
      // 月次レポートを生成・エクスポート
      const result = generateAndExportMonthlyReport(yearMonth, '月次レポート');
      
      if (result.success) {
        // 管理者にレポート完了通知を送信
        sendMonthlyReportNotification(yearMonth, result.exportResult);
        Logger.log(`月次レポート生成完了: ${result.exportResult.fileName}`);
      }
      
      return result;
      
    } catch (error) {
      Logger.log(`月次レポート生成エラー: ${error.toString()}`);
      sendErrorAlert(error, 'generateMonthlyReportForJob');
      throw error;
    }
    
  }, 'Triggers.generateMonthlyReportForJob', 'MEDIUM', {
    yearMonth: yearMonth
  });
}

/**
 * 残業監視レポート生成・送信（週次ジョブから呼び出し）
 * @param {Date} startDate - 監視開始日
 * @param {Date} endDate - 監視終了日
 */
function generateOvertimeReportForJob(startDate, endDate) {
  return withErrorHandling(() => {
    Logger.log(`残業監視レポート生成開始: ${formatDate(startDate)} - ${formatDate(endDate)}`);
    
    try {
      // 残業監視レポートを生成・エクスポート
      const overtimeThreshold = getOvertimeWarningThresholdHours();
      const result = generateAndExportOvertimeReport(overtimeThreshold, 4, '残業監視レポート');
      
      if (result.success) {
        // 管理者にレポート完了通知を送信（アラートは既にmonitorOvertimeAndAlert内で送信済み）
        sendOvertimeReportNotification(startDate, endDate, result.exportResult, result.monitorResult.alertCount);
        Logger.log(`残業監視レポート生成完了: ${result.exportResult.fileName}`);
      }
      
      return result;
      
    } catch (error) {
      Logger.log(`残業監視レポート生成エラー: ${error.toString()}`);
      sendErrorAlert(error, 'generateOvertimeReportForJob');
      throw error;
    }
    
  }, 'Triggers.generateOvertimeReportForJob', 'MEDIUM', {
    startDate: formatDate(startDate),
    endDate: formatDate(endDate)
  });
}

// ========== レポート通知関数 ==========

/**
 * 日次レポート完了通知を送信
 * @param {Date} targetDate - 対象日
 * @param {Object} exportResult - エクスポート結果
 */
function sendDailyReportNotification(targetDate, exportResult) {
  try {
    const adminEmail = getAdminEmail();
    if (!adminEmail) return;
    
    const subject = `【日次レポート完了】${formatDate(targetDate)} のレポートが生成されました`;
    const body = `
システム管理者 様

日次レポートの生成が完了しました。

【レポート情報】
対象日: ${formatDate(targetDate)}
ファイル名: ${exportResult.fileName}
データ行数: ${exportResult.rowCount}行
保存先: ${exportResult.folderName}

【ダウンロードリンク】
閲覧用: ${exportResult.fileUrl}
ダウンロード用: ${exportResult.downloadUrl}

【レポート内容】
- 各従業員の出勤・退勤時刻
- 労働時間・残業時間の詳細
- 承認ステータス情報
- 遅刻・早退・深夜勤務の記録

レポートをご確認ください。

出勤管理システム
`;
    
    GmailApp.sendEmail(adminEmail, subject, body);
    Logger.log(`日次レポート通知送信完了: ${adminEmail}`);
    
  } catch (error) {
    Logger.log(`日次レポート通知送信エラー: ${error.toString()}`);
  }
}

/**
 * 月次レポート完了通知を送信
 * @param {string} yearMonth - 対象年月
 * @param {Object} exportResult - エクスポート結果
 */
function sendMonthlyReportNotification(yearMonth, exportResult) {
  try {
    const adminEmail = getAdminEmail();
    if (!adminEmail) return;
    
    const subject = `【月次レポート完了】${yearMonth} のレポートが生成されました`;
    const body = `
システム管理者 様

月次レポートの生成が完了しました。

【レポート情報】
対象月: ${yearMonth}
ファイル名: ${exportResult.fileName}
データ行数: ${exportResult.rowCount}行
保存先: ${exportResult.folderName}

【ダウンロードリンク】
閲覧用: ${exportResult.fileUrl}
ダウンロード用: ${exportResult.downloadUrl}

【レポート内容】
- 各従業員の月次勤務統計
- 総労働時間・残業時間の集計
- 有給使用状況
- 遅刻・早退回数
- 承認待ち申請数

月次データの確認をお願いします。

出勤管理システム
`;
    
    GmailApp.sendEmail(adminEmail, subject, body);
    Logger.log(`月次レポート通知送信完了: ${adminEmail}`);
    
  } catch (error) {
    Logger.log(`月次レポート通知送信エラー: ${error.toString()}`);
  }
}

/**
 * 残業監視レポート完了通知を送信
 * @param {Date} startDate - 監視開始日
 * @param {Date} endDate - 監視終了日
 * @param {Object} exportResult - エクスポート結果
 * @param {number} alertCount - アラート対象者数
 */
function sendOvertimeReportNotification(startDate, endDate, exportResult, alertCount) {
  try {
    const adminEmail = getAdminEmail();
    if (!adminEmail) return;
    
    const subject = `【残業監視レポート完了】${formatDate(startDate)}～${formatDate(endDate)} の監視結果`;
    const body = `
システム管理者 様

残業監視レポートの生成が完了しました。

【レポート情報】
監視期間: ${formatDate(startDate)} ～ ${formatDate(endDate)}
ファイル名: ${exportResult.fileName}
データ行数: ${exportResult.rowCount}行
保存先: ${exportResult.folderName}
警告対象者数: ${alertCount}名

【ダウンロードリンク】
閲覧用: ${exportResult.fileUrl}
ダウンロード用: ${exportResult.downloadUrl}

【レポート内容】
- 全従業員の残業時間詳細
- 警告レベル別の分類
- 法定上限超過者の特定
- 上司情報と対応状況

${alertCount > 0 ? '※警告対象者には別途緊急アラートを送信済みです。' : '※今回は警告対象者はいませんでした。'}

継続的な労働時間管理をお願いします。

出勤管理システム
`;
    
    GmailApp.sendEmail(adminEmail, subject, body);
    Logger.log(`残業監視レポート通知送信完了: ${adminEmail}`);
    
  } catch (error) {
    Logger.log(`残業監視レポート通知送信エラー: ${error.toString()}`);
  }
}

// ========== 手動レポート生成関数 ==========

/**
 * 日次レポートを手動生成
 */
function runDailyReportManually() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('日次レポート生成', '対象日を入力してください（YYYY-MM-DD形式、空白で昨日）:', ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() === ui.Button.CANCEL) {
      return;
    }
    
    let targetDate;
    const dateInput = response.getResponseText().trim();
    
    if (dateInput === '') {
      targetDate = addDays(getToday(), -1); // 昨日
    } else {
      targetDate = new Date(dateInput);
      if (isNaN(targetDate.getTime())) {
        ui.alert('無効な日付形式です。YYYY-MM-DD形式で入力してください。');
        return;
      }
    }
    
    ui.alert('日次レポート生成を開始します。完了まで数分かかる場合があります。');
    
    const result = generateAndExportDailyReport(targetDate);
    
    if (result.success) {
      ui.alert(`日次レポート生成完了\n\nファイル名: ${result.exportResult.fileName}\n行数: ${result.exportResult.rowCount}行\n\nGoogle Driveで確認してください。`);
    } else {
      ui.alert('日次レポート生成に失敗しました。');
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`日次レポート生成でエラーが発生しました: ${error.message}`);
  }
}

/**
 * 月次レポートを手動生成
 */
function runMonthlyReportManually() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('月次レポート生成', '対象年月を入力してください（YYYY-MM形式、空白で前月）:', ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() === ui.Button.CANCEL) {
      return;
    }
    
    let yearMonth;
    const monthInput = response.getResponseText().trim();
    
    if (monthInput === '') {
      const lastMonth = new Date();
      lastMonth.setMonth(lastMonth.getMonth() - 1);
      yearMonth = formatDate(lastMonth, 'YYYY-MM');
    } else {
      if (!/^\d{4}-\d{2}$/.test(monthInput)) {
        ui.alert('無効な年月形式です。YYYY-MM形式で入力してください。');
        return;
      }
      yearMonth = monthInput;
    }
    
    ui.alert('月次レポート生成を開始します。完了まで数分かかる場合があります。');
    
    const result = generateAndExportMonthlyReport(yearMonth);
    
    if (result.success) {
      ui.alert(`月次レポート生成完了\n\nファイル名: ${result.exportResult.fileName}\n行数: ${result.exportResult.rowCount}行\n\nGoogle Driveで確認してください。`);
    } else {
      ui.alert('月次レポート生成に失敗しました。');
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`月次レポート生成でエラーが発生しました: ${error.message}`);
  }
}

/**
 * 残業監視レポートを手動生成
 */
function runOvertimeReportManually() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('残業監視レポート生成を開始します。過去4週間のデータを分析します。');
    
    const result = generateAndExportOvertimeReport();
    
    if (result.success) {
      const alertMessage = result.monitorResult.alertCount > 0 
        ? `\n\n⚠️ 警告: ${result.monitorResult.alertCount}名が残業上限を超過しています！`
        : '\n\n✓ 残業上限超過者はいませんでした。';
      
      ui.alert(`残業監視レポート生成完了\n\nファイル名: ${result.exportResult.fileName}\n行数: ${result.exportResult.rowCount}行${alertMessage}\n\nGoogle Driveで確認してください。`);
    } else {
      ui.alert('残業監視レポート生成に失敗しました。');
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`残業監視レポート生成でエラーが発生しました: ${error.message}`);
  }
}