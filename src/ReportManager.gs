/**
 * レポート生成・エクスポート機能モジュール
 * 
 * 機能:
 * - 日次サマリーレポート生成機能
 * - 月次サマリーレポート生成機能
 * - CSVエクスポート機能（Google Driveへの保存）
 * - 残業時間監視・アラート機能
 * - レポートへの承認ステータス表示機能
 * 
 * 要件: 8.1, 8.2, 8.3, 8.4, 8.5
 */

/**
 * 日次サマリーレポート生成機能
 * 各従業員の出勤、労働時間、残業を示すレポートを作成
 * 
 * @param {Date} targetDate - レポート対象日（省略時は昨日）
 * @return {Object} レポート生成結果
 * 
 * 要件: 8.1 - 日次サマリーを生成する時、システムは各従業員の出勤、労働時間、残業を示すレポートを作成する
 */
function generateDailyReport(targetDate) {
  return withErrorHandling(() => {
    if (!targetDate) {
      targetDate = addDays(getToday(), -1); // デフォルトは昨日
    }
    
    Logger.log(`日次レポート生成開始: ${formatDate(targetDate)}`);
    
    // 全従業員のリストを取得
    const employees = getAllEmployees();
    const reportData = [];
    
    // ヘッダー行を追加
    const headers = [
      '従業員ID',
      '氏名',
      '所属',
      '出勤時刻',
      '退勤時刻',
      '休憩時間(分)',
      '実働時間(分)',
      '実働時間(時:分)',
      '残業時間(分)',
      '残業時間(時:分)',
      '遅刻早退(分)',
      '深夜勤務(分)',
      '承認ステータス',
      '備考'
    ];
    reportData.push(headers);
    
    // 各従業員のデータを処理
    employees.forEach(employee => {
      try {
        // 日次労働時間を計算
        const workTimeResult = calculateDailyWorkTime(employee.employeeId, targetDate);
        
        // 打刻データを取得
        const timeEntries = getDailyTimeEntries(employee.employeeId, targetDate);
        
        // 承認ステータスを取得
        const approvalStatus = getDailyApprovalStatus(employee.employeeId, targetDate);
        
        // 深夜勤務時間を計算
        const nightWorkMinutes = calculateNightWorkTime(timeEntries.clockIn, timeEntries.clockOut);
        
        // レポート行を作成
        const reportRow = [
          employee.employeeId,
          employee.name,
          employee.department || '',
          timeEntries.clockIn || '',
          timeEntries.clockOut || '',
          workTimeResult.breakMinutes || 0,
          workTimeResult.workMinutes || 0,
          minutesToTime(workTimeResult.workMinutes || 0),
          workTimeResult.overtimeMinutes || 0,
          minutesToTime(workTimeResult.overtimeMinutes || 0),
          (workTimeResult.lateMinutes || 0) + (workTimeResult.earlyLeaveMinutes || 0),
          nightWorkMinutes,
          approvalStatus.status,
          workTimeResult.message || approvalStatus.notes || ''
        ];
        
        reportData.push(reportRow);
        
      } catch (error) {
        Logger.log(`従業員データ処理エラー (${employee.employeeId}): ${error.toString()}`);
        
        // エラー行を追加
        const errorRow = [
          employee.employeeId,
          employee.name,
          employee.department || '',
          '', '', 0, 0, '00:00', 0, '00:00', 0, 0,
          'ERROR',
          `データ処理エラー: ${error.message}`
        ];
        reportData.push(errorRow);
      }
    });
    
    Logger.log(`日次レポート生成完了: ${reportData.length - 1}名分のデータ`);
    
    return {
      success: true,
      targetDate: formatDate(targetDate),
      reportData: reportData,
      employeeCount: reportData.length - 1,
      headers: headers
    };
    
  }, 'ReportManager.generateDailyReport', 'HIGH', {
    targetDate: formatDate(targetDate)
  });
}

/**
 * 月次サマリーレポート生成機能
 * 勤務日数、総労働時間、残業、休暇使用を集計
 * 
 * @param {string} yearMonth - 対象年月（YYYY-MM形式、省略時は前月）
 * @return {Object} レポート生成結果
 * 
 * 要件: 8.2 - 月次サマリーを作成する時、システムは勤務日数、総労働時間、残業、休暇使用を集計する
 */
function generateMonthlyReport(yearMonth) {
  return withErrorHandling(() => {
    if (!yearMonth) {
      const lastMonth = new Date();
      lastMonth.setMonth(lastMonth.getMonth() - 1);
      yearMonth = formatDate(lastMonth, 'YYYY-MM');
    }
    
    Logger.log(`月次レポート生成開始: ${yearMonth}`);
    
    // 対象月の期間を計算
    const [year, month] = yearMonth.split('-').map(Number);
    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 0);
    
    Logger.log(`集計期間: ${formatDate(startDate)} ～ ${formatDate(endDate)}`);
    
    // 全従業員のリストを取得
    const employees = getAllEmployees();
    const reportData = [];
    
    // ヘッダー行を追加
    const headers = [
      '従業員ID',
      '氏名',
      '所属',
      '勤務日数',
      '総労働時間(分)',
      '総労働時間(時:分)',
      '残業時間(分)',
      '残業時間(時:分)',
      '深夜勤務時間(分)',
      '有給使用日数',
      '遅刻回数',
      '早退回数',
      '承認待ち申請数',
      '備考'
    ];
    reportData.push(headers);
    
    // 各従業員の月次データを処理
    employees.forEach(employee => {
      try {
        // 月次サマリーを計算
        const monthlySummary = calculateEmployeeMonthlySummary(employee.employeeId, startDate, endDate);
        
        // 追加統計を計算
        const additionalStats = calculateAdditionalMonthlyStats(employee.employeeId, startDate, endDate);
        
        // レポート行を作成
        const reportRow = [
          employee.employeeId,
          employee.name,
          employee.department || '',
          monthlySummary.workDays,
          monthlySummary.totalWorkMinutes,
          minutesToTime(monthlySummary.totalWorkMinutes),
          monthlySummary.totalOvertimeMinutes,
          minutesToTime(monthlySummary.totalOvertimeMinutes),
          additionalStats.nightWorkMinutes,
          monthlySummary.vacationDays,
          additionalStats.lateCount,
          additionalStats.earlyLeaveCount,
          additionalStats.pendingApprovalCount,
          monthlySummary.notes || ''
        ];
        
        reportData.push(reportRow);
        
      } catch (error) {
        Logger.log(`従業員月次データ処理エラー (${employee.employeeId}): ${error.toString()}`);
        
        // エラー行を追加
        const errorRow = [
          employee.employeeId,
          employee.name,
          employee.department || '',
          0, 0, '00:00', 0, '00:00', 0, 0, 0, 0, 0,
          `データ処理エラー: ${error.message}`
        ];
        reportData.push(errorRow);
      }
    });
    
    Logger.log(`月次レポート生成完了: ${reportData.length - 1}名分のデータ`);
    
    return {
      success: true,
      yearMonth: yearMonth,
      startDate: formatDate(startDate),
      endDate: formatDate(endDate),
      reportData: reportData,
      employeeCount: reportData.length - 1,
      headers: headers
    };
    
  }, 'ReportManager.generateMonthlyReport', 'HIGH', {
    yearMonth: yearMonth
  });
}

/**
 * CSVエクスポート機能（Google Driveへの保存）
 * CSVファイルを生成しGoogle Driveにアクセスリンク付きで保存
 * 
 * @param {Array} reportData - レポートデータ（2次元配列）
 * @param {string} fileName - ファイル名（拡張子なし）
 * @param {string} folderName - 保存先フォルダ名（省略時はルート）
 * @return {Object} エクスポート結果
 * 
 * 要件: 8.3 - データをエクスポートする時、システムはCSVファイルを生成しGoogle Driveにアクセスリンク付きで保存する
 */
function exportReportToCsv(reportData, fileName, folderName) {
  return withErrorHandling(() => {
    if (!reportData || reportData.length === 0) {
      throw new Error('エクスポートするデータがありません');
    }
    
    if (!fileName) {
      throw new Error('ファイル名が指定されていません');
    }
    
    Logger.log(`CSVエクスポート開始: ${fileName}`);
    
    // CSVデータを作成
    const csvContent = reportData.map(row => 
      row.map(cell => {
        // セルの値を文字列に変換し、カンマやダブルクォートをエスケープ
        const cellValue = String(cell || '');
        if (cellValue.includes(',') || cellValue.includes('"') || cellValue.includes('\n')) {
          return '"' + cellValue.replace(/"/g, '""') + '"';
        }
        return cellValue;
      }).join(',')
    ).join('\n');
    
    // BOMを追加してExcelでの文字化けを防止
    const bom = '\uFEFF';
    const csvWithBom = bom + csvContent;
    
    // ファイル名にタイムスタンプを追加
    const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
    const fullFileName = `${fileName}_${timestamp}.csv`;
    
    // Google Driveにファイルを作成
    const blob = Utilities.newBlob(csvWithBom, 'text/csv', fullFileName);
    let file = DriveApp.createFile(blob);
    
    // 指定されたフォルダに移動
    if (folderName) {
      try {
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) {
          const folder = folders.next();
          folder.addFile(file);
          DriveApp.getRootFolder().removeFile(file);
          Logger.log(`ファイルをフォルダに移動: ${folderName}`);
        } else {
          // フォルダが存在しない場合は作成
          const newFolder = DriveApp.createFolder(folderName);
          newFolder.addFile(file);
          DriveApp.getRootFolder().removeFile(file);
          Logger.log(`新しいフォルダを作成してファイルを移動: ${folderName}`);
        }
      } catch (error) {
        Logger.log(`フォルダ操作エラー: ${error.toString()}`);
        // エラーが発生してもルートフォルダにファイルは作成されている
      }
    }
    
    // ファイルを共有可能に設定
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (error) {
      Logger.log(`ファイル共有設定エラー: ${error.toString()}`);
    }
    
    const fileUrl = `https://drive.google.com/file/d/${file.getId()}/view`;
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${file.getId()}`;
    
    Logger.log(`CSVエクスポート完了: ${fullFileName} (ID: ${file.getId()})`);
    
    return {
      success: true,
      fileName: fullFileName,
      fileId: file.getId(),
      fileUrl: fileUrl,
      downloadUrl: downloadUrl,
      rowCount: reportData.length,
      folderName: folderName || 'ルートフォルダ'
    };
    
  }, 'ReportManager.exportReportToCsv', 'HIGH', {
    fileName: fileName,
    folderName: folderName
  });
}

/**
 * 残業時間監視・アラート機能
 * 残業閾値を超えた時、管理者に週次アラートメールを送信
 * 
 * @param {number} thresholdHours - 警告閾値（時間、デフォルト80時間）
 * @param {number} periodWeeks - 監視期間（週数、デフォルト4週間）
 * @return {Object} 監視結果
 * 
 * 要件: 8.4 - 残業閾値を超えた時、システムは管理者に週次アラートメールを送信する
 */
function monitorOvertimeAndAlert(thresholdHours = 80, periodWeeks = 4) {
  return withErrorHandling(() => {
    Logger.log(`残業監視開始: 閾値=${thresholdHours}時間, 期間=${periodWeeks}週間`);
    
    // 監視期間を計算
    const endDate = addDays(getToday(), -1); // 昨日まで
    const startDate = addDays(endDate, -(periodWeeks * 7)); // 指定週数前から
    
    Logger.log(`監視期間: ${formatDate(startDate)} ～ ${formatDate(endDate)}`);
    
    // 全従業員のリストを取得
    const employees = getAllEmployees();
    const overtimeAlerts = [];
    const overtimeReport = [];
    
    // レポートヘッダー
    const reportHeaders = [
      '従業員ID',
      '氏名',
      '所属',
      '残業時間(分)',
      '残業時間(時:分)',
      '警告レベル',
      '上司メール',
      '備考'
    ];
    overtimeReport.push(reportHeaders);
    
    // 各従業員の残業時間を監視
    employees.forEach(employee => {
      try {
        // 期間内の残業時間を計算
        const overtimeMinutes = calculateEmployeeOvertimeForPeriod(employee.employeeId, startDate, endDate);
        const overtimeHours = Math.round(overtimeMinutes / 60 * 10) / 10; // 小数点1桁
        
        // 警告レベルを判定
        let warningLevel = 'NORMAL';
        let notes = '';
        
        if (overtimeHours >= thresholdHours) {
          warningLevel = 'CRITICAL';
          notes = `法定上限(${thresholdHours}時間)超過`;
          
          // アラート対象に追加
          overtimeAlerts.push({
            employeeId: employee.employeeId,
            name: employee.name,
            email: employee.email,
            department: employee.department,
            supervisor: employee.supervisorEmail,
            overtimeHours: overtimeHours,
            overtimeMinutes: overtimeMinutes,
            warningLevel: warningLevel
          });
          
        } else if (overtimeHours >= thresholdHours * 0.8) {
          warningLevel = 'WARNING';
          notes = `警告レベル(${Math.round(thresholdHours * 0.8)}時間)到達`;
        } else if (overtimeHours >= thresholdHours * 0.6) {
          warningLevel = 'CAUTION';
          notes = `注意レベル(${Math.round(thresholdHours * 0.6)}時間)到達`;
        }
        
        // レポートに追加
        const reportRow = [
          employee.employeeId,
          employee.name,
          employee.department || '',
          overtimeMinutes,
          minutesToTime(overtimeMinutes),
          warningLevel,
          employee.supervisorEmail || '',
          notes
        ];
        overtimeReport.push(reportRow);
        
      } catch (error) {
        Logger.log(`従業員残業監視エラー (${employee.employeeId}): ${error.toString()}`);
        
        // エラー行を追加
        const errorRow = [
          employee.employeeId,
          employee.name,
          employee.department || '',
          0, '00:00', 'ERROR',
          employee.supervisorEmail || '',
          `監視エラー: ${error.message}`
        ];
        overtimeReport.push(errorRow);
      }
    });
    
    // アラートメールを送信
    if (overtimeAlerts.length > 0) {
      sendOvertimeAlertReport(overtimeAlerts, startDate, endDate, thresholdHours);
    }
    
    Logger.log(`残業監視完了: 警告対象=${overtimeAlerts.length}名, 総監視対象=${employees.length}名`);
    
    return {
      success: true,
      monitoringPeriod: {
        startDate: formatDate(startDate),
        endDate: formatDate(endDate),
        weeks: periodWeeks
      },
      thresholdHours: thresholdHours,
      totalEmployees: employees.length,
      alertCount: overtimeAlerts.length,
      alerts: overtimeAlerts,
      reportData: overtimeReport
    };
    
  }, 'ReportManager.monitorOvertimeAndAlert', 'HIGH', {
    thresholdHours: thresholdHours,
    periodWeeks: periodWeeks
  });
}

/**
 * レポートへの承認ステータス表示機能
 * すべての時間調整の承認ステータスを含める
 * 
 * @param {string} employeeId - 従業員ID
 * @param {Date} targetDate - 対象日
 * @return {Object} 承認ステータス情報
 * 
 * 要件: 8.5 - レポートを生成する時、システムはすべての時間調整の承認ステータスを含める
 */
function getDailyApprovalStatus(employeeId, targetDate) {
  return withErrorHandling(() => {
    const targetDateStr = formatDate(targetDate);
    const requestData = getSheetData('Request_Responses', 'A:G');
    
    const approvalInfo = {
      status: 'NORMAL',
      pendingCount: 0,
      approvedCount: 0,
      rejectedCount: 0,
      notes: '',
      requests: []
    };
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < requestData.length; i++) {
      const row = requestData[i];
      const requestDate = new Date(row[0]);
      const empId = row[1];
      const requestType = row[2];
      const details = row[3];
      const requestedValue = row[4];
      const approver = row[5];
      const status = row[6];
      
      // 該当従業員の該当日の申請を検索
      if (empId === employeeId && formatDate(requestDate) === targetDateStr) {
        const requestInfo = {
          type: requestType,
          details: details,
          requestedValue: requestedValue,
          approver: approver,
          status: status
        };
        
        approvalInfo.requests.push(requestInfo);
        
        // ステータス別カウント
        switch (status) {
          case 'Pending':
            approvalInfo.pendingCount++;
            break;
          case 'Approved':
            approvalInfo.approvedCount++;
            break;
          case 'Rejected':
            approvalInfo.rejectedCount++;
            break;
        }
      }
    }
    
    // 全体ステータスを決定
    if (approvalInfo.pendingCount > 0) {
      approvalInfo.status = 'PENDING';
      approvalInfo.notes = `承認待ち${approvalInfo.pendingCount}件`;
    } else if (approvalInfo.approvedCount > 0) {
      approvalInfo.status = 'APPROVED';
      approvalInfo.notes = `承認済み${approvalInfo.approvedCount}件`;
    } else if (approvalInfo.rejectedCount > 0) {
      approvalInfo.status = 'REJECTED';
      approvalInfo.notes = `却下${approvalInfo.rejectedCount}件`;
    }
    
    // 複数ステータスが混在する場合
    if (approvalInfo.approvedCount > 0 && approvalInfo.rejectedCount > 0) {
      approvalInfo.status = 'MIXED';
      approvalInfo.notes = `承認${approvalInfo.approvedCount}件/却下${approvalInfo.rejectedCount}件`;
    }
    
    return approvalInfo;
    
  }, 'ReportManager.getDailyApprovalStatus', 'LOW', {
    employeeId: employeeId,
    targetDate: formatDate(targetDate),
    suppressError: true,
    defaultValue: {
      status: 'NORMAL',
      pendingCount: 0,
      approvedCount: 0,
      rejectedCount: 0,
      notes: '',
      requests: []
    }
  });
}

// ========== 補助関数 ==========

/**
 * 追加の月次統計を計算
 * @param {string} employeeId - 従業員ID
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} 追加統計情報
 */
function calculateAdditionalMonthlyStats(employeeId, startDate, endDate) {
  try {
    const summaryData = getSheetData('Daily_Summary', 'A:I');
    const requestData = getSheetData('Request_Responses', 'A:G');
    
    // 従業員情報を一度だけ取得（ループ外で最適化）
    const employeeInfo = getEmployeeInfoById(employeeId);
    
    let nightWorkMinutes = 0;
    let lateCount = 0;
    let earlyLeaveCount = 0;
    let pendingApprovalCount = 0;
    
    // Daily_Summaryから統計を計算
    for (let i = 1; i < summaryData.length; i++) {
      const row = summaryData[i];
      const date = new Date(row[0]);
      const empId = row[1];
      const clockIn = row[2];
      const clockOut = row[3];
      const lateEarlyMinutes = row[7] || 0;
      
      if (empId === employeeId && date >= startDate && date <= endDate) {
        // 深夜勤務時間を計算
        if (clockIn && clockOut) {
          nightWorkMinutes += calculateNightWorkTime(clockIn, clockOut);
        }
        
        // 遅刻・早退回数をカウント
        if (Number(lateEarlyMinutes) > 0) {
          // 詳細な判定のため打刻データを確認
          const timeEntries = getDailyTimeEntries(employeeId, date);
          
          if (employeeInfo && timeEntries.clockIn) {
            const standardStartMinutes = timeToMinutes(employeeInfo.standardStartTime);
            const actualStartMinutes = timeToMinutes(timeEntries.clockIn);
            
            if (actualStartMinutes > standardStartMinutes) {
              lateCount++;
            }
          }
          
          if (employeeInfo && timeEntries.clockOut) {
            const standardEndMinutes = timeToMinutes(employeeInfo.standardEndTime);
            const actualEndMinutes = timeToMinutes(timeEntries.clockOut);
            
            if (actualEndMinutes < standardEndMinutes) {
              earlyLeaveCount++;
            }
          }
        }
      }
    }
    
    // Request_Responsesから承認待ち件数を計算
    for (let i = 1; i < requestData.length; i++) {
      const row = requestData[i];
      const requestDate = new Date(row[0]);
      const empId = row[1];
      const status = row[6];
      
      if (empId === employeeId && 
          requestDate >= startDate && 
          requestDate <= endDate &&
          status === 'Pending') {
        pendingApprovalCount++;
      }
    }
    
    return {
      nightWorkMinutes: nightWorkMinutes,
      lateCount: lateCount,
      earlyLeaveCount: earlyLeaveCount,
      pendingApprovalCount: pendingApprovalCount
    };
    
  } catch (error) {
    Logger.log(`追加統計計算エラー (${employeeId}): ${error.toString()}`);
    return {
      nightWorkMinutes: 0,
      lateCount: 0,
      earlyLeaveCount: 0,
      pendingApprovalCount: 0
    };
  }
}

/**
 * 残業アラートレポートを送信
 * @param {Array} overtimeAlerts - 残業アラート配列
 * @param {Date} startDate - 監視開始日
 * @param {Date} endDate - 監視終了日
 * @param {number} thresholdHours - 閾値時間
 */
function sendOvertimeAlertReport(overtimeAlerts, startDate, endDate, thresholdHours) {
  try {
    const adminEmail = getAdminEmail();
    if (!adminEmail) {
      Logger.log('管理者メールアドレスが設定されていません');
      return;
    }
    
    const subject = `【緊急】残業時間警告レポート - ${overtimeAlerts.length}名が閾値超過`;
    
    let body = `
システム管理者 様

残業時間監視の結果、法定上限を超過した従業員が検出されました。
至急対応をお願いします。

【監視結果】
監視期間: ${formatDate(startDate)} ～ ${formatDate(endDate)}
警告閾値: ${thresholdHours}時間
超過者数: ${overtimeAlerts.length}名

【超過者一覧】
`;

    overtimeAlerts.forEach((alert, index) => {
      body += `
${index + 1}. ${alert.name} (${alert.employeeId})
   - 残業時間: ${alert.overtimeHours}時間
   - 所属: ${alert.department || '不明'}
   - 上司: ${alert.supervisor || '不明'}
   - 警告レベル: ${alert.warningLevel}
`;
    });

    body += `

【対応事項】
1. 各従業員の業務量調整を検討してください
2. 必要に応じて人員配置の見直しを行ってください
3. 労働基準法遵守のため、適切な措置を講じてください
4. 従業員の健康管理にもご注意ください

【参考情報】
- 労働基準法では時間外労働の上限が月45時間、年360時間
- 特別な事情がある場合でも月80時間、年720時間が上限
- 継続的な長時間労働は健康リスクを高めます

このメールは自動送信されています。

出勤管理システム
`;

    GmailApp.sendEmail(adminEmail, subject, body);
    Logger.log(`残業アラートレポート送信完了: ${adminEmail}`);
    
  } catch (error) {
    Logger.log(`残業アラートレポート送信エラー: ${error.toString()}`);
    throw error;
  }
}

// ========== 統合レポート機能 ==========

/**
 * 日次レポートを生成してCSVエクスポート
 * @param {Date} targetDate - 対象日
 * @param {string} folderName - 保存先フォルダ名
 * @return {Object} 処理結果
 */
function generateAndExportDailyReport(targetDate, folderName = '日次レポート') {
  return withErrorHandling(() => {
    // 日次レポートを生成
    const reportResult = generateDailyReport(targetDate);
    
    if (!reportResult.success) {
      throw new Error('日次レポート生成に失敗しました');
    }
    
    // CSVエクスポート
    const fileName = `日次レポート_${reportResult.targetDate}`;
    const exportResult = exportReportToCsv(reportResult.reportData, fileName, folderName);
    
    if (!exportResult.success) {
      throw new Error('CSVエクスポートに失敗しました');
    }
    
    Logger.log(`日次レポート生成・エクスポート完了: ${exportResult.fileName}`);
    
    return {
      success: true,
      reportResult: reportResult,
      exportResult: exportResult
    };
    
  }, 'ReportManager.generateAndExportDailyReport', 'HIGH', {
    targetDate: formatDate(targetDate),
    folderName: folderName
  });
}

/**
 * 月次レポートを生成してCSVエクスポート
 * @param {string} yearMonth - 対象年月
 * @param {string} folderName - 保存先フォルダ名
 * @return {Object} 処理結果
 */
function generateAndExportMonthlyReport(yearMonth, folderName = '月次レポート') {
  return withErrorHandling(() => {
    // 月次レポートを生成
    const reportResult = generateMonthlyReport(yearMonth);
    
    if (!reportResult.success) {
      throw new Error('月次レポート生成に失敗しました');
    }
    
    // CSVエクスポート
    const fileName = `月次レポート_${reportResult.yearMonth}`;
    const exportResult = exportReportToCsv(reportResult.reportData, fileName, folderName);
    
    if (!exportResult.success) {
      throw new Error('CSVエクスポートに失敗しました');
    }
    
    Logger.log(`月次レポート生成・エクスポート完了: ${exportResult.fileName}`);
    
    return {
      success: true,
      reportResult: reportResult,
      exportResult: exportResult
    };
    
  }, 'ReportManager.generateAndExportMonthlyReport', 'HIGH', {
    yearMonth: yearMonth,
    folderName: folderName
  });
}

/**
 * 残業監視レポートを生成してCSVエクスポート
 * @param {number} thresholdHours - 警告閾値
 * @param {number} periodWeeks - 監視期間
 * @param {string} folderName - 保存先フォルダ名
 * @return {Object} 処理結果
 */
function generateAndExportOvertimeReport(thresholdHours = 80, periodWeeks = 4, folderName = '残業監視レポート') {
  return withErrorHandling(() => {
    // 残業監視を実行
    const monitorResult = monitorOvertimeAndAlert(thresholdHours, periodWeeks);
    
    if (!monitorResult.success) {
      throw new Error('残業監視に失敗しました');
    }
    
    // CSVエクスポート
    const fileName = `残業監視レポート_${monitorResult.monitoringPeriod.startDate}_${monitorResult.monitoringPeriod.endDate}`;
    const exportResult = exportReportToCsv(monitorResult.reportData, fileName, folderName);
    
    if (!exportResult.success) {
      throw new Error('CSVエクスポートに失敗しました');
    }
    
    Logger.log(`残業監視レポート生成・エクスポート完了: ${exportResult.fileName}`);
    
    return {
      success: true,
      monitorResult: monitorResult,
      exportResult: exportResult
    };
    
  }, 'ReportManager.generateAndExportOvertimeReport', 'HIGH', {
    thresholdHours: thresholdHours,
    periodWeeks: periodWeeks,
    folderName: folderName
  });
}