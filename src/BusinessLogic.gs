/**
 * 業務ロジックモジュール
 * 出勤管理システムの核となる計算処理を提供
 */

/**
 * 日次労働時間計算機能
 * 指定した従業員の指定日の労働時間を計算
 * @param {string} employeeId - 従業員ID
 * @param {Date} date - 計算対象日
 * @return {Object} 労働時間計算結果
 */
function calculateDailyWorkTime(employeeId, date) {
  return withErrorHandling(() => {
    if (!employeeId) {
      throw new Error('従業員IDが指定されていません');
    }
    
    if (!date || !(date instanceof Date)) {
      throw new Error('有効な日付オブジェクトが必要です');
    }
    
    // 従業員情報を取得
    const employeeInfo = getEmployeeInfoById(employeeId);
    if (!employeeInfo) {
      throw new Error(`従業員ID ${employeeId} が見つかりません`);
    }
    
    // 指定日の打刻データを取得
    const timeEntries = getDailyTimeEntries(employeeId, date);
    
    // 打刻データが不完全な場合の処理
    if (!timeEntries.clockIn) {
      return {
        workMinutes: 0,
        overtimeMinutes: 0,
        breakMinutes: 0,
        lateMinutes: 0,
        earlyLeaveMinutes: 0,
        nightWorkMinutes: 0,
        status: 'INCOMPLETE',
        message: '出勤打刻がありません'
      };
    }
    
    if (!timeEntries.clockOut) {
      return {
        workMinutes: 0,
        overtimeMinutes: 0,
        breakMinutes: 0,
        lateMinutes: 0,
        earlyLeaveMinutes: 0,
        nightWorkMinutes: 0,
        status: 'INCOMPLETE',
        message: '退勤打刻がありません'
      };
    }
    
    // 基本労働時間を計算
    const totalMinutes = calculateTimeDifference(timeEntries.clockIn, timeEntries.clockOut);
    
    // 休憩時間を計算・控除
    const breakMinutes = calculateBreakTime(timeEntries);
    const workMinutes = Math.max(0, totalMinutes - breakMinutes);
    
    // 休憩時間自動控除を適用
    const adjustedWorkTime = applyBreakDeduction(workMinutes);
    
    // 残業時間を計算
    const standardWorkMinutes = getStandardWorkMinutes(employeeInfo);
    const overtimeMinutes = calculateOvertime(adjustedWorkTime.workMinutes, standardWorkMinutes, date);
    
    // 遅刻・早退を判定
    const lateEarlyResult = calculateLateAndEarlyLeave(employeeInfo, timeEntries);
    
    // 深夜勤務時間を計算
    const nightWorkMinutes = calculateNightWorkTime(timeEntries.clockIn, timeEntries.clockOut);
    
    return {
      workMinutes: adjustedWorkTime.workMinutes,
      overtimeMinutes: overtimeMinutes,
      breakMinutes: adjustedWorkTime.breakMinutes,
      lateMinutes: lateEarlyResult.lateMinutes,
      earlyLeaveMinutes: lateEarlyResult.earlyLeaveMinutes,
      nightWorkMinutes: nightWorkMinutes,
      status: 'COMPLETE',
      message: '計算完了'
    };
    
  }, 'BusinessLogic.calculateDailyWorkTime', 'HIGH', {
    employeeId: employeeId,
    date: formatDate(date)
  });
}

/**
 * 残業時間計算機能
 * 基準労働時間を超過した時間を残業として計算
 * @param {number} workMinutes - 実労働時間（分）
 * @param {number} standardMinutes - 基準労働時間（分）
 * @param {Date} date - 対象日（祝日判定用）
 * @return {number} 残業時間（分）
 */
function calculateOvertime(workMinutes, standardMinutes, date) {
  try {
    if (typeof workMinutes !== 'number' || workMinutes < 0) {
      throw new Error('有効な労働時間が必要です');
    }
    
    if (typeof standardMinutes !== 'number' || standardMinutes < 0) {
      throw new Error('有効な基準労働時間が必要です');
    }
    
    // 祝日判定
    const holidayInfo = isHoliday(date);
    
    // 祝日の場合は全時間が残業扱い
    if (holidayInfo.isHoliday) {
      return workMinutes;
    }
    
    // 土日の判定
    const dayOfWeek = date.getDay();
    const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6); // 0=日曜, 6=土曜
    
    // 土日の場合は全時間が残業扱い
    if (isWeekend) {
      return workMinutes;
    }
    
    // 平日の場合は基準時間を超過した分が残業
    return Math.max(0, workMinutes - standardMinutes);
    
  } catch (error) {
    Logger.log('残業時間計算エラー: ' + error.toString());
    return 0;
  }
}

/**
 * 休憩時間自動控除機能
 * 労働時間に応じて法定休憩時間を自動控除
 * @param {number} workMinutes - 労働時間（分）
 * @return {Object} {workMinutes: 調整後労働時間, breakMinutes: 控除された休憩時間}
 */
function applyBreakDeduction(workMinutes) {
  try {
    if (typeof workMinutes !== 'number' || workMinutes < 0) {
      return {
        workMinutes: 0,
        breakMinutes: 0
      };
    }
    
    let breakMinutes = 0;
    
    // 6時間を超える場合は45分の休憩を控除
    if (workMinutes > 6 * 60) {
      breakMinutes = 45;
    }
    
    // 8時間を超える場合は追加で15分の休憩を控除（合計60分）
    if (workMinutes > 8 * 60) {
      breakMinutes = 60;
    }
    
    const adjustedWorkMinutes = Math.max(0, workMinutes - breakMinutes);
    
    return {
      workMinutes: adjustedWorkMinutes,
      breakMinutes: breakMinutes
    };
    
  } catch (error) {
    Logger.log('休憩時間控除エラー: ' + error.toString());
    return {
      workMinutes: workMinutes,
      breakMinutes: 0
    };
  }
}

/**
 * 遅刻・早退判定機能
 * 従業員の基準時刻と実際の打刻時刻を比較して遅刻・早退を判定
 * @param {Object} employeeInfo - 従業員情報
 * @param {Object} timeEntries - 打刻データ
 * @return {Object} {lateMinutes: 遅刻時間, earlyLeaveMinutes: 早退時間}
 */
function calculateLateAndEarlyLeave(employeeInfo, timeEntries) {
  try {
    if (!employeeInfo || !timeEntries) {
      return {
        lateMinutes: 0,
        earlyLeaveMinutes: 0
      };
    }
    
    let lateMinutes = 0;
    let earlyLeaveMinutes = 0;
    
    // 遅刻判定
    if (employeeInfo.standardStartTime && timeEntries.clockIn) {
      const standardStartMinutes = timeToMinutes(employeeInfo.standardStartTime);
      const actualStartMinutes = timeToMinutes(timeEntries.clockIn);
      
      if (actualStartMinutes > standardStartMinutes) {
        lateMinutes = actualStartMinutes - standardStartMinutes;
      }
    }
    
    // 早退判定
    if (employeeInfo.standardEndTime && timeEntries.clockOut) {
      const standardEndMinutes = timeToMinutes(employeeInfo.standardEndTime);
      const actualEndMinutes = timeToMinutes(timeEntries.clockOut);
      
      if (actualEndMinutes < standardEndMinutes) {
        earlyLeaveMinutes = standardEndMinutes - actualEndMinutes;
      }
    }
    
    return {
      lateMinutes: lateMinutes,
      earlyLeaveMinutes: earlyLeaveMinutes
    };
    
  } catch (error) {
    Logger.log('遅刻・早退判定エラー: ' + error.toString());
    return {
      lateMinutes: 0,
      earlyLeaveMinutes: 0
    };
  }
}

// ========== 補助関数 ==========

/**
 * 従業員情報を取得（ID検索）
 * @param {string} employeeId - 従業員ID
 * @return {Object|null} 従業員情報
 */
function getEmployeeInfoById(employeeId) {
  try {
    const employeeData = getSheetData('Master_Employee', 'A:H');
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < employeeData.length; i++) {
      const row = employeeData[i];
      if (row[0] === employeeId) {
        return {
          employeeId: row[0],
          name: row[1],
          email: row[2],
          department: row[3],
          employmentType: row[4],
          supervisorEmail: row[5],
          standardStartTime: row[6],
          standardEndTime: row[7]
        };
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log(`従業員情報取得エラー (${employeeId}): ` + error.toString());
    return null;
  }
}

/**
 * 指定日の打刻データを取得
 * @param {string} employeeId - 従業員ID
 * @param {Date} date - 対象日
 * @return {Object} 打刻データ
 */
function getDailyTimeEntries(employeeId, date) {
  try {
    const targetDateStr = formatDate(date);
    const logData = getSheetData('Log_Raw', 'A:F');
    
    const entries = {
      clockIn: null,
      clockOut: null,
      breakIn: null,
      breakOut: null
    };
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < logData.length; i++) {
      const row = logData[i];
      const timestamp = new Date(row[0]);
      const empId = row[1];
      const action = row[3];
      
      // 同じ従業員の同じ日のデータのみ処理
      if (empId === employeeId && formatDate(timestamp) === targetDateStr) {
        const timeStr = formatTime(timestamp);
        
        switch (action) {
          case 'IN':
            if (!entries.clockIn) entries.clockIn = timeStr;
            break;
          case 'OUT':
            entries.clockOut = timeStr; // 最後の退勤を記録
            break;
          case 'BRK_IN':
            entries.breakIn = timeStr;
            break;
          case 'BRK_OUT':
            entries.breakOut = timeStr;
            break;
        }
      }
    }
    
    return entries;
    
  } catch (error) {
    Logger.log(`打刻データ取得エラー (${employeeId}, ${formatDate(date)}): ` + error.toString());
    return {
      clockIn: null,
      clockOut: null,
      breakIn: null,
      breakOut: null
    };
  }
}

/**
 * 休憩時間を計算
 * @param {Object} timeEntries - 打刻データ
 * @return {number} 休憩時間（分）
 */
function calculateBreakTime(timeEntries) {
  try {
    if (!timeEntries.breakIn || !timeEntries.breakOut) {
      return 0;
    }
    
    return calculateTimeDifference(timeEntries.breakIn, timeEntries.breakOut);
    
  } catch (error) {
    Logger.log('休憩時間計算エラー: ' + error.toString());
    return 0;
  }
}

/**
 * 基準労働時間を取得
 * @param {Object} employeeInfo - 従業員情報
 * @return {number} 基準労働時間（分）
 */
function getStandardWorkMinutes(employeeInfo) {
  try {
    if (!employeeInfo.standardStartTime || !employeeInfo.standardEndTime) {
      return 8 * 60; // デフォルト8時間
    }
    
    return calculateTimeDifference(employeeInfo.standardStartTime, employeeInfo.standardEndTime);
    
  } catch (error) {
    Logger.log('基準労働時間取得エラー: ' + error.toString());
    return 8 * 60; // デフォルト8時間
  }
}

/**
 * Daily_Summaryシートを更新
 * @param {string} employeeId - 従業員ID
 * @param {Date} date - 対象日
 */
function updateDailySummary(employeeId, date) {
  return withErrorHandling(() => {
    // 労働時間を計算
    const workTimeResult = calculateDailyWorkTime(employeeId, date);
    
    if (workTimeResult.status === 'INCOMPLETE') {
      Logger.log(`Daily_Summary更新スキップ (${employeeId}, ${formatDate(date)}): ${workTimeResult.message}`);
      return;
    }
    
    // Daily_Summaryシートのデータを取得
    const summaryData = getSheetData('Daily_Summary', 'A:I');
    const targetDateStr = formatDate(date);
    
    // 既存データを検索
    let targetRowIndex = -1;
    for (let i = 1; i < summaryData.length; i++) {
      const row = summaryData[i];
      if (formatDate(new Date(row[0])) === targetDateStr && row[1] === employeeId) {
        targetRowIndex = i;
        break;
      }
    }
    
    // 打刻データを取得
    const timeEntries = getDailyTimeEntries(employeeId, date);
    
    // 更新データを準備
    const updateRow = [
      date,
      employeeId,
      timeEntries.clockIn || '',
      timeEntries.clockOut || '',
      workTimeResult.breakMinutes,
      workTimeResult.workMinutes,
      workTimeResult.overtimeMinutes,
      workTimeResult.lateMinutes + workTimeResult.earlyLeaveMinutes,
      'CALCULATED'
    ];
    
    if (targetRowIndex >= 0) {
      // 既存行を更新
      summaryData[targetRowIndex] = updateRow;
    } else {
      // 新規行を追加
      summaryData.push(updateRow);
    }
    
    // シートに書き戻し
    setSheetData('Daily_Summary', 'A:I', summaryData);
    
    Logger.log(`Daily_Summary更新完了 (${employeeId}, ${targetDateStr})`);
    
  }, 'BusinessLogic.updateDailySummary', 'HIGH', {
    employeeId: employeeId,
    date: formatDate(date)
  });
}