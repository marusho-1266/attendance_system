/**
 * フォーム管理・承認ワークフローモジュール
 * Googleフォームからの申請を受信し、承認ワークフローを管理する
 */

/**
 * Googleフォーム送信時のトリガー関数
 * フォーム申請を受信し、Request_Responsesシートに記録する
 * @param {Object} e - フォーム送信イベントオブジェクト
 */
function onRequestSubmit(e) {
  return withErrorHandling(() => {
    Logger.log('フォーム申請受信開始');
    
    // フォームデータの取得
    const formData = extractFormData(e);
    Logger.log('フォームデータ: ' + JSON.stringify(formData));
    
    // 申請データの検証
    validateRequestData(formData);
    
    // 承認者の特定
    const approver = identifyApprover(formData.employeeId);
    
    // Request_Responsesシートに記録
    recordRequestToSheet(formData, approver);
    
    Logger.log('フォーム申請処理完了: ' + formData.employeeId);
    
    return {
      success: true,
      employeeId: formData.employeeId,
      requestType: formData.requestType,
      approver: approver
    };
    
  }, 'FormManager.onRequestSubmit', 'HIGH', {
    formData: e ? JSON.stringify(e.values) : 'No form data'
  });
}

/**
 * フォームイベントからデータを抽出する
 * @param {Object} e - フォーム送信イベント
 * @returns {Object} 抽出されたフォームデータ
 */
function extractFormData(e) {
  if (!e || !e.values) {
    throw new Error('フォームデータが無効です');
  }
  
  const values = e.values;
  
  // フォームの構造に基づいてデータを抽出
  // 想定フォーム構造: [タイムスタンプ, 社員ID, 申請種別, 詳細, 希望値]
  const formData = {
    timestamp: values[0] || new Date(),
    employeeId: values[1] || '',
    requestType: values[2] || '',
    details: values[3] || '',
    requestedValue: values[4] || ''
  };
  
  return formData;
}

/**
 * 申請データの検証を行う
 * @param {Object} formData - 検証対象のフォームデータ
 * @throws {Error} 検証エラー時
 */
function validateRequestData(formData) {
  const errors = [];
  
  // 必須フィールドの検証
  if (!formData.employeeId || formData.employeeId.trim() === '') {
    errors.push('社員IDが入力されていません');
  }
  
  if (!formData.requestType || formData.requestType.trim() === '') {
    errors.push('申請種別が選択されていません');
  }
  
  if (!formData.details || formData.details.trim() === '') {
    errors.push('詳細が入力されていません');
  }
  
  // 申請種別の妥当性検証
  const validRequestTypes = ['修正', '残業', '休暇'];
  if (formData.requestType && !validRequestTypes.includes(formData.requestType)) {
    errors.push('無効な申請種別です: ' + formData.requestType);
  }
  
  // 社員IDの存在確認
  if (formData.employeeId) {
    const employee = getEmployeeById(formData.employeeId);
    if (!employee) {
      errors.push('存在しない社員IDです: ' + formData.employeeId);
    }
  }
  
  if (errors.length > 0) {
    throw new Error('申請データ検証エラー: ' + errors.join(', '));
  }
}

/**
 * 承認者を特定する
 * Master_Employeeシートから上司情報を取得
 * @param {string} employeeId - 申請者の社員ID
 * @returns {string} 承認者のメールアドレス
 */
function identifyApprover(employeeId) {
  try {
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      throw new Error('従業員情報が見つかりません: ' + employeeId);
    }
    
    const approverEmail = employee.supervisorEmail;
    if (!approverEmail || approverEmail.trim() === '') {
      throw new Error('承認者が設定されていません: ' + employeeId);
    }
    
    Logger.log('承認者特定完了: ' + approverEmail);
    return approverEmail;
    
  } catch (error) {
    Logger.log('承認者特定エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 社員IDから従業員情報を取得する
 * @param {string} employeeId - 社員ID
 * @returns {Object|null} 従業員情報オブジェクト
 */
function getEmployeeById(employeeId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Employee');
    if (!sheet) {
      throw new Error('Master_Employeeシートが見つかりません');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // ヘッダー行のインデックスを取得
    const idIndex = headers.indexOf('社員ID');
    const nameIndex = headers.indexOf('氏名');
    const emailIndex = headers.indexOf('Gmail');
    const supervisorIndex = headers.indexOf('上司Gmail');
    
    if (idIndex === -1) {
      throw new Error('社員ID列が見つかりません');
    }
    
    // 該当する従業員を検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === employeeId) {
        return {
          id: data[i][idIndex],
          name: nameIndex !== -1 ? data[i][nameIndex] : '',
          email: emailIndex !== -1 ? data[i][emailIndex] : '',
          supervisorEmail: supervisorIndex !== -1 ? data[i][supervisorIndex] : ''
        };
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log('従業員情報取得エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 申請データをRequest_Responsesシートに記録する
 * @param {Object} formData - フォームデータ
 * @param {string} approver - 承認者メールアドレス
 */
function recordRequestToSheet(formData, approver) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Request_Responses');
    if (!sheet) {
      throw new Error('Request_Responsesシートが見つかりません');
    }
    
    // 新しい行のデータを準備
    const newRow = [
      formData.timestamp,
      formData.employeeId,
      formData.requestType,
      formData.details,
      formData.requestedValue,
      approver,
      'Pending'  // 初期ステータス
    ];
    
    // シートに行を追加
    sheet.appendRow(newRow);
    
    Logger.log('申請データ記録完了: ' + formData.employeeId + ' - ' + formData.requestType);
    
  } catch (error) {
    Logger.log('申請データ記録エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 承認処理用の関数
 * 承認者がステータスを変更した際の処理
 * @param {string} requestId - 申請ID（行番号）
 * @param {string} newStatus - 新しいステータス
 */
function processApprovalStatusChange(requestId, newStatus) {
  return withErrorHandling(() => {
    Logger.log('承認ステータス変更処理開始: ' + requestId + ' -> ' + newStatus);
    
    if (!['Approved', 'Rejected'].includes(newStatus)) {
      throw new Error('無効なステータスです: ' + newStatus);
    }
    
    // 申請データの取得
    const requestData = getRequestData(requestId);
    if (!requestData) {
      throw new Error('申請データが見つかりません: ' + requestId);
    }
    
    // 現在のステータスと同じ場合は処理をスキップ
    if (requestData.status === newStatus) {
      Logger.log('ステータス変更なし: ' + requestId);
      return { success: true, message: 'No status change' };
    }
    
    // 承認履歴の記録
    recordApprovalHistory(requestId, requestData, newStatus);
    
    // ステータス変更の記録
    updateRequestStatus(requestId, newStatus);
    
    // 承認された場合は再計算をスケジュール
    if (newStatus === 'Approved') {
      scheduleRecalculation(requestData.employeeId, requestData.requestType, requestData);
    }
    
    // 申請者への通知送信
    sendApplicantNotification(requestData, newStatus);
    
    Logger.log('承認ステータス変更完了: ' + requestId);
    
    return {
      success: true,
      requestId: requestId,
      newStatus: newStatus,
      employeeId: requestData.employeeId
    };
    
  }, 'FormManager.processApprovalStatusChange', 'HIGH', {
    requestId: requestId,
    newStatus: newStatus
  });
}

/**
 * 申請データを取得する
 * @param {string} requestId - 申請ID（行番号）
 * @returns {Object|null} 申請データ
 */
function getRequestData(requestId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Request_Responses');
    if (!sheet) {
      return null;
    }
    
    const rowIndex = parseInt(requestId);
    if (isNaN(rowIndex) || rowIndex < 2) {
      return null;
    }
    
    const data = sheet.getRange(rowIndex, 1, 1, 7).getValues()[0];
    
    return {
      timestamp: data[0],
      employeeId: data[1],
      requestType: data[2],
      details: data[3],
      requestedValue: data[4],
      approver: data[5],
      status: data[6]
    };
    
  } catch (error) {
    Logger.log('申請データ取得エラー: ' + error.toString());
    return null;
  }
}

/**
 * 申請ステータスを更新する
 * @param {string} requestId - 申請ID（行番号）
 * @param {string} status - 新しいステータス
 */
function updateRequestStatus(requestId, status) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Request_Responses');
    if (!sheet) {
      throw new Error('Request_Responsesシートが見つかりません');
    }
    
    const rowIndex = parseInt(requestId);
    if (isNaN(rowIndex) || rowIndex < 2) {
      throw new Error('無効な申請IDです: ' + requestId);
    }
    
    // ステータス列（G列）を更新
    sheet.getRange(rowIndex, 7).setValue(status);
    
    Logger.log('ステータス更新完了: ' + requestId + ' -> ' + status);
    
  } catch (error) {
    Logger.log('ステータス更新エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 承認後の再計算をスケジュールする
 * @param {string} employeeId - 従業員ID
 * @param {string} requestType - 申請種別
 * @param {Object} requestData - 申請データ
 */
function scheduleRecalculation(employeeId, requestType, requestData) {
  try {
    Logger.log('再計算スケジュール: ' + employeeId + ' - ' + requestType);
    
    // 申請種別に応じた再計算処理
    switch (requestType) {
      case '修正':
        // 時刻修正の場合は該当日の再計算をマーク
        markForRecalculation(employeeId, extractDateFromRequest(requestData));
        break;
        
      case '残業':
        // 残業申請の場合は該当日の残業時間再計算をマーク
        markForOvertimeRecalculation(employeeId, extractDateFromRequest(requestData));
        break;
        
      case '休暇':
        // 休暇申請の場合は該当期間の勤怠再計算をマーク
        markForLeaveRecalculation(employeeId, extractDateRangeFromRequest(requestData));
        break;
        
      default:
        Logger.log('未対応の申請種別: ' + requestType);
    }
    
    Logger.log('再計算スケジュール完了');
    
  } catch (error) {
    Logger.log('再計算スケジュールエラー: ' + error.toString());
    throw error;
  }
}

/**
 * 申請データから対象日付を抽出する
 * @param {Object} requestData - 申請データ
 * @returns {Date} 対象日付
 */
function extractDateFromRequest(requestData) {
  try {
    // 詳細フィールドから日付を抽出（例: "2024/01/15の出勤時刻を9:00に修正"）
    const dateMatch = requestData.details.match(/(\d{4}\/\d{1,2}\/\d{1,2})/);
    if (dateMatch) {
      return new Date(dateMatch[1]);
    }
    
    // 希望値フィールドから日付を抽出
    const valueMatch = requestData.requestedValue.match(/(\d{4}\/\d{1,2}\/\d{1,2})/);
    if (valueMatch) {
      return new Date(valueMatch[1]);
    }
    
    // デフォルトは申請日
    return new Date(requestData.timestamp);
    
  } catch (error) {
    Logger.log('日付抽出エラー: ' + error.toString());
    return new Date(requestData.timestamp);
  }
}

/**
 * 申請データから対象期間を抽出する
 * @param {Object} requestData - 申請データ
 * @returns {Object} 開始日と終了日
 */
function extractDateRangeFromRequest(requestData) {
  try {
    // 期間指定の場合（例: "2024/01/15-2024/01/17"）
    const rangeMatch = requestData.details.match(/(\d{4}\/\d{1,2}\/\d{1,2})-(\d{4}\/\d{1,2}\/\d{1,2})/);
    if (rangeMatch) {
      return {
        startDate: new Date(rangeMatch[1]),
        endDate: new Date(rangeMatch[2])
      };
    }
    
    // 単日の場合
    const singleDate = extractDateFromRequest(requestData);
    return {
      startDate: singleDate,
      endDate: singleDate
    };
    
  } catch (error) {
    Logger.log('期間抽出エラー: ' + error.toString());
    const defaultDate = new Date(requestData.timestamp);
    return {
      startDate: defaultDate,
      endDate: defaultDate
    };
  }
}

/**
 * 日次再計算のマーキング
 * @param {string} employeeId - 従業員ID
 * @param {Date} targetDate - 対象日付
 */
function markForRecalculation(employeeId, targetDate) {
  try {
    const sheet = getOrCreateRecalculationSheet();
    const dateStr = Utilities.formatDate(targetDate, 'JST', 'yyyy/MM/dd');
    
    // 既存のマーキングをチェック
    const existingMark = findRecalculationMark(sheet, employeeId, dateStr, 'DAILY');
    if (existingMark) {
      Logger.log('既に再計算マーク済み: ' + employeeId + ' - ' + dateStr);
      return;
    }
    
    // 新しいマーキングを追加
    sheet.appendRow([
      new Date(),
      employeeId,
      dateStr,
      'DAILY',
      'PENDING',
      '時刻修正による再計算'
    ]);
    
    Logger.log('日次再計算マーク完了: ' + employeeId + ' - ' + dateStr);
    
  } catch (error) {
    Logger.log('日次再計算マークエラー: ' + error.toString());
    throw error;
  }
}

/**
 * 残業時間再計算のマーキング
 * @param {string} employeeId - 従業員ID
 * @param {Date} targetDate - 対象日付
 */
function markForOvertimeRecalculation(employeeId, targetDate) {
  try {
    const sheet = getOrCreateRecalculationSheet();
    const dateStr = Utilities.formatDate(targetDate, 'JST', 'yyyy/MM/dd');
    
    // 既存のマーキングをチェック
    const existingMark = findRecalculationMark(sheet, employeeId, dateStr, 'OVERTIME');
    if (existingMark) {
      Logger.log('既に残業再計算マーク済み: ' + employeeId + ' - ' + dateStr);
      return;
    }
    
    // 新しいマーキングを追加
    sheet.appendRow([
      new Date(),
      employeeId,
      dateStr,
      'OVERTIME',
      'PENDING',
      '残業申請による再計算'
    ]);
    
    Logger.log('残業再計算マーク完了: ' + employeeId + ' - ' + dateStr);
    
  } catch (error) {
    Logger.log('残業再計算マークエラー: ' + error.toString());
    throw error;
  }
}

/**
 * 休暇再計算のマーキング
 * @param {string} employeeId - 従業員ID
 * @param {Object} dateRange - 対象期間
 */
function markForLeaveRecalculation(employeeId, dateRange) {
  try {
    const sheet = getOrCreateRecalculationSheet();
    
    // 期間内の各日をマーキング
    const currentDate = new Date(dateRange.startDate);
    while (currentDate <= dateRange.endDate) {
      const dateStr = Utilities.formatDate(currentDate, 'JST', 'yyyy/MM/dd');
      
      // 既存のマーキングをチェック
      const existingMark = findRecalculationMark(sheet, employeeId, dateStr, 'LEAVE');
      if (!existingMark) {
        // 新しいマーキングを追加
        sheet.appendRow([
          new Date(),
          employeeId,
          dateStr,
          'LEAVE',
          'PENDING',
          '休暇申請による再計算'
        ]);
      }
      
      // 次の日へ
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    Logger.log('休暇再計算マーク完了: ' + employeeId + ' - ' + 
               Utilities.formatDate(dateRange.startDate, 'JST', 'yyyy/MM/dd') + ' to ' +
               Utilities.formatDate(dateRange.endDate, 'JST', 'yyyy/MM/dd'));
    
  } catch (error) {
    Logger.log('休暇再計算マークエラー: ' + error.toString());
    throw error;
  }
}

/**
 * 再計算管理シートを取得または作成
 * @returns {Sheet} 再計算管理シート
 */
function getOrCreateRecalculationSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('Recalculation_Queue');
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = spreadsheet.insertSheet('Recalculation_Queue');
      
      // ヘッダー行を設定
      const headers = [
        'マーク日時',
        '社員ID',
        '対象日付',
        '種別',
        'ステータス',
        '備考'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダー行のフォーマット
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      
      Logger.log('再計算管理シート作成完了');
    }
    
    return sheet;
    
  } catch (error) {
    Logger.log('再計算管理シート取得エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 既存の再計算マークを検索
 * @param {Sheet} sheet - 再計算管理シート
 * @param {string} employeeId - 従業員ID
 * @param {string} dateStr - 対象日付文字列
 * @param {string} type - 再計算種別
 * @returns {boolean} 既存マークの有無
 */
function findRecalculationMark(sheet, employeeId, dateStr, type) {
  try {
    const data = sheet.getDataRange().getValues();
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === employeeId && 
          data[i][2] === dateStr && 
          data[i][3] === type &&
          data[i][4] === 'PENDING') {
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    Logger.log('再計算マーク検索エラー: ' + error.toString());
    return false;
  }
}

/**
 * バッチ承認通知処理
 * 同じ承認者への複数申請を1通のメールにまとめる
 */
function batchApprovalNotifications() {
  try {
    Logger.log('バッチ承認通知処理開始');
    
    const pendingRequests = getPendingRequests();
    if (pendingRequests.length === 0) {
      Logger.log('承認待ち申請がありません');
      return;
    }
    
    // 承認者別にグループ化
    const groupedRequests = groupRequestsByApprover(pendingRequests);
    
    // 各承認者に通知メールを送信
    Object.keys(groupedRequests).forEach(approver => {
      const requests = groupedRequests[approver];
      sendApprovalNotification(approver, requests);
    });
    
    Logger.log('バッチ承認通知処理完了');
    
  } catch (error) {
    Logger.log('バッチ承認通知処理エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 承認待ち申請を取得する
 * @returns {Array} 承認待ち申請の配列
 */
function getPendingRequests() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Request_Responses');
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const pendingRequests = [];
    
    // ヘッダー行をスキップして処理
    for (let i = 1; i < data.length; i++) {
      if (data[i][6] === 'Pending') {  // ステータス列
        pendingRequests.push({
          rowIndex: i + 1,
          timestamp: data[i][0],
          employeeId: data[i][1],
          requestType: data[i][2],
          details: data[i][3],
          requestedValue: data[i][4],
          approver: data[i][5],
          status: data[i][6]
        });
      }
    }
    
    return pendingRequests;
    
  } catch (error) {
    Logger.log('承認待ち申請取得エラー: ' + error.toString());
    return [];
  }
}

/**
 * 申請を承認者別にグループ化する
 * @param {Array} requests - 申請の配列
 * @returns {Object} 承認者別にグループ化された申請
 */
function groupRequestsByApprover(requests) {
  const grouped = {};
  
  requests.forEach(request => {
    if (!grouped[request.approver]) {
      grouped[request.approver] = [];
    }
    grouped[request.approver].push(request);
  });
  
  return grouped;
}

/**
 * 承認通知メールを送信する
 * @param {string} approver - 承認者メールアドレス
 * @param {Array} requests - 承認対象の申請配列
 */
function sendApprovalNotification(approver, requests) {
  try {
    if (!approver || requests.length === 0) {
      return;
    }
    
    const subject = `【出勤管理】承認依頼 (${requests.length}件)`;
    let body = `承認者様\n\n以下の申請について承認をお願いいたします。\n\n`;
    
    requests.forEach((request, index) => {
      body += `${index + 1}. ${request.requestType}申請\n`;
      body += `   申請者: ${request.employeeId}\n`;
      body += `   詳細: ${request.details}\n`;
      body += `   希望値: ${request.requestedValue}\n`;
      body += `   申請日時: ${Utilities.formatDate(request.timestamp, 'JST', 'yyyy/MM/dd HH:mm')}\n\n`;
    });
    
    body += `承認は以下のシートで行ってください:\n`;
    body += `${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=Request_Responses\n\n`;
    body += `※ステータス列で「Approved」または「Rejected」を選択してください。`;
    
    // メール送信（MailManagerが利用可能な場合）
    if (typeof MailManager !== 'undefined' && MailManager.sendBatchNotification) {
      MailManager.sendBatchNotification([approver], subject, body);
    } else {
      // 直接Gmail送信
      GmailApp.sendEmail(approver, subject, body);
    }
    
    Logger.log('承認通知送信完了: ' + approver + ' (' + requests.length + '件)');
    
  } catch (error) {
    Logger.log('承認通知送信エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 承認履歴を記録する
 * @param {string} requestId - 申請ID
 * @param {Object} requestData - 申請データ
 * @param {string} newStatus - 新しいステータス
 */
function recordApprovalHistory(requestId, requestData, newStatus) {
  try {
    Logger.log('承認履歴記録開始: ' + requestId);
    
    const sheet = getOrCreateApprovalHistorySheet();
    const approver = Session.getActiveUser().getEmail();
    
    // 承認履歴データを準備
    const historyRow = [
      new Date(),                    // 処理日時
      requestId,                     // 申請ID
      requestData.employeeId,        // 申請者ID
      requestData.requestType,       // 申請種別
      requestData.status,            // 変更前ステータス
      newStatus,                     // 変更後ステータス
      approver,                      // 承認者
      requestData.details,           // 申請詳細
      requestData.requestedValue,    // 希望値
      '承認処理'                     // 処理種別
    ];
    
    // 履歴シートに記録
    sheet.appendRow(historyRow);
    
    Logger.log('承認履歴記録完了: ' + requestId);
    
  } catch (error) {
    Logger.log('承認履歴記録エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 承認履歴管理シートを取得または作成
 * @returns {Sheet} 承認履歴シート
 */
function getOrCreateApprovalHistorySheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('Approval_History');
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = spreadsheet.insertSheet('Approval_History');
      
      // ヘッダー行を設定
      const headers = [
        '処理日時',
        '申請ID',
        '申請者ID',
        '申請種別',
        '変更前ステータス',
        '変更後ステータス',
        '承認者',
        '申請詳細',
        '希望値',
        '処理種別'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダー行のフォーマット
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#34a853');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      
      // 列幅の調整
      sheet.setColumnWidth(1, 150);  // 処理日時
      sheet.setColumnWidth(2, 80);   // 申請ID
      sheet.setColumnWidth(3, 100);  // 申請者ID
      sheet.setColumnWidth(4, 100);  // 申請種別
      sheet.setColumnWidth(7, 200);  // 承認者
      sheet.setColumnWidth(8, 300);  // 申請詳細
      
      Logger.log('承認履歴シート作成完了');
    }
    
    return sheet;
    
  } catch (error) {
    Logger.log('承認履歴シート取得エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 申請者への処理完了通知を送信
 * @param {Object} requestData - 申請データ
 * @param {string} status - 処理結果ステータス
 */
function sendApplicantNotification(requestData, status) {
  try {
    Logger.log('申請者通知送信開始: ' + requestData.employeeId + ' - ' + status);
    
    // 申請者の情報を取得
    const applicant = getEmployeeById(requestData.employeeId);
    if (!applicant || !applicant.email) {
      Logger.log('申請者のメールアドレスが見つかりません: ' + requestData.employeeId);
      return;
    }
    
    // 通知メールの内容を作成
    const subject = `【出勤管理】申請処理完了通知 - ${getStatusDisplayName(status)}`;
    const body = createApplicantNotificationBody(requestData, status);
    
    // メール送信
    if (typeof MailManager !== 'undefined' && MailManager.sendBatchNotification) {
      MailManager.sendBatchNotification([applicant.email], subject, body);
    } else {
      GmailApp.sendEmail(applicant.email, subject, body);
    }
    
    Logger.log('申請者通知送信完了: ' + applicant.email);
    
  } catch (error) {
    Logger.log('申請者通知送信エラー: ' + error.toString());
    // 通知送信エラーは処理を停止させない
  }
}

/**
 * ステータスの表示名を取得
 * @param {string} status - ステータス
 * @returns {string} 表示名
 */
function getStatusDisplayName(status) {
  const statusMap = {
    'Approved': '承認',
    'Rejected': '却下',
    'Pending': '承認待ち'
  };
  
  return statusMap[status] || status;
}

/**
 * 申請者向け通知メール本文を作成
 * @param {Object} requestData - 申請データ
 * @param {string} status - 処理結果ステータス
 * @returns {string} メール本文
 */
function createApplicantNotificationBody(requestData, status) {
  const statusDisplay = getStatusDisplayName(status);
  const isApproved = status === 'Approved';
  
  let body = `${requestData.employeeId} 様\n\n`;
  body += `以下の申請について処理が完了いたしました。\n\n`;
  body += `■ 処理結果\n`;
  body += `${statusDisplay}\n\n`;
  body += `■ 申請内容\n`;
  body += `申請種別: ${requestData.requestType}\n`;
  body += `申請日時: ${Utilities.formatDate(requestData.timestamp, 'JST', 'yyyy/MM/dd HH:mm')}\n`;
  body += `詳細: ${requestData.details}\n`;
  body += `希望値: ${requestData.requestedValue}\n\n`;
  
  if (isApproved) {
    body += `■ 処理について\n`;
    body += `承認されました。次回の日次処理で勤怠データに反映されます。\n\n`;
    
    // 申請種別に応じた追加情報
    switch (requestData.requestType) {
      case '修正':
        body += `時刻修正が承認されました。該当日の勤怠サマリーが自動更新されます。\n`;
        break;
      case '残業':
        body += `残業申請が承認されました。残業時間が勤怠記録に反映されます。\n`;
        break;
      case '休暇':
        body += `休暇申請が承認されました。該当期間の勤怠記録が更新されます。\n`;
        break;
    }
  } else {
    body += `■ 処理について\n`;
    body += `申請は却下されました。詳細については承認者にお問い合わせください。\n\n`;
  }
  
  body += `■ 確認\n`;
  body += `勤怠記録の確認は以下のシートで行えます:\n`;
  body += `${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=Daily_Summary\n\n`;
  body += `※このメールは自動送信されています。\n`;
  body += `※ご不明な点がございましたら、システム管理者までお問い合わせください。`;
  
  return body;
}

/**
 * ステータス変更を検出するトリガー関数
 * Request_Responsesシートの編集時に呼び出される
 * @param {Object} e - 編集イベント
 */
function onRequestResponsesEdit(e) {
  return withErrorHandling(() => {
    Logger.log('Request_Responsesシート編集検出');
    
    if (!e || !e.range) {
      Logger.log('無効な編集イベント');
      return;
    }
    
    const range = e.range;
    const sheet = range.getSheet();
    
    // Request_Responsesシート以外は処理しない
    if (sheet.getName() !== 'Request_Responses') {
      return;
    }
    
    // ステータス列（G列）の編集のみ処理
    if (range.getColumn() !== 7) {
      return;
    }
    
    const row = range.getRow();
    const newStatus = range.getValue();
    
    // ヘッダー行は処理しない
    if (row <= 1) {
      return;
    }
    
    // 有効なステータス値のみ処理
    if (!['Approved', 'Rejected'].includes(newStatus)) {
      return;
    }
    
    Logger.log('ステータス変更検出: 行' + row + ' -> ' + newStatus);
    
    // 承認処理を実行
    processApprovalStatusChange(row.toString(), newStatus);
    
  }, 'FormManager.onRequestResponsesEdit', 'MEDIUM', {
    range: e.range ? e.range.getA1Notation() : 'Unknown',
    value: e.range ? e.range.getValue() : 'Unknown'
  });
}

/**
 * 承認待ち申請の一括処理
 * 管理者が手動で実行する関数
 */
function processPendingApprovals() {
  return withErrorHandling(() => {
    Logger.log('承認待ち申請一括処理開始');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Request_Responses');
    if (!sheet) {
      throw new Error('Request_Responsesシートが見つかりません');
    }
    
    const data = sheet.getDataRange().getValues();
    let processedCount = 0;
    
    // ヘッダー行をスキップして処理
    for (let i = 1; i < data.length; i++) {
      const status = data[i][6]; // ステータス列
      
      // Pendingから変更されたものを検出
      if (['Approved', 'Rejected'].includes(status)) {
        const rowIndex = i + 1;
        
        // 既に処理済みかチェック（承認履歴を確認）
        if (!isAlreadyProcessed(rowIndex.toString(), status)) {
          Logger.log('未処理の承認を検出: 行' + rowIndex + ' - ' + status);
          processApprovalStatusChange(rowIndex.toString(), status);
          processedCount++;
        }
      }
    }
    
    Logger.log('承認待ち申請一括処理完了: ' + processedCount + '件処理');
    
    return {
      success: true,
      processedCount: processedCount
    };
    
  }, 'FormManager.processPendingApprovals', 'HIGH');
}

/**
 * 申請が既に処理済みかチェック
 * @param {string} requestId - 申請ID
 * @param {string} status - ステータス
 * @returns {boolean} 処理済みの場合true
 */
function isAlreadyProcessed(requestId, status) {
  try {
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approval_History');
    if (!historySheet) {
      return false;
    }
    
    const data = historySheet.getDataRange().getValues();
    
    // 履歴から該当する処理を検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === requestId && data[i][5] === status) {
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    Logger.log('処理済みチェックエラー: ' + error.toString());
    return false;
  }
}

/**
 * 再計算待ちキューから処理対象を取得
 * 日次ジョブから呼び出される
 * @returns {Array} 再計算対象の配列
 */
function getPendingRecalculations() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Recalculation_Queue');
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const pending = [];
    
    // ヘッダー行をスキップして処理
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] === 'PENDING') { // ステータス列
        pending.push({
          rowIndex: i + 1,
          markTime: data[i][0],
          employeeId: data[i][1],
          targetDate: data[i][2],
          type: data[i][3],
          status: data[i][4],
          note: data[i][5]
        });
      }
    }
    
    return pending;
    
  } catch (error) {
    Logger.log('再計算待ちキュー取得エラー: ' + error.toString());
    return [];
  }
}

/**
 * 再計算完了をマーク
 * @param {number} rowIndex - 行インデックス
 */
function markRecalculationCompleted(rowIndex) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Recalculation_Queue');
    if (!sheet) {
      return;
    }
    
    // ステータスを'COMPLETED'に更新
    sheet.getRange(rowIndex, 5).setValue('COMPLETED');
    
    // 完了日時を備考に追加
    const currentNote = sheet.getRange(rowIndex, 6).getValue();
    const completedNote = currentNote + ' (完了: ' + 
                         Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm') + ')';
    sheet.getRange(rowIndex, 6).setValue(completedNote);
    
    Logger.log('再計算完了マーク: 行' + rowIndex);
    
  } catch (error) {
    Logger.log('再計算完了マークエラー: ' + error.toString());
  }
}