/**
 * メール管理・通知システム
 * 
 * 機能:
 * - バッチメール送信（クォータ節約）
 * - 承認依頼メールのバッチ処理
 * - エラー通知メール
 * - 未退勤者通知メール
 * 
 * 要件: 2.3, 4.2, 6.2
 */

/**
 * バッチメール送信機能（クォータ節約）
 * 同一受信者への複数メールを1通にまとめて送信
 * 
 * @param {Array} notifications - 通知オブジェクトの配列
 * @param {string} notifications[].recipient - 受信者メールアドレス
 * @param {string} notifications[].subject - 件名
 * @param {string} notifications[].body - 本文
 */
function sendBatchNotifications(notifications) {
  return withErrorHandling(() => {
    if (!notifications || notifications.length === 0) {
      Logger.log('MailManager: 送信する通知がありません');
      return;
    }

    // 受信者ごとにグループ化
    const groupedByRecipient = {};
    notifications.forEach(notification => {
      const recipient = notification.recipient;
      if (!groupedByRecipient[recipient]) {
        groupedByRecipient[recipient] = [];
      }
      groupedByRecipient[recipient].push(notification);
    });

    // バッチ処理でメール送信（クォータ制限対策）
    const recipients = Object.keys(groupedByRecipient);
    const sendResult = processBatch(recipients, (recipient, index) => {
      const messages = groupedByRecipient[recipient];
      
      if (messages.length === 1) {
        // 単一メッセージの場合はそのまま送信
        const msg = messages[0];
        GmailApp.sendEmail(recipient, msg.subject, msg.body);
        Logger.log(`MailManager: 単一メール送信完了 - ${recipient}`);
      } else {
        // 複数メッセージの場合は結合して送信
        const combinedMessage = combineMessages(messages);
        GmailApp.sendEmail(recipient, combinedMessage.subject, combinedMessage.body);
        Logger.log(`MailManager: バッチメール送信完了 - ${recipient} (${messages.length}件)`);
      }
      
      return { recipient: recipient, messageCount: messages.length };
      
    }, {
      context: 'MailManager.BatchSend',
      batchSize: 5, // Gmail送信制限を考慮
      delay: 2000,  // 2秒間隔
      maxExecutionTime: ERROR_CONFIG.MAX_EXECUTION_TIME
    });

    Logger.log(`MailManager: バッチ送信完了 - 総通知数: ${notifications.length}, 送信メール数: ${sendResult.processedCount}`);
    
    return {
      success: true,
      totalNotifications: notifications.length,
      sentEmails: sendResult.processedCount,
      errors: sendResult.errors
    };

  }, 'MailManager.sendBatchNotifications', 'HIGH');
}

/**
 * 複数メッセージを1通のメールに結合
 * 
 * @param {Array} messages - メッセージ配列
 * @return {Object} 結合されたメッセージ
 */
function combineMessages(messages) {
  const subjects = messages.map(msg => msg.subject);
  const uniqueSubjects = [...new Set(subjects)];
  
  let combinedSubject;
  if (uniqueSubjects.length === 1) {
    combinedSubject = uniqueSubjects[0];
  } else {
    combinedSubject = '【まとめ通知】出勤管理システムからの通知';
  }

  let combinedBody = '出勤管理システムからの通知をまとめてお送りします。\n\n';
  combinedBody += '=' * 50 + '\n\n';

  messages.forEach((msg, index) => {
    combinedBody += `【通知 ${index + 1}】${msg.subject}\n`;
    combinedBody += '-' * 30 + '\n';
    combinedBody += msg.body + '\n\n';
  });

  combinedBody += '=' * 50 + '\n';
  combinedBody += '※このメールは自動送信されています。\n';
  combinedBody += `送信時刻: ${Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss')}`;

  return {
    subject: combinedSubject,
    body: combinedBody
  };
}

/**
 * 承認依頼メールのバッチ処理機能
 * 同日の同一承認者への申請をまとめて通知
 * 
 * @param {Array} approvalRequests - 承認依頼の配列
 */
function sendApprovalRequestBatch(approvalRequests) {
  try {
    if (!approvalRequests || approvalRequests.length === 0) {
      Logger.log('MailManager: 承認依頼がありません');
      return;
    }

    const notifications = approvalRequests.map(request => {
      const subject = `【承認依頼】${request.employeeName}さんからの${request.requestType}申請`;
      const body = createApprovalRequestBody(request);
      
      return {
        recipient: request.approverEmail,
        subject: subject,
        body: body
      };
    });

    sendBatchNotifications(notifications);
    Logger.log(`MailManager: 承認依頼バッチ送信完了 - ${approvalRequests.length}件`);

  } catch (error) {
    Logger.log(`MailManager: 承認依頼バッチ送信エラー - ${error.toString()}`);
    sendErrorAlert(error, 'sendApprovalRequestBatch');
    throw error;
  }
}

/**
 * 承認依頼メール本文を作成
 * 
 * @param {Object} request - 承認依頼オブジェクト
 * @return {string} メール本文
 */
function createApprovalRequestBody(request) {
  const body = `
${request.approverName} 様

お疲れ様です。
${request.employeeName}さんから${request.requestType}の申請が提出されました。

【申請詳細】
申請者: ${request.employeeName}
申請種別: ${request.requestType}
申請日時: ${Utilities.formatDate(request.requestDate, 'JST', 'yyyy/MM/dd HH:mm')}
対象日: ${request.targetDate}
詳細: ${request.details}
希望値: ${request.requestedValue}

【承認方法】
以下のリンクから承認シートにアクセスし、該当行のStatus列で承認・却下を選択してください。

承認シート: ${getApprovalSheetUrl()}

※このメールは自動送信されています。
※承認処理後、申請者に自動で通知されます。

出勤管理システム
`;

  return body;
}

/**
 * エラー通知メール機能
 * システムエラー発生時に管理者に即座に通知
 * 
 * @param {Error} error - エラーオブジェクト
 * @param {string} context - エラー発生コンテキスト
 */
function sendErrorAlert(error, context) {
  try {
    const adminEmail = getAdminEmail();
    if (!adminEmail) {
      Logger.log('MailManager: 管理者メールアドレスが設定されていません');
      return;
    }

    const subject = `【緊急】出勤管理システムエラー: ${context}`;
    const body = createErrorAlertBody(error, context);

    GmailApp.sendEmail(adminEmail, subject, body);
    Logger.log(`MailManager: エラー通知送信完了 - ${context}`);

  } catch (mailError) {
    // エラー通知の送信に失敗した場合はログのみ記録
    Logger.log(`MailManager: エラー通知送信失敗 - 元エラー: ${error.toString()}, 送信エラー: ${mailError.toString()}`);
  }
}

/**
 * エラー通知メール本文を作成
 * 
 * @param {Error} error - エラーオブジェクト
 * @param {string} context - エラー発生コンテキスト
 * @return {string} メール本文
 */
function createErrorAlertBody(error, context) {
  const body = `
システム管理者 様

出勤管理システムでエラーが発生しました。
至急確認をお願いします。

【エラー詳細】
発生時刻: ${Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss')}
コンテキスト: ${context}
エラーメッセージ: ${error.message}
エラー名: ${error.name}

【スタックトレース】
${error.stack || 'スタックトレース情報なし'}

【対応方法】
1. Google Apps Scriptの実行ログを確認
2. 該当する処理の動作状況を確認
3. 必要に応じてシステムの一時停止を検討

【システム情報】
スクリプトID: ${ScriptApp.getScriptId()}
実行ユーザー: ${Session.getActiveUser().getEmail()}

このメールは自動送信されています。

出勤管理システム
`;

  return body;
}

/**
 * 未退勤者通知メール機能
 * 退勤打刻漏れの従業員と管理者に通知
 * 
 * @param {Array} missingClockOutEmployees - 未退勤者の配列
 */
function sendMissingClockOutNotifications(missingClockOutEmployees) {
  try {
    if (!missingClockOutEmployees || missingClockOutEmployees.length === 0) {
      Logger.log('MailManager: 未退勤者はいません');
      return;
    }

    const notifications = [];

    // 各従業員への個別通知
    missingClockOutEmployees.forEach(employee => {
      const subject = '【重要】退勤打刻漏れのお知らせ';
      const body = createMissingClockOutEmployeeBody(employee);
      
      notifications.push({
        recipient: employee.email,
        subject: subject,
        body: body
      });
    });

    // 管理者への一括通知
    const adminEmail = getAdminEmail();
    if (adminEmail) {
      const subject = `【管理者通知】未退勤者一覧 (${missingClockOutEmployees.length}名)`;
      const body = createMissingClockOutAdminBody(missingClockOutEmployees);
      
      notifications.push({
        recipient: adminEmail,
        subject: subject,
        body: body
      });
    }

    sendBatchNotifications(notifications);
    Logger.log(`MailManager: 未退勤者通知送信完了 - 対象者: ${missingClockOutEmployees.length}名`);

  } catch (error) {
    Logger.log(`MailManager: 未退勤者通知送信エラー - ${error.toString()}`);
    sendErrorAlert(error, 'sendMissingClockOutNotifications');
    throw error;
  }
}

/**
 * 従業員向け未退勤通知メール本文を作成
 * 
 * @param {Object} employee - 従業員オブジェクト
 * @return {string} メール本文
 */
function createMissingClockOutEmployeeBody(employee) {
  const body = `
${employee.name} 様

お疲れ様です。

昨日（${employee.targetDate}）の退勤打刻が記録されていません。
打刻漏れの可能性がありますので、ご確認をお願いします。

【確認事項】
- 退勤時に打刻を行ったかどうか
- システムの不具合により打刻が記録されなかった可能性

【対応方法】
打刻漏れの場合は、以下のフォームから時刻修正申請を提出してください。

時刻修正申請フォーム: ${getTimeAdjustmentFormUrl()}

【打刻記録】
出勤時刻: ${employee.clockInTime || '記録なし'}
退勤時刻: 記録なし

ご不明な点がございましたら、管理者までお問い合わせください。

出勤管理システム
`;

  return body;
}

/**
 * 管理者向け未退勤者一覧メール本文を作成
 * 
 * @param {Array} employees - 未退勤者配列
 * @return {string} メール本文
 */
function createMissingClockOutAdminBody(employees) {
  let body = `
システム管理者 様

昨日の未退勤者一覧をお知らせします。

【未退勤者一覧】（${employees.length}名）
`;

  employees.forEach((employee, index) => {
    body += `
${index + 1}. ${employee.name} (${employee.employeeId})
   - 出勤時刻: ${employee.clockInTime || '記録なし'}
   - メール: ${employee.email}
   - 所属: ${employee.department || '不明'}
`;
  });

  body += `

【対応状況】
- 各従業員には個別に通知メールを送信済みです
- 時刻修正申請の提出を促しています

【確認事項】
- システム障害による打刻記録漏れの可能性
- 従業員の打刻忘れの可能性
- 残業や深夜勤務による特殊な勤務パターン

必要に応じて個別に確認をお願いします。

出勤管理システム
`;

  return body;
}

/**
 * 残業警告メール送信
 * 週次で残業時間が閾値を超えた従業員に警告
 * 
 * @param {Array} overtimeEmployees - 残業超過従業員の配列
 */
function sendOvertimeWarnings(overtimeEmployees) {
  try {
    if (!overtimeEmployees || overtimeEmployees.length === 0) {
      Logger.log('MailManager: 残業警告対象者はいません');
      return;
    }

    const notifications = [];

    // 各従業員への警告通知
    overtimeEmployees.forEach(employee => {
      const subject = '【重要】残業時間に関するお知らせ';
      const body = createOvertimeWarningBody(employee);
      
      notifications.push({
        recipient: employee.email,
        subject: subject,
        body: body
      });
    });

    // 管理者への報告
    const adminEmail = getAdminEmail();
    if (adminEmail) {
      const subject = `【管理者通知】残業警告対象者 (${overtimeEmployees.length}名)`;
      const body = createOvertimeWarningAdminBody(overtimeEmployees);
      
      notifications.push({
        recipient: adminEmail,
        subject: subject,
        body: body
      });
    }

    sendBatchNotifications(notifications);
    Logger.log(`MailManager: 残業警告送信完了 - 対象者: ${overtimeEmployees.length}名`);

  } catch (error) {
    Logger.log(`MailManager: 残業警告送信エラー - ${error.toString()}`);
    sendErrorAlert(error, 'sendOvertimeWarnings');
    throw error;
  }
}

/**
 * 従業員向け残業警告メール本文を作成
 * 
 * @param {Object} employee - 従業員オブジェクト
 * @return {string} メール本文
 */
function createOvertimeWarningBody(employee) {
  const body = `
${employee.name} 様

お疲れ様です。

過去4週間の残業時間が${employee.overtimeHours}時間となり、
労働基準法の上限（月80時間）に近づいています。

【残業時間詳細】
過去4週間の残業時間: ${employee.overtimeHours}時間
警告基準: 80時間
残り時間: ${80 - employee.overtimeHours}時間

【お願い】
- 業務の効率化や優先順位の見直しをご検討ください
- 必要に応じて上司や管理者にご相談ください
- 健康管理にもご注意ください

【参考情報】
労働基準法では、時間外労働の上限が月45時間、年360時間と定められています。
特別な事情がある場合でも月80時間、年720時間が上限となります。

ご不明な点がございましたら、管理者までお問い合わせください。

出勤管理システム
`;

  return body;
}

/**
 * 管理者向け残業警告一覧メール本文を作成
 * 
 * @param {Array} employees - 残業超過従業員配列
 * @return {string} メール本文
 */
function createOvertimeWarningAdminBody(employees) {
  let body = `
システム管理者 様

残業時間警告対象者をお知らせします。

【警告対象者一覧】（${employees.length}名）
`;

  employees.forEach((employee, index) => {
    body += `
${index + 1}. ${employee.name} (${employee.employeeId})
   - 過去4週間残業時間: ${employee.overtimeHours}時間
   - 所属: ${employee.department || '不明'}
   - 上司: ${employee.supervisor || '不明'}
`;
  });

  body += `

【対応状況】
- 各従業員には個別に警告メールを送信済みです
- 業務効率化の検討を促しています

【管理者対応事項】
- 業務量の調整検討
- 人員配置の見直し
- 労働環境の改善検討

労働基準法遵守のため、適切な対応をお願いします。

出勤管理システム
`;

  return body;
}

/**
 * 管理者メールアドレスを取得
 * 
 * @return {string} 管理者メールアドレス
 */
function getAdminEmail() {
  try {
    // Config.gsから管理者メールアドレスを取得
    if (typeof getConfig === 'function') {
      return getConfig('SYSTEM', 'ADMIN_EMAIL');
    }
    
    // フォールバック: 直接設定値を参照
    if (typeof SYSTEM_CONFIG !== 'undefined') {
      return SYSTEM_CONFIG.ADMIN_EMAIL;
    }
    
    // 最終フォールバック: スクリプト所有者のメールアドレス
    return Session.getActiveUser().getEmail();
    
  } catch (error) {
    Logger.log(`MailManager: 管理者メールアドレス取得エラー - ${error.toString()}`);
    return Session.getActiveUser().getEmail();
  }
}

/**
 * 承認シートのURLを取得
 * 
 * @return {string} 承認シートURL
 */
function getApprovalSheetUrl() {
  try {
    const spreadsheetId = getSpreadsheetId();
    const sheetId = getSheetId('Request_Responses');
    return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetId}`;
  } catch (error) {
    Logger.log(`MailManager: 承認シートURL取得エラー - ${error.toString()}`);
    return 'https://docs.google.com/spreadsheets/';
  }
}

/**
 * 時刻修正申請フォームのURLを取得
 * 
 * @return {string} フォームURL
 */
function getTimeAdjustmentFormUrl() {
  try {
    // Config.gsからフォームURLを取得
    if (typeof getConfig === 'function') {
      const systemConfig = getConfig('SYSTEM');
      return systemConfig.TIME_ADJUSTMENT_FORM_URL || 'フォームURLが設定されていません';
    }
    
    return 'フォームURLが設定されていません';
    
  } catch (error) {
    Logger.log(`MailManager: フォームURL取得エラー - ${error.toString()}`);
    return 'フォームURLの取得に失敗しました';
  }
}

/**
 * スプレッドシートIDを取得
 * 
 * @return {string} スプレッドシートID
 */
function getSpreadsheetId() {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

/**
 * シートIDを取得
 * 
 * @param {string} sheetName - シート名
 * @return {number} シートID
 */
function getSheetId(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet ? sheet.getSheetId() : 0;
}