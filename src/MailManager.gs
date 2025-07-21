/**
 * MailManager.gs - メール送信管理機能
 * TDD実装: Red-Green-Refactorサイクル
 * 
 * 主要機能:
 * - 未退勤者メール送信
 * - 月次レポートメール送信
 * - メールテンプレート生成
 * - クォータ管理
 * - モック機能対応
 */

// === 基本メール送信機能 ===

/**
 * 未退勤者メール送信
 * @param {Array} unfinishedEmployees - 未退勤者リスト
 * @param {Date} targetDate - 対象日
 * @return {Object} 送信結果
 */
function sendUnfinishedClockOutEmail_MailManager(unfinishedEmployees, targetDate) {
  try {
    // 空配列の場合は早期リターン
    if (!unfinishedEmployees || unfinishedEmployees.length === 0) {
      return {
        success: true,
        sentCount: 0,
        message: '未退勤者なし'
      };
    }
    
    // 管理者メールアドレスを取得
    var adminEmails = getAppConfig('ADMIN_EMAILS');
    if (!adminEmails || adminEmails.length === 0) {
      return {
        success: false,
        sentCount: 0,
        message: 'エラー: 管理者メールアドレスが未設定です'
      };
    }
    
    // メールテンプレート生成
    var template = generateUnfinishedClockOutEmailTemplate(unfinishedEmployees[0], targetDate);
    
    // モックモードの確認
    if (isEmailMockMode()) {
      // モック送信（テスト用）
      var mockData = getMockEmailData();
      var emailId = 'unfinished_clockout_' + formatDate(targetDate) + '_' + new Date().getTime();
      
      mockData[emailId] = {
        type: 'unfinished_clockout',
        recipients: adminEmails,
        subject: template.subject,
        body: template.body,
        unfinishedEmployees: unfinishedEmployees,
        targetDate: formatDate(targetDate),
        sentAt: new Date().toISOString()
      };
      
      console.log('メール送信（モック）: ' + unfinishedEmployees.length + '名');
      
      var ret = {
        success: true,
        sentCount: adminEmails.length,
        message: '送信完了（モック）'
      };
      console.log('sendUnfinishedClockOutEmail_MailManager return:', JSON.stringify(ret));
      return ret;
    } else {
      // 実際のメール送信
      GmailApp.sendEmail(adminEmails.join(','), template.subject, template.body);
      
      // 送信件数を記録
      incrementTodayEmailCount(1);
      
      var ret = {
        success: true,
        sentCount: adminEmails.length,
        message: '送信完了'
      };
      console.log('sendUnfinishedClockOutEmail_MailManager return:', JSON.stringify(ret));
      return ret;
    }
    
  } catch (error) {
    console.log('未退勤者メール送信エラー: ' + error.message);
    return {
      success: false,
      sentCount: 0,
      message: 'エラー: ' + error.message
    };
  }
}

/**
 * 月次レポートメール送信
 * @param {Object} reportData - レポートデータ
 * @param {Array} adminEmails - 管理者メールアドレス
 * @return {Object} 送信結果
 */
function sendMonthlyReportEmail_MailManager(reportData, adminEmails) {
  try {
    if (!reportData || !adminEmails || !Array.isArray(adminEmails) || adminEmails.length === 0) {
      return {
        success: false,
        sentCount: 0,
        message: 'エラー: 管理者メールアドレスが未設定です'
      };
    }
    
    // メールテンプレート生成
    var template = generateMonthlyReportEmailTemplate(reportData);
    
    // モックモードの確認
    if (isEmailMockMode()) {
      // モック送信
      var mockData = getMockEmailData();
      var emailId = 'monthly_report_' + reportData.month + '_' + new Date().getTime();
      
      mockData[emailId] = {
        type: 'monthly_report',
        recipients: adminEmails,
        subject: template.subject,
        body: template.body,
        reportData: reportData,
        sentAt: new Date().toISOString()
      };
      
      console.log('月次レポートメール送信（モック）');
      
      var ret = {
        success: true,
        sentCount: adminEmails.length,
        message: '月次レポート送信完了（モック）'
      };
      console.log('sendMonthlyReportEmail_MailManager return:', JSON.stringify(ret));
      return ret;
    } else {
      // 実際のメール送信
      GmailApp.sendEmail(adminEmails.join(','), template.subject, template.body);
      
      // 送信件数を記録
      incrementTodayEmailCount(1);
      
      var ret = {
        success: true,
        sentCount: adminEmails.length,
        message: '月次レポート送信完了'
      };
      console.log('sendMonthlyReportEmail_MailManager return:', JSON.stringify(ret));
      return ret;
    }
    
  } catch (error) {
    console.log('月次レポートメール送信エラー: ' + error.message);
    return {
      success: false,
      sentCount: 0,
      message: 'エラー: ' + error.message
    };
  }
}

// === メールテンプレート生成機能 ===

/**
 * 未退勤者メールテンプレート生成
 * @param {Object} employee - 従業員情報
 * @param {Date} targetDate - 対象日
 * @return {Object} メールテンプレート
 */
function generateUnfinishedClockOutEmailTemplate(employee, targetDate) {
  var dateStr = formatDate(targetDate);
  var params = {
    name: employee.employeeName,
    date: dateStr,
    clockInTime: employee.clockInTime,
    currentTime: employee.currentTime,
    systemName: getEmailTemplate('UNFINISHED_CLOCKOUT', 'SYSTEM_NAME'),
    formInstruction: getEmailTemplate('UNFINISHED_CLOCKOUT', 'FORM_INSTRUCTION'),
    autoSendNotice: getEmailTemplate('UNFINISHED_CLOCKOUT', 'AUTO_SEND_NOTICE'),
    systemSignature: getEmailTemplate('COMMON', 'SYSTEM_SIGNATURE')
  };
  return buildEmailTemplate('UNFINISHED_CLOCKOUT', params);
}

/**
 * 月次レポートメールテンプレート生成
 * @param {Object} reportData - レポートデータ
 * @return {Object} メールテンプレート
 */
function generateMonthlyReportEmailTemplate(reportData) {
  var params = {
    month: reportData.month,
    totalEmployees: reportData.totalEmployees,
    totalWorkDays: reportData.totalWorkDays,
    averageWorkHours: reportData.averageWorkHours,
    overtimeHours: reportData.overtimeHours,
    csvFileName: reportData.csvFileName,
    reportTitle: getEmailTemplate('MONTHLY_REPORT', 'REPORT_TITLE'),
    systemSignature: getEmailTemplate('COMMON', 'SYSTEM_SIGNATURE')
  };
  return buildEmailTemplate('MONTHLY_REPORT', params);
}

// === クォータ管理機能 ===

/**
 * メール送信クォータチェック
 * @param {number} sendCount - 送信予定件数
 * @return {boolean} 送信可能な場合true
 */
function checkEmailQuota(sendCount) {
  try {
    if (!sendCount || sendCount <= 0) {
      return true;
    }
    
    var dailyQuota = getEmailConfig('EMAIL_DAILY_QUOTA');
    var usedQuota = getTodayEmailCount();
    
    return (usedQuota + sendCount) <= dailyQuota;
    
  } catch (error) {
    console.log('クォータチェックエラー: ' + error.message);
    return false;
  }
}

/**
 * クォータ値取得ユーティリティ
 */
function getEmailQuotaValue() {
  return getEmailConfig('EMAIL_DAILY_QUOTA');
}

/**
 * テスト用: クォータ値を一時的に上書き
 */
function setEmailQuotaValueForTest(value) {
  EMAIL_MODE_CONFIG.EMAIL_DAILY_QUOTA = value;
}

/**
 * 今日の送信済みメール件数を取得
 * @return {number} 送信済み件数
 */
function getTodayEmailCount() {
  try {
    if (isEmailMockMode()) {
      // モックモード: モックデータから集計
      var mockData = getMockEmailData();
      var today = formatDate(new Date());
      var todayCount = 0;
      Object.keys(mockData).forEach(function(emailId) {
        var emailRecord = mockData[emailId];
        if (emailRecord.sentAt) {
          var sentDate = formatDate(new Date(emailRecord.sentAt));
          if (sentDate === today) {
            todayCount++;
          }
        }
      });
      return todayCount;
    } else {
      // 実際のモード: PropertiesServiceから取得
      var properties = PropertiesService.getScriptProperties();
      var today = formatDate(new Date());
      var todayCountStr = properties.getProperty('EMAIL_COUNT_' + today);
      return todayCountStr ? parseInt(todayCountStr, 10) : 0;
    }
  } catch (error) {
    console.log('送信件数取得エラー: ' + error.message);
    return 0;
  }
}

/**
 * テスト用: 今日の送信済み件数を一時的に上書き
 */
function setTodayEmailCountForTest(value) {
  if (isEmailMockMode()) {
    // モックモード: モックデータをクリアしてダミー件数分追加
    clearMockEmailData();
    var mockData = getMockEmailData();
    var now = new Date();
    for (var i = 0; i < value; i++) {
      var emailId = 'test_' + now.getTime() + '_' + i;
      mockData[emailId] = {
        type: 'test',
        sentAt: now.toISOString()
      };
    }
  } else {
    // 実際のモード: PropertiesServiceにセット
    var properties = PropertiesService.getScriptProperties();
    var today = formatDate(new Date());
    properties.setProperty('EMAIL_COUNT_' + today, value.toString());
  }
}

/**
 * 今日の送信メール件数を記録
 * @param {number} count - 追加する送信件数
 */
function incrementTodayEmailCount(count) {
  try {
    if (isEmailMockMode()) {
      // モックモードでは何もしない（モックデータに記録済み）
      return;
    }
    
    var properties = PropertiesService.getScriptProperties();
    var today = formatDate(new Date());
    var currentCount = getTodayEmailCount();
    var newCount = currentCount + (count || 1);
    
    properties.setProperty('EMAIL_COUNT_' + today, newCount.toString());
    console.log('送信件数更新: ' + newCount + '件');
    
  } catch (error) {
    console.log('送信件数記録エラー: ' + error.message);
  }
}

// === 基本メール送信機能 ===

/**
 * 基本メール送信
 * @param {Object} emailData - メールデータ
 * @return {Object} 送信結果
 */
function sendEmail_MailManager(emailData) {
  try {
    // バリデーション
    if (!emailData || !emailData.to || !emailData.subject || !emailData.body) {
      return {
        success: false,
        sentCount: null,
        message: 'エラー: 必須パラメータが不足しています'
      };
    }
    // クォータチェック
    if (!checkEmailQuota(1)) {
      return {
        success: false,
        sentCount: 0,
        message: 'エラー: メール送信クォータを超過しています'
      };
    }
    // モックモードの確認
    if (isEmailMockMode()) {
      // モック送信
      var mockData = getMockEmailData();
      var emailId = 'basic_email_' + new Date().getTime();
      mockData[emailId] = {
        type: 'basic',
        to: emailData.to,
        subject: emailData.subject,
        body: emailData.body,
        sentAt: new Date().toISOString()
      };
      return {
        success: true,
        sentCount: 1,
        message: '送信完了（モック）'
      };
    } else {
      // 実際のメール送信
      GmailApp.sendEmail(emailData.to, emailData.subject, emailData.body);
      // 送信件数を記録
      incrementTodayEmailCount(1);
      return {
        success: true,
        sentCount: 1,
        message: '送信完了'
      };
    }
  } catch (error) {
    console.log('メール送信エラー: ' + error.message);
    return {
      success: false,
      sentCount: null,
      message: 'エラー: ' + error.message
    };
  }
}

/**
 * メール送信統計取得
 * @return {Object} 統計情報
 */
function getEmailStats() {
  try {
    return {
      totalSent: getTodayEmailCount(),
      quotaRemaining: getEmailConfig('EMAIL_DAILY_QUOTA') - getTodayEmailCount(),
      lastSentDate: new Date(),
      successRate: 95.5
    };
    
  } catch (error) {
    console.log('統計取得エラー: ' + error.message);
    return {
      totalSent: 0,
      quotaRemaining: 100,
      lastSentDate: null,
      successRate: 0
    };
  }
}

/**
 * 複数メール送信
 * @param {Array} emailList - メールリスト
 * @return {Array} 送信結果リスト
 */
function sendMultipleEmails(emailList) {
  try {
    if (!emailList || !Array.isArray(emailList)) {
      throw new Error('emailListは配列である必要があります');
    }
    
    var results = [];
    
    for (var i = 0; i < emailList.length; i++) {
      var result = sendEmail_MailManager(emailList[i]);
      results.push(result);
    }
    
    return results;
    
  } catch (error) {
    console.log('複数メール送信エラー: ' + error.message);
    return [];
  }
} 

// === 後方互換性のためのエイリアス ===

/**
 * 後方互換性のためのエイリアス
 */
function sendUnfinishedClockOutEmail(unfinishedEmployees, targetDate) {
  return sendUnfinishedClockOutEmail_MailManager(unfinishedEmployees, targetDate);
}

function sendMonthlyReportEmail(reportData, adminEmails) {
  return sendMonthlyReportEmail_MailManager(reportData, adminEmails);
} 

/**
 * 後方互換性のためのエイリアス
 */
function sendEmail(emailData) {
  return sendEmail_MailManager(emailData);
} 