// ========== ワークフロー関連テストヘルパー関数 ==========

/**
 * 統一された関数存在チェックヘルパー
 * @param {string[]} functionNames - チェックする関数名の配列
 * @returns {Object} チェック結果オブジェクト
 */
function checkFunctionExistence(functionNames) {
  try {
    const results = {};
    let allExist = true;
    
    functionNames.forEach(funcName => {
      try {
        // 安全な関数存在チェック - thisコンテキストとglobalThisの両方をチェック
        const func = this[funcName] || globalThis[funcName];
        results[funcName] = {
          exists: typeof func === 'function',
          type: typeof func
        };
        if (typeof func !== 'function') {
          allExist = false;
        }
      } catch (error) {
        results[funcName] = {
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
 * エラーハンドリング関数の存在確認
 */
function testErrorHandlingFunctionExists() {
  const errorFunctions = ['withErrorHandling', 'sendErrorAlert', 'logError'];
  return checkFunctionExistence(errorFunctions);
}

/**
 * エラーログ機能の確認
 */
function testErrorLogging() {
  try {
    // withErrorHandling関数の存在チェック
    const withErrorHandlingFunc = this['withErrorHandling'] || globalThis['withErrorHandling'];
    
    if (typeof withErrorHandlingFunc !== 'function') {
      return {
        success: false,
        error: 'withErrorHandling関数が定義されていません',
        withErrorHandlingExists: false
      };
    }
    
    // エラーログ機能のテスト（実際のエラーは発生させない）
    const testResult = withErrorHandlingFunc(() => {
      return { success: true, message: 'テスト成功' };
    }, 'TestErrorLogging', 'LOW');
    
    return {
      success: testResult && testResult.success,
      testResult: testResult,
      withErrorHandlingExists: true
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      withErrorHandlingExists: false
    };
  }
}

/**
 * エラー通知設定の確認
 */
function testErrorNotificationSetup() {
  try {
    const adminEmail = getAdminEmail();
    const isConfigured = adminEmail && adminEmail !== 'admin@example.com';
    
    // sendErrorAlert関数の存在確認
    const functionCheck = checkFunctionExistence(['sendErrorAlert']);
    const sendErrorAlertExists = functionCheck.results['sendErrorAlert']?.exists || false;
    
    return {
      success: isConfigured && sendErrorAlertExists,
      adminEmailConfigured: isConfigured,
      sendErrorAlertExists: sendErrorAlertExists
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * バッチ処理エラーハンドリングの確認
 */
function testBatchProcessingErrorHandling() {
  try {
    // processBatch関数の存在確認
    const functionCheck = checkFunctionExistence(['processBatch']);
    const processBatchExists = functionCheck.results['processBatch']?.exists || false;
    
    return {
      success: processBatchExists,
      processBatchExists: processBatchExists
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 認証ワークフローの確認
 */
function testAuthenticationWorkflow() {
  const authFunctions = ['authenticateUser', 'validateEmployeeAccess', 'getEmployeeInfo'];
  return checkFunctionExistence(authFunctions);
}

/**
 * 承認ワークフローの設定確認
 */
function testApprovalWorkflowSetup() {
  const approvalFunctions = ['onRequestSubmit', 'processApprovalRequest', 'onRequestResponsesEdit'];
  return checkFunctionExistence(approvalFunctions);
}

/**
 * 計算ワークフローの確認
 */
function testCalculationWorkflow() {
  const calcFunctions = ['calculateDailyWorkTime', 'calculateOvertime', 'updateDailySummary'];
  return checkFunctionExistence(calcFunctions);
}

/**
 * 通知ワークフローの設定確認
 */
function testNotificationWorkflowSetup() {
  const notificationFunctions = ['sendBatchNotification', 'sendErrorAlert', 'sendOvertimeWarning'];
  return checkFunctionExistence(notificationFunctions);
} 