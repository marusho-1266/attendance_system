// ========== ワークフロー関連テストヘルパー関数 ==========

/**
 * エラーハンドリング関数の存在確認
 */
function testErrorHandlingFunctionExists() {
  try {
    const errorFunctions = ['withErrorHandling', 'sendErrorAlert', 'logError'];
    const results = {};
    let allExist = true;
    
    errorFunctions.forEach(funcName => {
      try {
        // 安全な関数存在チェック - eval()の代わりにグローバルスコープから直接アクセス
        const func = this[funcName] || (typeof window !== 'undefined' ? window[funcName] : null);
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
 * エラーログ機能の確認
 */
function testErrorLogging() {
  try {
    // エラーログ機能のテスト（実際のエラーは発生させない）
    const testResult = withErrorHandling(() => {
      return { success: true, message: 'テスト成功' };
    }, 'TestErrorLogging', 'LOW');
    
    return {
      success: testResult && testResult.success,
      testResult: testResult
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
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
    let sendErrorAlertExists = false;
    try {
      // 安全な関数存在チェック - eval()の代わりにグローバルスコープから直接アクセス
      const func = this['sendErrorAlert'] || (typeof window !== 'undefined' ? window['sendErrorAlert'] : null);
      sendErrorAlertExists = typeof func === 'function';
    } catch (error) {
      sendErrorAlertExists = false;
    }
    
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
    let processBatchExists = false;
    try {
      // 安全な関数存在チェック - eval()の代わりにグローバルスコープから直接アクセス
      const func = this['processBatch'] || (typeof window !== 'undefined' ? window['processBatch'] : null);
      processBatchExists = typeof func === 'function';
    } catch (error) {
      processBatchExists = false;
    }
    
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
  try {
    const authFunctions = ['authenticateUser', 'validateEmployeeAccess', 'getEmployeeInfo'];
    const results = {};
    let allExist = true;
    
    authFunctions.forEach(funcName => {
      try {
        // 安全な関数存在チェック - eval()の代わりにグローバルスコープから直接アクセス
        const func = this[funcName] || (typeof window !== 'undefined' ? window[funcName] : null);
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
 * 承認ワークフローの設定確認
 */
function testApprovalWorkflowSetup() {
  try {
    const approvalFunctions = ['onRequestSubmit', 'processApprovalRequest', 'onRequestResponsesEdit'];
    const results = {};
    let allExist = true;
    
    approvalFunctions.forEach(funcName => {
      try {
        const func = eval(funcName);
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
 * 計算ワークフローの確認
 */
function testCalculationWorkflow() {
  try {
    const calcFunctions = ['calculateDailyWorkTime', 'calculateOvertime', 'updateDailySummary'];
    const results = {};
    let allExist = true;
    
    calcFunctions.forEach(funcName => {
      try {
        const func = eval(funcName);
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
 * 通知ワークフローの設定確認
 */
function testNotificationWorkflowSetup() {
  try {
    const notificationFunctions = ['sendBatchNotification', 'sendErrorAlert', 'sendOvertimeWarning'];
    const results = {};
    let allExist = true;
    
    notificationFunctions.forEach(funcName => {
      try {
        const func = eval(funcName);
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