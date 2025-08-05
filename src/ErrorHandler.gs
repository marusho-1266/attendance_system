/**
 * エラーハンドリングシステム
 * 
 * 機能:
 * - 包括的なエラーハンドリング
 * - エラーログ記録
 * - 管理者への即時エラー通知
 * - クォータ制限対応のバッチ処理
 * - 実行時間制限対応のチャンク処理
 * 
 * 要件: 6.1, 6.3, 6.4, 6.5
 */

// ========== エラー処理設定 ==========

/**
 * エラー処理設定
 */
const ERROR_CONFIG = {
  // 実行時間制限（ミリ秒）
  MAX_EXECUTION_TIME: 5 * 60 * 1000, // 5分
  
  // チャンク処理のデフォルトサイズ
  DEFAULT_CHUNK_SIZE: 50,
  
  // バッチ処理のデフォルトサイズ
  DEFAULT_BATCH_SIZE: 20,
  
  // エラー通知の間隔（ミリ秒）
  ERROR_NOTIFICATION_INTERVAL: 5 * 60 * 1000, // 5分
  
  // ログ保持期間（日）
  LOG_RETENTION_DAYS: 30,
  
  // 重要度レベル
  SEVERITY_LEVELS: {
    LOW: 'LOW',
    MEDIUM: 'MEDIUM',
    HIGH: 'HIGH',
    CRITICAL: 'CRITICAL'
  }
};

// ========== エラーログ記録機能 ==========

/**
 * エラーログを記録する
 * @param {Error} error - エラーオブジェクト
 * @param {string} context - エラー発生コンテキスト
 * @param {string} severity - 重要度 (LOW/MEDIUM/HIGH/CRITICAL)
 * @param {Object} additionalInfo - 追加情報
 */
function logError(error, context, severity = 'MEDIUM', additionalInfo = {}) {
  try {
    // エラーログシートの取得または作成
    const logSheet = getOrCreateErrorLogSheet();
    
    // エラー情報を構築
    const errorInfo = {
      timestamp: new Date(),
      context: context || 'Unknown',
      severity: severity,
      errorName: error.name || 'Error',
      errorMessage: error.message || 'Unknown error',
      stackTrace: error.stack || 'No stack trace',
      userEmail: Session.getActiveUser().getEmail() || 'Unknown',
      scriptId: ScriptApp.getScriptId(),
      additionalInfo: JSON.stringify(additionalInfo)
    };
    
    // ログエントリを作成
    const logEntry = [
      errorInfo.timestamp,
      errorInfo.context,
      errorInfo.severity,
      errorInfo.errorName,
      errorInfo.errorMessage,
      errorInfo.stackTrace,
      errorInfo.userEmail,
      errorInfo.scriptId,
      errorInfo.additionalInfo
    ];
    
    // ログシートに記録
    logSheet.appendRow(logEntry);
    
    // コンソールログにも記録
    Logger.log(`[${severity}] ${context}: ${error.message}`);
    
    // 重要度が高い場合は即座に通知
    if (severity === 'HIGH' || severity === 'CRITICAL') {
      sendImmediateErrorNotification(errorInfo);
    }
    
  } catch (logError) {
    // ログ記録自体が失敗した場合はコンソールログのみ
    Logger.log(`エラーログ記録失敗: ${logError.toString()}`);
    Logger.log(`元のエラー: ${error.toString()}`);
  }
}

/**
 * エラーログシートを取得または作成
 * @return {Sheet} エラーログシート
 */
function getOrCreateErrorLogSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = spreadsheet.getSheetByName('Error_Log');
    
    if (!logSheet) {
      // エラーログシートを作成
      logSheet = spreadsheet.insertSheet('Error_Log');
      
      // ヘッダー行を設定
      const headers = [
        'タイムスタンプ',
        'コンテキスト',
        '重要度',
        'エラー名',
        'エラーメッセージ',
        'スタックトレース',
        'ユーザー',
        'スクリプトID',
        '追加情報'
      ];
      
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダー行のフォーマット
      const headerRange = logSheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f0f0f0');
      headerRange.setWrap(true);
      
      // 列幅を調整
      logSheet.setColumnWidth(1, 150); // タイムスタンプ
      logSheet.setColumnWidth(2, 120); // コンテキスト
      logSheet.setColumnWidth(3, 80);  // 重要度
      logSheet.setColumnWidth(4, 100); // エラー名
      logSheet.setColumnWidth(5, 200); // エラーメッセージ
      logSheet.setColumnWidth(6, 300); // スタックトレース
      logSheet.setColumnWidth(7, 150); // ユーザー
      logSheet.setColumnWidth(8, 150); // スクリプトID
      logSheet.setColumnWidth(9, 200); // 追加情報
    }
    
    return logSheet;
    
  } catch (error) {
    Logger.log(`エラーログシート作成エラー: ${error.toString()}`);
    throw error;
  }
}

// ========== 管理者への即時エラー通知 ==========

/**
 * 管理者への即時エラー通知
 * @param {Object} errorInfo - エラー情報オブジェクト
 */
function sendImmediateErrorNotification(errorInfo) {
  try {
    // 通知頻度制限チェック
    if (!shouldSendErrorNotification(errorInfo.context)) {
      Logger.log(`エラー通知スキップ（頻度制限）: ${errorInfo.context}`);
      return;
    }
    
    const adminEmail = getAdminEmail();
    if (!adminEmail) {
      Logger.log('管理者メールアドレスが設定されていません');
      return;
    }
    
    const subject = `【${errorInfo.severity}】出勤管理システムエラー: ${errorInfo.context}`;
    const body = createErrorNotificationBody(errorInfo);
    
    // メール送信
    GmailApp.sendEmail(adminEmail, subject, body);
    
    // 通知履歴を記録
    recordErrorNotification(errorInfo.context);
    
    Logger.log(`エラー通知送信完了: ${errorInfo.context}`);
    
  } catch (error) {
    Logger.log(`エラー通知送信失敗: ${error.toString()}`);
  }
}

/**
 * エラー通知メール本文を作成
 * @param {Object} errorInfo - エラー情報
 * @return {string} メール本文
 */
function createErrorNotificationBody(errorInfo) {
  return `
システム管理者 様

出勤管理システムで${errorInfo.severity}レベルのエラーが発生しました。
至急確認をお願いします。

【エラー詳細】
発生時刻: ${Utilities.formatDate(errorInfo.timestamp, 'JST', 'yyyy/MM/dd HH:mm:ss')}
コンテキスト: ${errorInfo.context}
重要度: ${errorInfo.severity}
エラー名: ${errorInfo.errorName}
エラーメッセージ: ${errorInfo.errorMessage}

【実行環境】
ユーザー: ${errorInfo.userEmail}
スクリプトID: ${errorInfo.scriptId}

【スタックトレース】
${errorInfo.stackTrace}

【追加情報】
${errorInfo.additionalInfo}

【対応方法】
1. Google Apps Scriptの実行ログを確認
2. エラーログシート（Error_Log）で詳細を確認
3. 必要に応じてシステムの一時停止を検討
4. 問題解決後、関連する処理の再実行を検討

【システム情報】
エラーログシート: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=${getSheetId('Error_Log')}

このメールは自動送信されています。

出勤管理システム
`;
}

/**
 * エラー通知の送信可否を判定（頻度制限）
 * @param {string} context - エラーコンテキスト
 * @return {boolean} 送信可否
 */
function shouldSendErrorNotification(context) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const lastNotificationKey = `last_error_notification_${context}`;
    const lastNotificationTime = properties.getProperty(lastNotificationKey);
    
    if (!lastNotificationTime) {
      return true; // 初回通知
    }
    
    const lastTime = new Date(lastNotificationTime);
    const now = new Date();
    const timeDiff = now.getTime() - lastTime.getTime();
    
    return timeDiff >= ERROR_CONFIG.ERROR_NOTIFICATION_INTERVAL;
    
  } catch (error) {
    Logger.log(`通知頻度チェックエラー: ${error.toString()}`);
    return true; // エラー時は通知を許可
  }
}

/**
 * エラー通知履歴を記録
 * @param {string} context - エラーコンテキスト
 */
function recordErrorNotification(context) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const lastNotificationKey = `last_error_notification_${context}`;
    properties.setProperty(lastNotificationKey, new Date().toISOString());
  } catch (error) {
    Logger.log(`通知履歴記録エラー: ${error.toString()}`);
  }
}

// ========== 包括的エラーハンドリングラッパー ==========

/**
 * 関数を包括的エラーハンドリングでラップ
 * @param {Function} func - 実行する関数
 * @param {string} context - エラーコンテキスト
 * @param {string} severity - エラー重要度
 * @param {Object} options - オプション設定
 * @return {*} 関数の戻り値
 */
function withErrorHandling(func, context, severity = 'MEDIUM', options = {}) {
  const startTime = new Date();
  
  try {
    Logger.log(`${context} 開始`);
    
    const result = func();
    
    const endTime = new Date();
    const executionTime = endTime.getTime() - startTime.getTime();
    
    Logger.log(`${context} 完了 (実行時間: ${executionTime}ms)`);
    
    return result;
    
  } catch (error) {
    const endTime = new Date();
    const executionTime = endTime.getTime() - startTime.getTime();
    
    // エラー情報を記録
    logError(error, context, severity, {
      executionTime: executionTime,
      options: options
    });
    
    // オプションで再スローを制御
    if (options.suppressError !== true) {
      throw error;
    }
    
    return options.defaultValue || null;
  }
}

/**
 * 非同期処理用エラーハンドリングラッパー
 * @param {Function} func - 実行する関数
 * @param {string} context - エラーコンテキスト
 * @param {string} severity - エラー重要度
 * @param {Object} options - オプション設定
 * @return {Promise} Promise結果
 */
function withAsyncErrorHandling(func, context, severity = 'MEDIUM', options = {}) {
  return new Promise((resolve, reject) => {
    try {
      const result = withErrorHandling(func, context, severity, options);
      resolve(result);
    } catch (error) {
      if (options.suppressError === true) {
        resolve(options.defaultValue || null);
      } else {
        reject(error);
      }
    }
  });
}

// ========== クォータ制限対応のバッチ処理 ==========

/**
 * クォータ制限対応のバッチ処理
 * @param {Array} items - 処理対象のアイテム配列
 * @param {Function} processor - 各アイテムを処理する関数
 * @param {Object} options - バッチ処理オプション
 * @return {Object} 処理結果
 */
function processBatch(items, processor, options = {}) {
  const batchSize = options.batchSize || ERROR_CONFIG.DEFAULT_BATCH_SIZE;
  const context = options.context || 'BatchProcess';
  const delayBetweenBatches = options.delay || 1000; // 1秒
  
  return withErrorHandling(() => {
    if (!Array.isArray(items)) {
      throw new Error('処理対象はArray型である必要があります');
    }
    
    if (typeof processor !== 'function') {
      throw new Error('プロセッサーは関数である必要があります');
    }
    
    const results = [];
    const errors = [];
    let processedCount = 0;
    
    Logger.log(`バッチ処理開始: 総アイテム数=${items.length}, バッチサイズ=${batchSize}`);
    
    // バッチごとに処理
    for (let i = 0; i < items.length; i += batchSize) {
      const batch = items.slice(i, i + batchSize);
      const batchNumber = Math.floor(i / batchSize) + 1;
      const totalBatches = Math.ceil(items.length / batchSize);
      
      Logger.log(`バッチ ${batchNumber}/${totalBatches} 処理中 (${batch.length}件)`);
      
      // バッチ内の各アイテムを処理
      batch.forEach((item, index) => {
        try {
          const result = processor(item, i + index);
          results.push(result);
          processedCount++;
        } catch (error) {
          errors.push({
            item: item,
            index: i + index,
            error: error
          });
          
          logError(error, `${context}.BatchItem`, 'LOW', {
            itemIndex: i + index,
            batchNumber: batchNumber
          });
        }
      });
      
      // バッチ間の遅延（クォータ制限対策）
      if (i + batchSize < items.length && delayBetweenBatches > 0) {
        Utilities.sleep(delayBetweenBatches);
      }
      
      // 実行時間チェック
      if (options.maxExecutionTime && 
          (new Date().getTime() - options.startTime) > options.maxExecutionTime) {
        Logger.log(`実行時間制限により処理を中断: ${i + batchSize}/${items.length}`);
        break;
      }
    }
    
    Logger.log(`バッチ処理完了: 成功=${processedCount}件, エラー=${errors.length}件`);
    
    return {
      success: true,
      processedCount: processedCount,
      totalCount: items.length,
      results: results,
      errors: errors
    };
    
  }, context, 'HIGH');
}

// ========== 実行時間制限対応のチャンク処理 ==========

/**
 * 実行時間制限対応のチャンク処理
 * @param {Array} items - 処理対象のアイテム配列
 * @param {Function} processor - 各アイテムを処理する関数
 * @param {Object} options - チャンク処理オプション
 * @return {Object} 処理結果
 */
function processChunks(items, processor, options = {}) {
  const chunkSize = options.chunkSize || ERROR_CONFIG.DEFAULT_CHUNK_SIZE;
  const maxExecutionTime = options.maxExecutionTime || ERROR_CONFIG.MAX_EXECUTION_TIME;
  const context = options.context || 'ChunkProcess';
  const startTime = new Date().getTime();
  
  return withErrorHandling(() => {
    if (!Array.isArray(items)) {
      throw new Error('処理対象はArray型である必要があります');
    }
    
    if (typeof processor !== 'function') {
      throw new Error('プロセッサーは関数である必要があります');
    }
    
    const results = [];
    const errors = [];
    let processedCount = 0;
    let lastProcessedIndex = 0;
    
    Logger.log(`チャンク処理開始: 総アイテム数=${items.length}, チャンクサイズ=${chunkSize}`);
    
    // チャンクごとに処理
    for (let i = 0; i < items.length; i += chunkSize) {
      const currentTime = new Date().getTime();
      const elapsedTime = currentTime - startTime;
      
      // 実行時間制限チェック
      if (elapsedTime > maxExecutionTime) {
        Logger.log(`実行時間制限により処理を中断: ${i}/${items.length} (経過時間: ${elapsedTime}ms)`);
        
        // 継続処理をスケジュール
        scheduleChunkContinuation(items, processor, options, i);
        break;
      }
      
      const chunk = items.slice(i, i + chunkSize);
      const chunkNumber = Math.floor(i / chunkSize) + 1;
      const totalChunks = Math.ceil(items.length / chunkSize);
      
      Logger.log(`チャンク ${chunkNumber}/${totalChunks} 処理中 (${chunk.length}件)`);
      
      // チャンク内の各アイテムを処理
      chunk.forEach((item, index) => {
        try {
          const result = processor(item, i + index);
          results.push(result);
          processedCount++;
          lastProcessedIndex = i + index;
        } catch (error) {
          errors.push({
            item: item,
            index: i + index,
            error: error
          });
          
          logError(error, `${context}.ChunkItem`, 'LOW', {
            itemIndex: i + index,
            chunkNumber: chunkNumber
          });
        }
      });
    }
    
    const endTime = new Date().getTime();
    const totalExecutionTime = endTime - startTime;
    
    Logger.log(`チャンク処理完了: 成功=${processedCount}件, エラー=${errors.length}件, 実行時間=${totalExecutionTime}ms`);
    
    return {
      success: true,
      processedCount: processedCount,
      totalCount: items.length,
      lastProcessedIndex: lastProcessedIndex,
      results: results,
      errors: errors,
      executionTime: totalExecutionTime,
      completed: processedCount === items.length
    };
    
  }, context, 'HIGH');
}

/**
 * チャンク処理の継続をスケジュール
 * @param {Array} items - 処理対象のアイテム配列
 * @param {Function} processor - 処理関数
 * @param {Object} options - オプション
 * @param {number} startIndex - 開始インデックス
 */
function scheduleChunkContinuation(items, processor, options, startIndex) {
  try {
    // 継続処理用のトリガーを設定
    const trigger = ScriptApp.newTrigger('continueChunkProcessing')
      .timeBased()
      .after(60000) // 1分後に実行
      .create();
    
    // 関数名の検証
    const processorName = processor.name;
    if (!processorName) {
      Logger.log('警告: 匿名関数は継続処理で使用できません。処理を中断します。');
      logError(new Error('匿名関数は継続処理で使用できません'), 'scheduleChunkContinuation.AnonymousFunction', 'HIGH', {
        startIndex: startIndex,
        itemsCount: items.length
      });
      return;
    }
    
    // 関数がFUNCTION_MAPPINGに登録されているかチェック
    if (!FUNCTION_MAPPING[processorName]) {
      Logger.log(`警告: 関数 '${processorName}' がFUNCTION_MAPPINGに登録されていません。処理を中断します。`);
      logError(new Error(`関数 '${processorName}' がFUNCTION_MAPPINGに登録されていません`), 'scheduleChunkContinuation.UnregisteredFunction', 'HIGH', {
        processorName: processorName,
        startIndex: startIndex,
        itemsCount: items.length
      });
      return;
    }
    
    // 継続処理用のデータを保存
    const continuationData = {
      items: items.slice(startIndex),
      processorName: processorName,
      options: options,
      originalStartIndex: startIndex
    };
    
    PropertiesService.getScriptProperties().setProperty(
      `chunk_continuation_${trigger.getUniqueId()}`,
      JSON.stringify(continuationData)
    );
    
    Logger.log(`チャンク処理継続をスケジュール: トリガーID=${trigger.getUniqueId()}, 開始インデックス=${startIndex}, 関数名=${processorName}`);
    
  } catch (error) {
    logError(error, 'scheduleChunkContinuation', 'HIGH');
  }
}

/**
 * 関数名から関数参照を取得する安全なマッピング
 * @type {Object.<string, Function>}
 */
const FUNCTION_MAPPING = {
  // 従業員関連の処理関数
  'processEmployeeData': function(employee, index) {
    // 従業員データ処理のデフォルト実装
    return { employeeId: employee.employeeId, processed: true, index: index };
  },
  
  // 残業時間計算関数
  'calculateOvertime': function(employee, index) {
    // 残業時間計算のデフォルト実装
    return { employeeId: employee.employeeId, overtimeHours: 0, index: index };
  },
  
  // 月次サマリー計算関数
  'calculateMonthlySummary': function(employee, index) {
    // 月次サマリー計算のデフォルト実装
    return { employeeId: employee.employeeId, monthlyData: {}, index: index };
  },
  
  // 承認処理関数
  'processApproval': function(request, index) {
    // 承認処理のデフォルト実装
    return { requestId: request.requestId, approved: false, index: index };
  },
  
  // データ検証関数
  'validateData': function(item, index) {
    // データ検証のデフォルト実装
    return { item: item, valid: true, index: index };
  },
  
  // レポート生成関数
  'generateReport': function(data, index) {
    // レポート生成のデフォルト実装
    return { data: data, reportGenerated: true, index: index };
  }
};

/**
 * 関数名から関数参照を安全に取得
 * @param {string} functionName - 関数名
 * @return {Function} 関数参照
 * @throws {Error} 関数が見つからない場合
 */
function getFunctionByName(functionName) {
  if (!functionName) {
    throw new Error('関数名が指定されていません');
  }
  
  const func = FUNCTION_MAPPING[functionName];
  if (!func) {
    throw new Error(`関数 '${functionName}' が見つかりません。FUNCTION_MAPPINGに追加してください。`);
  }
  
  return func;
}

/**
 * FUNCTION_MAPPINGに新しい関数を追加
 * @param {string} functionName - 関数名
 * @param {Function} func - 関数参照
 */
function registerFunction(functionName, func) {
  if (!functionName || typeof func !== 'function') {
    throw new Error('関数名と関数参照の両方が必要です');
  }
  
  FUNCTION_MAPPING[functionName] = func;
  Logger.log(`関数 '${functionName}' をFUNCTION_MAPPINGに登録しました`);
}

/**
 * FUNCTION_MAPPINGから関数を削除
 * @param {string} functionName - 関数名
 */
function unregisterFunction(functionName) {
  if (FUNCTION_MAPPING[functionName]) {
    delete FUNCTION_MAPPING[functionName];
    Logger.log(`関数 '${functionName}' をFUNCTION_MAPPINGから削除しました`);
  }
}

/**
 * FUNCTION_MAPPINGの現在の状態を表示
 * @return {Object} 登録されている関数の一覧
 */
function showFunctionMapping() {
  const mapping = {};
  for (const [name, func] of Object.entries(FUNCTION_MAPPING)) {
    mapping[name] = typeof func;
  }
  
  Logger.log('FUNCTION_MAPPING の現在の状態:');
  Logger.log(JSON.stringify(mapping, null, 2));
  
  return mapping;
}

/**
 * チャンク処理の継続実行
 * @param {string} triggerId - トリガーID
 */
function continueChunkProcessing(triggerId) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const continuationDataJson = properties.getProperty(`chunk_continuation_${triggerId}`);
    
    if (!continuationDataJson) {
      Logger.log(`継続データが見つかりません: ${triggerId}`);
      return;
    }
    
    const continuationData = JSON.parse(continuationDataJson);
    
    // 関数名から関数を安全に取得
    let processor;
    try {
      processor = getFunctionByName(continuationData.processorName);
    } catch (error) {
      Logger.log(`関数の復元に失敗: ${error.message}`);
      logError(error, 'continueChunkProcessing.FunctionRestoration', 'HIGH', { 
        triggerId: triggerId,
        processorName: continuationData.processorName 
      });
      return;
    }
    
    // 継続処理を実行
    const result = processChunks(
      continuationData.items,
      processor,
      continuationData.options
    );
    
    // 継続データを削除
    properties.deleteProperty(`chunk_continuation_${triggerId}`);
    
    Logger.log(`チャンク処理継続完了: ${result.processedCount}件処理`);
    
  } catch (error) {
    logError(error, 'continueChunkProcessing', 'HIGH', { triggerId: triggerId });
  }
}

// ========== エラー復旧機能 ==========

/**
 * 自動復旧機能付きの関数実行
 * @param {Function} func - 実行する関数
 * @param {Object} options - 復旧オプション
 * @return {*} 関数の戻り値
 */
function executeWithRetry(func, options = {}) {
  const maxRetries = options.maxRetries || 3;
  const retryDelay = options.retryDelay || 1000;
  const context = options.context || 'RetryableFunction';
  
  let lastError;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      Logger.log(`${context} 実行試行 ${attempt}/${maxRetries}`);
      
      const result = func();
      
      if (attempt > 1) {
        Logger.log(`${context} 復旧成功 (試行回数: ${attempt})`);
      }
      
      return result;
      
    } catch (error) {
      lastError = error;
      
      Logger.log(`${context} 試行 ${attempt} 失敗: ${error.message}`);
      
      // 最後の試行でない場合は遅延後に再試行
      if (attempt < maxRetries) {
        Logger.log(`${retryDelay}ms後に再試行します`);
        Utilities.sleep(retryDelay);
      }
    }
  }
  
  // 全ての試行が失敗した場合
  logError(lastError, `${context}.AllRetriesFailed`, 'HIGH', {
    maxRetries: maxRetries,
    retryDelay: retryDelay
  });
  
  throw lastError;
}

// ========== システム監視機能 ==========

/**
 * システムヘルスチェック
 * @return {Object} ヘルスチェック結果
 */
function performHealthCheck() {
  return withErrorHandling(() => {
    const healthStatus = {
      timestamp: new Date(),
      overall: 'HEALTHY',
      checks: []
    };
    
    // 必須シートの存在確認
    const requiredSheets = [
      'Master_Employee', 'Master_Holiday', 'Log_Raw',
      'Daily_Summary', 'Monthly_Summary', 'Request_Responses'
    ];
    
    requiredSheets.forEach(sheetName => {
      try {
        getSheet(sheetName);
        healthStatus.checks.push({
          name: `Sheet.${sheetName}`,
          status: 'OK',
          message: 'シートアクセス正常'
        });
      } catch (error) {
        healthStatus.checks.push({
          name: `Sheet.${sheetName}`,
          status: 'ERROR',
          message: error.message
        });
        healthStatus.overall = 'UNHEALTHY';
      }
    });
    
    // 設定確認
    try {
      const adminEmail = getAdminEmail();
      healthStatus.checks.push({
        name: 'Config.AdminEmail',
        status: adminEmail ? 'OK' : 'WARNING',
        message: adminEmail || '管理者メール未設定'
      });
    } catch (error) {
      healthStatus.checks.push({
        name: 'Config.AdminEmail',
        status: 'ERROR',
        message: error.message
      });
    }
    
    // トリガー確認
    try {
      const triggers = ScriptApp.getProjectTriggers();
      healthStatus.checks.push({
        name: 'Triggers',
        status: triggers.length > 0 ? 'OK' : 'WARNING',
        message: `${triggers.length}個のトリガーが設定済み`
      });
    } catch (error) {
      healthStatus.checks.push({
        name: 'Triggers',
        status: 'ERROR',
        message: error.message
      });
    }
    
    return healthStatus;
    
  }, 'SystemHealthCheck', 'MEDIUM');
}

// ========== エラーログ管理機能 ==========

/**
 * 古いエラーログを削除
 */
function cleanupOldErrorLogs() {
  return withErrorHandling(() => {
    const logSheet = getOrCreateErrorLogSheet();
    const data = logSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('削除対象のエラーログがありません');
      return;
    }
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - ERROR_CONFIG.LOG_RETENTION_DAYS);
    
    const rowsToDelete = [];
    
    // ヘッダー行をスキップして確認
    for (let i = 1; i < data.length; i++) {
      const timestamp = new Date(data[i][0]);
      if (timestamp < cutoffDate) {
        rowsToDelete.push(i + 1); // 1ベースのインデックス
      }
    }
    
    // 古い行を削除（後ろから削除）
    rowsToDelete.reverse().forEach(rowIndex => {
      logSheet.deleteRow(rowIndex);
    });
    
    Logger.log(`古いエラーログを削除: ${rowsToDelete.length}件`);
    
  }, 'CleanupErrorLogs', 'LOW');
}

// ========== ユーティリティ関数 ==========

/**
 * 管理者メールアドレスを取得
 * @return {string} 管理者メールアドレス
 */
function getAdminEmail() {
  try {
    // Config.gsから取得を試行
    if (typeof getConfig === 'function') {
      const config = getConfig();
      if (config && config.adminEmail) {
        return config.adminEmail;
      }
    }
    
    // フォールバック: スクリプト所有者
    const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
    return owner ? owner.getEmail() : Session.getActiveUser().getEmail();
    
  } catch (error) {
    Logger.log(`管理者メール取得エラー: ${error.toString()}`);
    return Session.getActiveUser().getEmail();
  }
}

/**
 * エラーハンドリングシステムの初期化
 */
function initializeErrorHandling() {
  return withErrorHandling(() => {
    // エラーログシートを作成
    getOrCreateErrorLogSheet();
    
    // 古いログをクリーンアップ
    cleanupOldErrorLogs();
    
    // ヘルスチェックを実行
    const healthStatus = performHealthCheck();
    
    Logger.log('エラーハンドリングシステム初期化完了');
    Logger.log(`システム状態: ${healthStatus.overall}`);
    
    return {
      success: true,
      healthStatus: healthStatus
    };
    
  }, 'InitializeErrorHandling', 'HIGH');
}

// ========== テスト・デモ機能 ==========

/**
 * エラーハンドリングシステムのテスト
 */
function testErrorHandlingSystem() {
  try {
    Logger.log('=== エラーハンドリングシステムテスト開始 ===');
    
    // 1. 基本的なエラーログテスト
    Logger.log('1. エラーログテスト');
    try {
      throw new Error('テスト用エラー');
    } catch (error) {
      logError(error, 'TestContext', 'LOW', { testData: 'sample' });
      Logger.log('✓ エラーログ記録成功');
    }
    
    // 2. withErrorHandlingテスト
    Logger.log('2. withErrorHandlingテスト');
    const result = withErrorHandling(() => {
      return 'テスト成功';
    }, 'TestFunction', 'LOW');
    Logger.log(`✓ withErrorHandling結果: ${result}`);
    
    // 3. バッチ処理テスト
    Logger.log('3. バッチ処理テスト');
    const testItems = [1, 2, 3, 4, 5];
    const batchResult = processBatch(testItems, (item, index) => {
      return item * 2;
    }, {
      context: 'TestBatch',
      batchSize: 2
    });
    Logger.log(`✓ バッチ処理結果: 成功=${batchResult.processedCount}件`);
    
    // 4. チャンク処理テスト
    Logger.log('4. チャンク処理テスト');
    const chunkResult = processChunks(testItems, (item, index) => {
      return item * 3;
    }, {
      context: 'TestChunk',
      chunkSize: 2
    });
    Logger.log(`✓ チャンク処理結果: 成功=${chunkResult.processedCount}件`);
    
    // 5. 再試行機能テスト
    Logger.log('5. 再試行機能テスト');
    let attemptCount = 0;
    const retryResult = executeWithRetry(() => {
      attemptCount++;
      if (attemptCount < 2) {
        throw new Error('意図的な失敗');
      }
      return '再試行成功';
    }, {
      context: 'TestRetry',
      maxRetries: 3
    });
    Logger.log(`✓ 再試行結果: ${retryResult}`);
    
    // 6. ヘルスチェックテスト
    Logger.log('6. ヘルスチェックテスト');
    const healthStatus = performHealthCheck();
    Logger.log(`✓ ヘルスチェック結果: ${healthStatus.overall}`);
    
    Logger.log('=== エラーハンドリングシステムテスト完了 ===');
    
    return {
      success: true,
      message: 'すべてのテストが正常に完了しました'
    };
    
  } catch (error) {
    Logger.log(`テスト実行エラー: ${error.toString()}`);
    return {
      success: false,
      message: `テスト実行エラー: ${error.message}`
    };
  }
}

/**
 * エラーハンドリングシステムのデモ
 * 意図的にエラーを発生させてシステムの動作を確認
 */
function demoErrorHandlingSystem() {
  Logger.log('=== エラーハンドリングシステムデモ開始 ===');
  
  // 1. 低重要度エラーのデモ
  Logger.log('1. 低重要度エラーのデモ');
  withErrorHandling(() => {
    throw new Error('低重要度のテストエラー');
  }, 'DemoLowSeverity', 'LOW', {
    suppressError: true,
    defaultValue: 'デフォルト値'
  });
  
  // 2. 中重要度エラーのデモ
  Logger.log('2. 中重要度エラーのデモ');
  try {
    withErrorHandling(() => {
      throw new Error('中重要度のテストエラー');
    }, 'DemoMediumSeverity', 'MEDIUM');
  } catch (error) {
    Logger.log('中重要度エラーをキャッチしました');
  }
  
  // 3. バッチ処理でのエラーハンドリングデモ
  Logger.log('3. バッチ処理エラーハンドリングデモ');
  const testItems = [1, 2, 'error', 4, 5];
  const batchResult = processBatch(testItems, (item, index) => {
    if (typeof item !== 'number') {
      throw new Error(`無効なアイテム: ${item}`);
    }
    return item * 2;
  }, {
    context: 'DemoBatchError',
    batchSize: 2
  });
  
  Logger.log(`バッチ処理結果: 成功=${batchResult.processedCount}件, エラー=${batchResult.errors.length}件`);
  
  Logger.log('=== エラーハンドリングシステムデモ完了 ===');
  
  return {
    success: true,
    message: 'デモが正常に完了しました',
    batchResult: batchResult
  };
}

/**
 * エラーログの表示
 * 最新のエラーログを表示
 * @param {number} limit - 表示件数制限
 */
function showRecentErrorLogs(limit = 10) {
  try {
    const logSheet = getOrCreateErrorLogSheet();
    const data = logSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('エラーログがありません');
      return [];
    }
    
    // ヘッダー行を除いて最新のログを取得
    const logs = data.slice(1).reverse().slice(0, limit);
    
    Logger.log(`=== 最新のエラーログ (${logs.length}件) ===`);
    
    logs.forEach((log, index) => {
      const timestamp = Utilities.formatDate(new Date(log[0]), 'JST', 'yyyy/MM/dd HH:mm:ss');
      Logger.log(`${index + 1}. [${log[2]}] ${log[1]}: ${log[4]} (${timestamp})`);
    });
    
    return logs;
    
  } catch (error) {
    Logger.log(`エラーログ表示エラー: ${error.toString()}`);
    return [];
  }
}

/**
 * エラー統計の表示
 */
function showErrorStatistics() {
  try {
    const logSheet = getOrCreateErrorLogSheet();
    const data = logSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('エラーログがありません');
      return {};
    }
    
    const logs = data.slice(1); // ヘッダー行を除く
    const stats = {
      total: logs.length,
      bySeverity: {},
      byContext: {},
      recent24h: 0
    };
    
    const now = new Date();
    const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    
    logs.forEach(log => {
      const severity = log[2];
      const context = log[1];
      const timestamp = new Date(log[0]);
      
      // 重要度別統計
      stats.bySeverity[severity] = (stats.bySeverity[severity] || 0) + 1;
      
      // コンテキスト別統計
      stats.byContext[context] = (stats.byContext[context] || 0) + 1;
      
      // 24時間以内のエラー
      if (timestamp > yesterday) {
        stats.recent24h++;
      }
    });
    
    Logger.log('=== エラー統計 ===');
    Logger.log(`総エラー数: ${stats.total}件`);
    Logger.log(`24時間以内: ${stats.recent24h}件`);
    Logger.log('重要度別:');
    Object.keys(stats.bySeverity).forEach(severity => {
      Logger.log(`  ${severity}: ${stats.bySeverity[severity]}件`);
    });
    Logger.log('コンテキスト別:');
    Object.keys(stats.byContext).forEach(context => {
      Logger.log(`  ${context}: ${stats.byContext[context]}件`);
    });
    
    return stats;
    
  } catch (error) {
    Logger.log(`エラー統計表示エラー: ${error.toString()}`);
    return {};
  }
}

/**
 * エラーハンドリングシステムの設定表示
 */
function showErrorHandlingConfig() {
  Logger.log('=== エラーハンドリングシステム設定 ===');
  Logger.log(`最大実行時間: ${ERROR_CONFIG.MAX_EXECUTION_TIME}ms`);
  Logger.log(`デフォルトチャンクサイズ: ${ERROR_CONFIG.DEFAULT_CHUNK_SIZE}`);
  Logger.log(`デフォルトバッチサイズ: ${ERROR_CONFIG.DEFAULT_BATCH_SIZE}`);
  Logger.log(`エラー通知間隔: ${ERROR_CONFIG.ERROR_NOTIFICATION_INTERVAL}ms`);
  Logger.log(`ログ保持期間: ${ERROR_CONFIG.LOG_RETENTION_DAYS}日`);
  Logger.log('重要度レベル:', Object.keys(ERROR_CONFIG.SEVERITY_LEVELS).join(', '));
  Logger.log('================================');
}

/**
 * 構文チェック用テスト関数
 */
function syntaxTest() {
  try {
    Logger.log('ErrorHandler.gs構文チェック: OK');
    return true;
  } catch (error) {
    Logger.log('ErrorHandler.gs構文エラー: ' + error.toString());
    return false;
  }
}

/**
 * eval()置き換え機能のテスト
 */
function testEvalReplacement() {
  try {
    Logger.log('=== eval()置き換え機能のテスト開始 ===');
    
    // 1. 関数マッピングの表示
    Logger.log('1. FUNCTION_MAPPINGの状態確認:');
    showFunctionMapping();
    
    // 2. 正常な関数取得のテスト
    Logger.log('2. 正常な関数取得のテスト:');
    const processEmployeeData = getFunctionByName('processEmployeeData');
    Logger.log(`取得した関数の型: ${typeof processEmployeeData}`);
    
    // 3. 関数実行のテスト
    Logger.log('3. 関数実行のテスト:');
    const testEmployee = { employeeId: 'TEST001', name: 'テスト太郎' };
    const result = processEmployeeData(testEmployee, 0);
    Logger.log(`実行結果: ${JSON.stringify(result)}`);
    
    // 4. 存在しない関数のテスト
    Logger.log('4. 存在しない関数のテスト:');
    try {
      getFunctionByName('nonExistentFunction');
      Logger.log('エラー: 存在しない関数が取得できてしまいました');
      return false;
    } catch (error) {
      Logger.log(`期待されるエラー: ${error.message}`);
    }
    
    // 5. 空の関数名のテスト
    Logger.log('5. 空の関数名のテスト:');
    try {
      getFunctionByName('');
      Logger.log('エラー: 空の関数名が処理されてしまいました');
      return false;
    } catch (error) {
      Logger.log(`期待されるエラー: ${error.message}`);
    }
    
    // 6. 関数登録のテスト
    Logger.log('6. 関数登録のテスト:');
    const testFunction = function(item, index) {
      return { custom: true, item: item, index: index };
    };
    registerFunction('testCustomFunction', testFunction);
    
    const registeredFunction = getFunctionByName('testCustomFunction');
    const customResult = registeredFunction({ id: 'CUSTOM001' }, 1);
    Logger.log(`カスタム関数実行結果: ${JSON.stringify(customResult)}`);
    
    // 7. 関数削除のテスト
    Logger.log('7. 関数削除のテスト:');
    unregisterFunction('testCustomFunction');
    try {
      getFunctionByName('testCustomFunction');
      Logger.log('エラー: 削除された関数が取得できてしまいました');
      return false;
    } catch (error) {
      Logger.log(`期待されるエラー: ${error.message}`);
    }
    
    Logger.log('=== eval()置き換え機能のテスト完了 ===');
    return true;
    
  } catch (error) {
    Logger.log(`テストエラー: ${error.message}`);
    logError(error, 'testEvalReplacement', 'HIGH');
    return false;
  }
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

/**
 * エラーログシートURL構築のテスト
 */
function testErrorLogUrlConstruction() {
  try {
    Logger.log('=== エラーログシートURL構築のテスト開始 ===');
    
    // 1. getSheetId関数のテスト
    Logger.log('1. getSheetId関数のテスト:');
    const errorLogSheetId = getSheetId('Error_Log');
    Logger.log(`Error_LogシートID: ${errorLogSheetId}`);
    
    // 2. URL構築のテスト
    Logger.log('2. URL構築のテスト:');
    const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    const errorLogUrl = `${spreadsheetUrl}#gid=${errorLogSheetId}`;
    Logger.log(`構築されたURL: ${errorLogUrl}`);
    
    // 3. 他のシートでもテスト
    Logger.log('3. 他のシートでのテスト:');
    const requestResponsesSheetId = getSheetId('Request_Responses');
    const dailySummarySheetId = getSheetId('Daily_Summary');
    Logger.log(`Request_ResponsesシートID: ${requestResponsesSheetId}`);
    Logger.log(`Daily_SummaryシートID: ${dailySummarySheetId}`);
    
    // 4. 存在しないシートのテスト
    Logger.log('4. 存在しないシートのテスト:');
    const nonExistentSheetId = getSheetId('NonExistentSheet');
    Logger.log(`存在しないシートID: ${nonExistentSheetId}`);
    
    Logger.log('=== エラーログシートURL構築のテスト完了 ===');
    return true;
    
  } catch (error) {
    Logger.log(`テストエラー: ${error.message}`);
    logError(error, 'testErrorLogUrlConstruction', 'HIGH');
    return false;
  }
}
