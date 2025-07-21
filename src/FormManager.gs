/**
 * Google フォーム連携管理
 * TDD実装: Red-Green-Refactorサイクル
 * 
 * 機能:
 * - フォーム応答の受信処理
 * - フォームデータの変換処理
 * - スプレッドシートへの保存処理
 * - エラーハンドリング
 */

/**
 * フォーム応答を処理する
 * @param {Object} formResponse - フォーム応答データ
 * @return {Object} 処理結果
 */
function processFormResponse(formResponse) {
  try {
    // バリデーション
    if (!validateFormData(formResponse)) {
      return {
        success: false,
        message: 'フォームデータのバリデーションエラー',
        data: null
      };
    }
    
    // 重複チェック
    if (checkDuplicateFormResponse(formResponse)) {
      return {
        success: false,
        message: '重複するフォーム応答です',
        data: null
      };
    }
    
    // データ変換
    var logData = convertFormDataToLogRaw(formResponse);
    
    // スプレッドシートに保存
    var saveResult = saveFormDataToSheet(logData);
    
    if (!saveResult.success) {
      return {
        success: false,
        message: 'スプレッドシートへの保存に失敗しました',
        data: null
      };
    }
    
    return {
      success: true,
      message: 'フォーム応答が正常に処理されました',
      data: {
        employeeId: formResponse.employeeId,
        employeeName: formResponse.employeeName,
        action: formResponse.action,
        timestamp: formResponse.timestamp,
        rowNumber: saveResult.rowNumber
      }
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'フォーム応答処理でエラーが発生しました: ' + error.message,
      data: null
    };
  }
}

/**
 * フォームデータをLog_Raw形式に変換する
 * @param {Object} formData - フォームデータ
 * @return {Array} Log_Raw形式のデータ配列
 */
function convertFormDataToLogRaw(formData) {
  return [
    formData.timestamp,      // タイムスタンプ
    formData.employeeId,     // 社員ID
    formData.employeeName,   // 氏名
    formData.action,         // アクション
    formData.ipAddress,      // 端末IP
    formData.remarks         // 備考
  ];
}

/**
 * フォームデータのバリデーション
 * @param {Object} formData - フォームデータ
 * @return {boolean} バリデーション結果
 */
function validateFormData(formData) {
  // 必須フィールドのチェック
  if (!formData.timestamp || !formData.employeeId || !formData.employeeName || 
      !formData.action || !formData.ipAddress) {
    return false;
  }
  
  // アクションの有効性チェック
  var validActions = ['IN', 'OUT', 'BRK_IN', 'BRK_OUT'];
  if (validActions.indexOf(formData.action) === -1) {
    return false;
  }
  
  // 社員IDの形式チェック
  if (typeof formData.employeeId !== 'string' || formData.employeeId.trim() === '') {
    return false;
  }
  
  // 氏名の形式チェック
  if (typeof formData.employeeName !== 'string' || formData.employeeName.trim() === '') {
    return false;
  }
  
  return true;
}

/**
 * フォームデータをスプレッドシートに保存する
 * @param {Array} logData - Log_Raw形式のデータ
 * @return {Object} 保存結果
 */
function saveFormDataToSheet(logData) {
  try {
    var sheet = getSheet(getSheetName('LOG_RAW'));
    var lastRow = sheet.getLastRow();
    var targetRow = lastRow + 1;
    
    // データを書き込み
    var range = sheet.getRange(targetRow, 1, 1, logData.length);
    range.setValues([logData]);
    
    return {
      success: true,
      message: 'データが正常に保存されました',
      rowNumber: targetRow
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'スプレッドシートへの保存に失敗しました: ' + error.message,
      rowNumber: null
    };
  }
}

/**
 * フォーム応答の重複チェック
 * @param {Object} formData - フォームデータ
 * @return {boolean} 重複の場合true
 */
function checkDuplicateFormResponse(formData) {
  try {
    var sheet = getSheet(getSheetName('LOG_RAW'));
    var data = sheet.getDataRange().getValues();
    
    // ヘッダー行をスキップ
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // 同じ社員ID、アクション、タイムスタンプ（1分以内）の組み合わせをチェック
      // テスト環境ではより厳密な重複チェックを行う
      if (row[1] === formData.employeeId && 
          row[3] === formData.action &&
          Math.abs(new Date(row[0]) - formData.timestamp) < 1 * 60 * 1000) {
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    // エラーの場合は重複とみなして安全側に倒す
    console.log('重複チェックでエラーが発生しました: ' + error.message);
    return true;
  }
}

/**
 * 複数のフォーム応答を一括処理する
 * @param {Array} formResponses - フォーム応答データの配列
 * @return {Array} 処理結果の配列
 */
function processMultipleFormResponses(formResponses) {
  var results = [];
  var batchSize = 10; // バッチサイズ制限
  
  console.log('一括処理開始: ' + formResponses.length + '件の応答');
  
  for (var i = 0; i < formResponses.length; i++) {
    try {
      var result = processFormResponse(formResponses[i]);
      results.push(result);
      
      // バッチサイズ制限のチェック
      if ((i + 1) % batchSize === 0) {
        console.log('バッチ処理進捗: ' + (i + 1) + '/' + formResponses.length);
      }
      
    } catch (error) {
      console.log('一括処理でエラーが発生しました (index: ' + i + '): ' + error.message);
      results.push({
        success: false,
        message: '一括処理エラー: ' + error.message,
        data: null,
        index: i
      });
    }
  }
  
  console.log('一括処理完了: ' + results.length + '件処理');
  return results;
}

/**
 * フォーム応答の統計情報を取得する
 * @return {Object} 統計情報
 */
function getFormResponseStats() {
  try {
    var sheet = getSheet(getSheetName('LOG_RAW'));
    var data = sheet.getDataRange().getValues();
    
    var totalResponses = data.length - 1; // ヘッダー行を除く
    var successCount = 0;
    var errorCount = 0;
    
    // 実際の統計計算（エラーレコードの検出）
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // エラーレコードの判定（備考欄にエラー情報がある場合）
      if (row[5] && row[5].toString().toLowerCase().includes('error')) {
        errorCount++;
      } else {
        successCount++;
      }
    }
    
    return {
      totalResponses: totalResponses,
      successCount: successCount,
      errorCount: errorCount,
      successRate: totalResponses > 0 ? (successCount / totalResponses) * 100 : 0,
      lastUpdated: new Date().toISOString()
    };
    
  } catch (error) {
    console.log('統計情報取得でエラーが発生しました: ' + error.message);
    return {
      totalResponses: 0,
      successCount: 0,
      errorCount: 0,
      successRate: 0,
      lastUpdated: new Date().toISOString(),
      error: error.message
    };
  }
}

/**
 * Google フォーム応答を処理する（実際のフォーム応答用）
 * @param {Object} e - フォーム応答イベント
 * @return {Object} 処理結果
 */
function processGoogleFormResponse(e) {
  try {
    console.log('Google フォーム応答を受信しました');
    
    // フォーム応答データの取得
    var formResponse = e.response;
    var itemResponses = formResponse.getItemResponses();
    
    // フォームデータの構築
    var formData = {
      timestamp: formResponse.getTimestamp(),
      employeeId: '',
      employeeName: '',
      action: '',
      ipAddress: '',
      remarks: ''
    };
    
    // フォーム項目からデータを抽出
    for (var i = 0; i < itemResponses.length; i++) {
      var itemResponse = itemResponses[i];
      var question = itemResponse.getItem().getTitle();
      var answer = itemResponse.getResponse();
      
      // 質問タイトルに基づいてデータを設定
      if (question.includes('社員ID') || question.includes('ID')) {
        formData.employeeId = answer;
      } else if (question.includes('氏名') || question.includes('名前')) {
        formData.employeeName = answer;
      } else if (question.includes('アクション') || question.includes('打刻')) {
        formData.action = answer;
      } else if (question.includes('備考') || question.includes('コメント')) {
        formData.remarks = answer;
      }
    }
    
    // IPアドレスの設定（フォームから取得できない場合はデフォルト値）
    formData.ipAddress = '192.168.1.100'; // 実際の環境では適切に取得
    
    console.log('抽出されたフォームデータ:', formData);
    
    // フォーム応答を処理
    var result = processFormResponse(formData);
    
    // 処理結果をログに記録
    console.log('フォーム応答処理結果:', result);
    
    return result;
    
  } catch (error) {
    console.log('Google フォーム応答処理でエラーが発生しました: ' + error.message);
    return {
      success: false,
      message: 'Google フォーム応答処理でエラーが発生しました: ' + error.message,
      data: null
    };
  }
}

/**
 * フォーム応答データをクリーンアップする
 * @param {Object} formData - フォームデータ
 * @return {Object} クリーンアップされたデータ
 */
function cleanupFormData(formData) {
  var cleanedData = {};
  
  // 各フィールドをクリーンアップ
  for (var key in formData) {
    if (formData.hasOwnProperty(key)) {
      var value = formData[key];
      
      // 文字列の場合、前後の空白を削除
      if (typeof value === 'string') {
        cleanedData[key] = value.trim();
      } else {
        cleanedData[key] = value;
      }
    }
  }
  
  return cleanedData;
} 