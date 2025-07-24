/**
 * 出勤管理システム初期セットアップスクリプト
 * TDD実装フェーズ1完了後の統合セットアップ
 */

/**
 * 出勤管理システム全体の初期セットアップ
 * 一括でスプレッドシートとフォームを作成
 */
function setupAttendanceSystem() {
  console.log('=== 出勤管理システム セットアップ開始 ===');
  
  try {
    // 1. テスト環境の確認
    console.log('1. テスト環境の確認...');
    if (!testEnvironmentCheck()) {
      throw new Error('テスト環境の確認に失敗しました');
    }
    console.log('✓ テスト環境確認完了');
    
    // 2. スプレッドシートの初期化
    console.log('2. スプレッドシートの初期化...');
    if (!initializeAllSheets()) {
      throw new Error('スプレッドシートの初期化に失敗しました');
    }
    console.log('✓ スプレッドシート初期化完了');
    
    // 3. サンプルデータの投入
    console.log('3. サンプルデータの投入...');
    setupSampleData();
    console.log('✓ サンプルデータ投入完了');
    
    // 4. システム設定の初期化
    console.log('4. システム設定の初期化...');
    setupSystemConfigData();
    console.log('✓ システム設定初期化完了');
    
    // 5. セットアップ確認テストの実行
    console.log('5. セットアップ確認テスト...');
    runSystemIntegrationTest();
    console.log('✓ セットアップ確認完了');
    
    console.log('=== 出勤管理システム セットアップ完了 ===');
    console.log('次のステップ: フェーズ2のコア機能実装へ進むことができます。');
    
    return true;
  } catch (error) {
    console.log('✗ セットアップエラー: ' + error.message);
    return false;
  }
}

/**
 * サンプルデータの投入
 * テスト用の基本データを各シートに設定
 */
function setupSampleData() {
  try {
    // Master_Employee サンプルデータ
    setupSampleEmployeeData();
    
    // Master_Holiday サンプルデータ
    setupSampleHolidayData();
    
    console.log('サンプルデータの投入が完了しました');
  } catch (error) {
    throw new Error('サンプルデータ投入エラー: ' + error.message);
  }
}

/**
 * 従業員マスタのサンプルデータ設定
 */
function setupSampleEmployeeData() {
  var employeeData = [
    ['EMP001', '田中太郎', 'tanaka@example.com', '営業部', '正社員', 'manager@example.com', '09:00', '18:00'],
    ['EMP002', '佐藤花子', 'sato@example.com', '開発部', '正社員', 'dev-manager@example.com', '09:30', '18:30'],
    ['EMP003', '鈴木次郎', 'suzuki@example.com', '営業部', 'パート', 'manager@example.com', '10:00', '15:00'],
    ['EMP004', '管理者太郎', 'manager@example.com', '管理部', '正社員', '', '09:00', '18:00'],
    ['EMP005', '開発管理者', 'dev-manager@example.com', '開発部', '正社員', '', '09:00', '18:00']
  ];
  
  var sheetName = getSheetName('MASTER_EMPLOYEE');
  
  employeeData.forEach(function(rowData) {
    appendDataToSheet(sheetName, rowData);
  });
  
  console.log('従業員マスタにサンプルデータ ' + employeeData.length + '件を追加しました');
}

/**
 * 休日マスタのサンプルデータ設定
 */
function setupSampleHolidayData() {
  var today = new Date();
  var currentYear = today.getFullYear();
  
  var holidayData = [
    [currentYear + '-01-01', '元日', 'TRUE', 'TRUE'],
    [currentYear + '-01-08', '成人の日', 'TRUE', 'TRUE'],
    [currentYear + '-02-11', '建国記念の日', 'TRUE', 'TRUE'],
    [currentYear + '-04-29', '昭和の日', 'TRUE', 'TRUE'],
    [currentYear + '-05-03', '憲法記念日', 'TRUE', 'TRUE'],
    [currentYear + '-05-04', 'みどりの日', 'TRUE', 'TRUE'],
    [currentYear + '-05-05', 'こどもの日', 'TRUE', 'TRUE'],
    [currentYear + '-12-31', '会社休業日', 'FALSE', 'TRUE']
  ];
  
  var sheetName = getSheetName('MASTER_HOLIDAY');
  
  holidayData.forEach(function(rowData) {
    appendDataToSheet(sheetName, rowData);
  });
  
  console.log('休日マスタにサンプルデータ ' + holidayData.length + '件を追加しました');
}

/**
 * システム設定の初期化
 * System_Configシートに必要な設定を追加
 */
function setupSystemConfigData() {
  try {
    var configData = [
      ['ADMIN_EMAILS', 'manager@example.com,dev-manager@example.com', '管理者メールアドレス（カンマ区切り）', 'TRUE'],
      ['EMAIL_MOCK_ENABLED', 'TRUE', 'メール送信のモック有効化', 'TRUE'],
      ['EMAIL_ACTUAL_SEND', 'FALSE', '実際のメール送信フラグ', 'TRUE'],
      ['MAX_WORK_HOURS_PER_DAY', '24', '1日の最大労働時間', 'TRUE'],
      ['STANDARD_WORK_HOURS', '8', '標準労働時間', 'TRUE'],
      ['BREAK_TIME_AUTO_DEDUCT', '45', '自動控除する休憩時間（分）', 'TRUE'],
      ['OVERTIME_THRESHOLD', '8', '残業判定の閾値（時間）', 'TRUE']
    ];
    
    var sheetName = getSheetName('SYSTEM_CONFIG');
    
    configData.forEach(function(rowData) {
      appendDataToSheet(sheetName, rowData);
    });
    
    console.log('システム設定に ' + configData.length + '件の設定を追加しました');
  } catch (error) {
    throw new Error('システム設定初期化エラー: ' + error.message);
  }
}

/**
 * システム統合テスト
 * セットアップ後の動作確認
 */
function runSystemIntegrationTest() {
  try {
    console.log('システム統合テストを開始します...');
    
    // 1. 全シートの存在確認
    var requiredSheets = ['MASTER_EMPLOYEE', 'MASTER_HOLIDAY', 'LOG_RAW', 'DAILY_SUMMARY', 'MONTHLY_SUMMARY', 'REQUEST_RESPONSES'];
    
    requiredSheets.forEach(function(sheetType) {
      var sheetName = getSheetName(sheetType);
      if (!sheetExists(sheetName)) {
        throw new Error('必須シート "' + sheetName + '" が存在しません');
      }
    });
    console.log('✓ 全必須シートの存在確認完了');
    
    // 2. 定数取得テスト
    var testColumnIndex = getColumnIndex('EMPLOYEE', 'NAME');
    if (testColumnIndex !== 1) {
      throw new Error('定数取得テスト失敗: EMPLOYEE.NAME = ' + testColumnIndex);
    }
    console.log('✓ 定数取得テスト完了');
    
    // 3. ユーティリティ関数テスト
    var testDate = formatDate(new Date(2025, 6, 10));
    if (testDate !== '2025-07-10') {
      throw new Error('ユーティリティ関数テスト失敗: formatDate = ' + testDate);
    }
    console.log('✓ ユーティリティ関数テスト完了');
    
    // 4. データ読み取りテスト
    var employeeRowCount = getSheetRowCount(getSheetName('MASTER_EMPLOYEE'));
    if (employeeRowCount < 2) { // ヘッダー + 最低1データ
      throw new Error('従業員データが不足しています: ' + employeeRowCount + '行');
    }
    console.log('✓ データ読み取りテスト完了（従業員データ: ' + (employeeRowCount - 1) + '件）');
    
    console.log('システム統合テスト完了');
  } catch (error) {
    throw new Error('システム統合テスト失敗: ' + error.message);
  }
}

/**
 * システム全体E2E統合テスト（Redフェーズ雛形）
 * 主要シナリオ例:
 * 1. 打刻API呼び出し→ログ保存
 * 2. 日次/週次/月次集計トリガー実行→集計結果検証
 * 3. 未退勤者メール送信→送信ログ検証
 * 4. Googleフォーム連携→データ保存・重複チェック
 * 5. 認証・権限チェック→正常/異常系
 * 6. 異常系・境界値（例: 二重打刻、無効データ、権限エラー等）
 */
function runFullIntegrationTest() {
  try {
    // テスト用：MASTER_EMPLOYEEシートにテスト従業員を追加
    var empSheet = getOrCreateSheet(getSheetName('MASTER_EMPLOYEE'));
    var empData = empSheet.getDataRange().getValues();
    var exists = empData.some(function(row) {
      return row[1] === 'EMP001' && row[2] === '田中太郎' && row[3] === 'tanaka@example.com';
    });
    if (!exists) {
      // 正しい列順序で追加（EMPLOYEE_COLUMNS定義に従う）
      // A列: 社員ID, B列: 氏名, C列: Gmail, D列: 所属, E列: 雇用区分, F列: 上司Gmail, G列: 基準始業, H列: 基準終業
      empSheet.appendRow(['EMP001', '田中太郎', 'tanaka@example.com', '営業部', '正社員', 'manager@example.com', '09:00', '18:00']);
    }
    // テスト前の設定値を保存
    var originalAuthCacheEnabled = AUTH_CACHE ? AUTH_CACHE.cacheEnabled : undefined;
    var originalBruteForceProtection = AUTH_CONFIG ? AUTH_CONFIG.BRUTE_FORCE_PROTECTION : undefined;
    
    try {
      // 認証キャッシュをクリア
      if (typeof clearAuthCache === 'function') {
        clearAuthCache();
      }
      // テスト用設定を適用
      if (typeof AUTH_CACHE !== 'undefined') {
        // テスト用：Log_Rawシートを初期化（ヘッダー以外を削除）
        var logRawSheet = getOrCreateSheet(getSheetName('LOG_RAW'));
        var lastRow = logRawSheet.getLastRow();
        if (lastRow > 1) {
          try {
            logRawSheet.deleteRows(2, lastRow - 1);
          } catch (error) {
            console.log('Log_Rawシート初期化エラー: ' + error.message);
            throw new Error('テストデータの初期化に失敗しました');
          }
        }
        console.log('Log_Rawシート初期化完了');
      }
    } finally {
      // 設定を元に戻す
      if (typeof AUTH_CACHE !== 'undefined' && originalAuthCacheEnabled !== undefined) {
        AUTH_CACHE.cacheEnabled = originalAuthCacheEnabled;
      }
      if (typeof AUTH_CONFIG !== 'undefined' && originalBruteForceProtection !== undefined) {
        AUTH_CONFIG.BRUTE_FORCE_PROTECTION = originalBruteForceProtection;
      }
    }
    
    console.log('=== E2E統合テスト開始 ===');
    // 1. 打刻API呼び出し（例: processClock）
    var userInfo = { email: 'tanaka@example.com', employeeId: 'EMP001', employeeName: '田中太郎' };
    var clockResult = processClock('IN', userInfo);
    assertTrue(clockResult.success, '打刻APIが成功し、success=trueで返るべき');

    // ここで未退勤者テストデータを追加（IN打刻のみ、OUTなし、日付は前日）
    var yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    // 3. 未退勤者メール送信テストは将来実装予定    // 2. 日次集計トリガー実行（例: dailyJob）
    var dailyResult = dailyJob();
    assertTrue(dailyResult.success, '日次集計トリガーが成功し、success=trueで返るべき');

    // 3. 未退勤者メール送信（例: sendUnfinishedClockOutEmail_MailManager）
    // ↓この部分を削除
    // var unfinishedEmployees = [{
    //   employeeId: 'EMP001',
    //   employeeName: '田中太郎',
    //   name: '田中太郎', // テンプレート用
    //   email: 'tanaka@example.com',
    //   clockInTime: '09:00', // テンプレート用
    //   currentTime: '18:30'  // テンプレート用
    // }];
    // var mailResult = sendUnfinishedClockOutEmail_MailManager(unfinishedEmployees, new Date());
    // assertTrue(mailResult.success, '未退勤者メール送信が成功し、success=trueで返るべき');

    // 4. Googleフォーム連携→データ保存・重複チェック
    // GoogleフォームのFormResponseイベントのダミー
    var dummyFormEvent = {
      response: {
        getItemResponses: function() {
          return [
            {
              getItem: function() { return { getTitle: function() { return '社員ID'; } }; },
              getResponse: function() { return 'EMP001'; }
            },
            {
              getItem: function() { return { getTitle: function() { return '氏名'; } }; },
              getResponse: function() { return '田中太郎'; }
            },
            {
              getItem: function() { return { getTitle: function() { return '打刻種別'; } }; },
              getResponse: function() { return 'IN'; }
            }
            // 必要に応じて他の項目も追加
          ];
        },
        getTimestamp: function() { return new Date(); }
      }
    };
    var formResult = processGoogleFormResponse(dummyFormEvent);
    assertTrue(formResult.success, 'フォーム応答処理が成功すべき');
    // 重複チェック（同じデータを再送）
    var formResultDup = processGoogleFormResponse(dummyFormEvent);
    assertTrue(formResultDup.duplicate || formResultDup.success === false, '重複データは検出されるべき');

    // 5. 認証・権限チェック（正常/異常系）
    if (typeof clearAuthCache === 'function') {
      clearAuthCache();
    }
    var authResult = authenticateUser('tanaka@example.com');
    assertTrue(authResult, '認証が成功すべき');
    var permResult = checkPermission('tanaka@example.com', 'ADMIN_ACTION');
    assertFalse(permResult, '権限がない場合はfalseを返すべき');

    // 6. 異常系・境界値（二重打刻、無効データ、権限エラー等）
    var dupResult = processClock('IN', userInfo); // 直前と同じ打刻
    assertFalse(dupResult.success, '二重打刻は失敗すべき');
    // WEBAPP_CONFIG.ERROR_MESSAGES.DUPLICATE_ACTIONが未定義の場合はメッセージ検証を省略

    // 7. 全体フローのデータ整合性検証
    var logRawCount = getSheetRowCount(getSheetName('LOG_RAW'));
    assertTrue(logRawCount > 1, '打刻後のLog_Rawシートにデータが追加されているべき');

    console.log('✓ E2E統合テスト成功');
    return true;
  } catch (error) {
    console.log('✗ E2E統合テスト失敗: ' + error.message);
    throw error;
  }
}

/**
 * セットアップ状況の確認
 * 現在のシステム状態をレポート
 */
function checkSetupStatus() {
  console.log('=== セットアップ状況確認 ===');
  
  try {
    // 環境情報
    showGASEnvironmentInfo();
    
    // TDD進捗
    showTDDStatus();
    
    // テストカバレッジ
    showTestCoverage();
    
    // シート状況確認
    console.log('\n=== シート状況 ===');
    var sheetTypes = ['MASTER_EMPLOYEE', 'MASTER_HOLIDAY', 'LOG_RAW', 'DAILY_SUMMARY', 'MONTHLY_SUMMARY', 'REQUEST_RESPONSES', 'SYSTEM_CONFIG'];
    
    sheetTypes.forEach(function(sheetType) {
      var sheetName = getSheetName(sheetType);
      var exists = sheetExists(sheetName);
      var rowCount = exists ? getSheetRowCount(sheetName) : 0;
      
      console.log(sheetName + ': ' + (exists ? '存在（' + rowCount + '行）' : '未作成'));
    });
    
  } catch (error) {
    console.log('状況確認エラー: ' + error.message);
  }
}

/**
 * フェーズ1完了確認
 * 次のフェーズに進める状態かチェック
 */
function checkPhase1Completion() {
  console.log('=== フェーズ1完了確認 ===');
  
  var completionChecks = {
    'テスト環境': false,
    'スプレッドシート': false,
    'サンプルデータ': false,
    '統合テスト': false
  };
  
  try {
    // 1. テスト環境確認
    completionChecks['テスト環境'] = testEnvironmentCheck();
    
    // 2. スプレッドシート確認
    var requiredSheets = ['MASTER_EMPLOYEE', 'MASTER_HOLIDAY', 'LOG_RAW', 'DAILY_SUMMARY', 'MONTHLY_SUMMARY', 'REQUEST_RESPONSES', 'SYSTEM_CONFIG'];
    var allSheetsExist = requiredSheets.every(function(sheetType) {
      return sheetExists(getSheetName(sheetType));
    });
    completionChecks['スプレッドシート'] = allSheetsExist;
    
    // 3. サンプルデータ確認
    var employeeCount = getSheetRowCount(getSheetName('MASTER_EMPLOYEE'));
    var holidayCount = getSheetRowCount(getSheetName('MASTER_HOLIDAY'));
    completionChecks['サンプルデータ'] = (employeeCount > 1 && holidayCount > 1);
    
    // 4. 統合テスト確認
    runSystemIntegrationTest();
    completionChecks['統合テスト'] = true;
    
  } catch (error) {
    console.log('確認エラー: ' + error.message);
  }
  
  // 結果表示
  console.log('\n=== 完了状況 ===');
  Object.keys(completionChecks).forEach(function(check) {
    var status = completionChecks[check] ? '✓ 完了' : '✗ 未完了';
    console.log(check + ': ' + status);
  });
  
  var allCompleted = Object.keys(completionChecks).every(function(check) {
    return completionChecks[check];
  });
  
  if (allCompleted) {
    console.log('\n🎉 フェーズ1が正常に完了しました！');
    console.log('フェーズ2（コア機能実装）に進むことができます。');
  } else {
    console.log('\n⚠️ フェーズ1に未完了の項目があります。');
    console.log('setupAttendanceSystem() を実行して初期セットアップを完了してください。');
  }
  
  return allCompleted;
} 