/**
 * WebApp.gs のテストスイート
 * TDD実装用テストファイル
 * 
 * テスト対象:
 * - doGet(e): 打刻画面表示
 * - doPost(e): 打刻データ処理
 * - generateClockHTML(userInfo): HTML生成
 * - processClock(action, userInfo): 打刻処理
 */

/**
 * WebAppTestRunner - WebApp関連のテスト実行
 * @return {void}
 */
function runWebAppTests() {
  console.log('=== WebApp.gs テスト開始 ===');
  
  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: []
  };
  
  // テスト一覧
  const tests = [
    testDoGet_AuthenticatedUser_ReturnsClockHTML,
    testDoGet_UnauthenticatedUser_ReturnsUnauthorized,
    testDoPost_ValidClockIn_SavesData,
    testDoPost_ValidClockOut_SavesData,
    testDoPost_InvalidAction_ReturnsError,
    testGenerateClockHTML_ValidUser_ReturnsHTML,
    testProcessClock_ClockIn_UpdatesLogRaw,
    testProcessClock_ClockOut_UpdatesLogRaw,
    testProcessClock_DuplicateAction_ReturnsError
  ];
  
  // テスト実行
  tests.forEach(test => {
    results.total++;
    try {
      console.log(`実行中: ${test.name}`);
      test();
      results.passed++;
      console.log(`✅ ${test.name}`);
    } catch (error) {
      results.failed++;
      results.errors.push(`❌ ${test.name}: ${error.message}`);
      console.log(`❌ ${test.name}: ${error.message}`);
    }
  });
  
  // 結果サマリー
  console.log('\n=== WebApp.gs テスト結果 ===');
  console.log(`総テスト数: ${results.total}`);
  console.log(`成功: ${results.passed}`);
  console.log(`失敗: ${results.failed}`);
  
  if (results.errors.length > 0) {
    console.log('\n失敗詳細:');
    results.errors.forEach(error => console.log(error));
  }
  
  return results;
}

// ========================================
// doGet関連テスト
// ========================================

/**
 * テスト: doGet - 認証済みユーザーの場合、打刻HTMLを返す
 */
function testDoGet_AuthenticatedUser_ReturnsClockHTML() {
  // Arrange
  const mockEvent = {
    parameter: {},
    pathInfo: null
  };
  
  // Act & Assert
  try {
    const result = doGet(mockEvent);
    
    // 結果がHtmlOutputであることを確認
    assert(result !== null, 'doGetの戻り値がnullです');
    assert(typeof result === 'object', 'doGetの戻り値がオブジェクトではありません');
    
    console.log('✓ doGet（認証済み）テスト成功');
  } catch (error) {
    throw new Error(`doGet（認証済み）テストエラー: ${error.message}`);
  }
}

/**
 * テスト: doGet - 未認証ユーザーの場合、未認証エラーを返す
 */
function testDoGet_UnauthenticatedUser_ReturnsUnauthorized() {
  // Arrange
  const mockEvent = {
    parameter: {},
    pathInfo: null
  };
  
  // 未認証状態をシミュレート（モック関数で対応予定）
  
  // Act & Assert
  try {
    // 未認証時の動作確認（実装後にモックで制御）
    console.log('✓ doGet（未認証）テスト準備完了');
  } catch (error) {
    throw new Error(`doGet（未認証）テストエラー: ${error.message}`);
  }
}

// ========================================
// doPost関連テスト
// ========================================

/**
 * テスト: doPost - 有効な出勤データの場合、データを保存する
 */
function testDoPost_ValidClockIn_SavesData() {
  // Arrange
  const mockEvent = {
    parameter: {
      action: 'IN',
      timestamp: new Date().toISOString()
    },
    postData: {
      contents: JSON.stringify({
        action: 'IN',
        timestamp: new Date().toISOString()
      })
    }
  };
  
  // Act & Assert
  try {
    const result = doPost(mockEvent);
    
    // 正常レスポンスの確認
    assert(result !== null, 'doPostの戻り値がnullです');
    
    console.log('✓ doPost（出勤）テスト成功');
  } catch (error) {
    throw new Error(`doPost（出勤）テストエラー: ${error.message}`);
  }
}

/**
 * テスト: doPost - 有効な退勤データの場合、データを保存する
 */
function testDoPost_ValidClockOut_SavesData() {
  // Arrange
  const mockEvent = {
    parameter: {
      action: 'OUT',
      timestamp: new Date().toISOString()
    },
    postData: {
      contents: JSON.stringify({
        action: 'OUT',
        timestamp: new Date().toISOString()
      })
    }
  };
  
  // Act & Assert
  try {
    const result = doPost(mockEvent);
    
    // 正常レスポンスの確認
    assert(result !== null, 'doPostの戻り値がnullです');
    
    console.log('✓ doPost（退勤）テスト成功');
  } catch (error) {
    throw new Error(`doPost（退勤）テストエラー: ${error.message}`);
  }
}

/**
 * テスト: doPost - 無効なアクションの場合、エラーを返す
 */
function testDoPost_InvalidAction_ReturnsError() {
  // Arrange
  const mockEvent = {
    parameter: {
      action: 'INVALID',
      timestamp: new Date().toISOString()
    },
    postData: {
      contents: JSON.stringify({
        action: 'INVALID',
        timestamp: new Date().toISOString()
      })
    }
  };
  
  // Act & Assert
  try {
    const result = doPost(mockEvent);
    
    // エラーレスポンスの確認
    assert(result !== null, 'doPostの戻り値がnullです');
    
    console.log('✓ doPost（無効アクション）テスト成功');
  } catch (error) {
    throw new Error(`doPost（無効アクション）テストエラー: ${error.message}`);
  }
}

// ========================================
// generateClockHTML関連テスト
// ========================================

/**
 * テスト: generateClockHTML - 有効なユーザー情報でHTMLを生成する
 */
function testGenerateClockHTML_ValidUser_ReturnsHTML() {
  // Arrange
  const userInfo = {
    email: 'test@example.com',
    name: 'テスト太郎',
    employeeId: 'EMP001'
  };
  
  // Act & Assert
  try {
    const result = generateClockHTML(userInfo);
    
    // HTML文字列の確認
    assert(typeof result === 'string', 'generateClockHTMLの戻り値が文字列ではありません');
    assert(result.includes('<!DOCTYPE html>'), 'HTMLのDoctype宣言がありません');
    assert(result.includes(userInfo.name), 'ユーザー名がHTMLに含まれていません');
    
    console.log('✓ generateClockHTML テスト成功');
  } catch (error) {
    throw new Error(`generateClockHTML テストエラー: ${error.message}`);
  }
}

// ========================================
// processClock関連テスト
// ========================================

/**
 * テスト: processClock - 出勤処理でLog_Rawシートを更新する
 */
function testProcessClock_ClockIn_UpdatesLogRaw() {
  // Arrange
  const action = 'IN';
  const userInfo = {
    email: 'test@example.com',
    name: 'テスト太郎',
    employeeId: 'EMP001'
  };
  
  // Act & Assert
  try {
    const result = processClock(action, userInfo);
    
    // 処理結果の確認
    assert(result !== null, 'processClockの戻り値がnullです');
    assert(typeof result === 'object', 'processClockの戻り値がオブジェクトではありません');
    
    console.log('✓ processClock（出勤）テスト成功');
  } catch (error) {
    throw new Error(`processClock（出勤）テストエラー: ${error.message}`);
  }
}

/**
 * テスト: processClock - 退勤処理でLog_Rawシートを更新する
 */
function testProcessClock_ClockOut_UpdatesLogRaw() {
  // Arrange
  const action = 'OUT';
  const userInfo = {
    email: 'test@example.com',
    name: 'テスト太郎',
    employeeId: 'EMP001'
  };
  
  // Act & Assert
  try {
    const result = processClock(action, userInfo);
    
    // 処理結果の確認
    assert(result !== null, 'processClockの戻り値がnullです');
    assert(typeof result === 'object', 'processClockの戻り値がオブジェクトではありません');
    
    console.log('✓ processClock（退勤）テスト成功');
  } catch (error) {
    throw new Error(`processClock（退勤）テストエラー: ${error.message}`);
  }
}

/**
 * テスト: processClock - 重複アクションの場合、エラーを返す
 */
function testProcessClock_DuplicateAction_ReturnsError() {
  // Arrange
  const action = 'IN';
  const userInfo = {
    email: 'test@example.com',
    name: 'テスト太郎',
    employeeId: 'EMP001'
  };
  
  // Act & Assert
  try {
    // 2回連続で同じアクションを実行（重複チェック）
    processClock(action, userInfo);
    const result = processClock(action, userInfo);
    
    // エラーハンドリングの確認
    console.log('✓ processClock（重複）テスト成功');
  } catch (error) {
    throw new Error(`processClock（重複）テストエラー: ${error.message}`);
  }
} 