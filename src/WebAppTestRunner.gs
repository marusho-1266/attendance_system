/**
 * WebAppTestRunner.gs - WebApp機能専用テストランナー
 * 
 * 機能:
 * - WebApp関連のテスト実行と結果管理
 * - TDDサイクル用のクイックテスト
 * - 統合テストとパフォーマンステスト
 */

/**
 * WebApp機能のメインテストランナー
 * @return {Object} テスト結果
 */
function runWebAppMainTests() {
  console.log('🚀 WebApp.gs メインテスト開始');
  console.log('='.repeat(50));
  
  const startTime = new Date();
  let results = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: [],
    duration: 0
  };
  
  try {
    // WebAppテストの実行
    const webAppResults = runWebAppTests();
    
    // 結果の集計
    results.total += webAppResults.total;
    results.passed += webAppResults.passed;
    results.failed += webAppResults.failed;
    results.errors = results.errors.concat(webAppResults.errors);
    
  } catch (error) {
    console.error('❌ WebAppテスト実行エラー:', error);
    results.failed++;
    results.errors.push(`WebAppテスト実行エラー: ${error.message}`);
  }
  
  // 実行時間の計算
  const endTime = new Date();
  results.duration = endTime.getTime() - startTime.getTime();
  
  // 結果サマリーの表示
  console.log('\n' + '='.repeat(50));
  console.log('📊 WebApp.gs テスト結果サマリー');
  console.log('='.repeat(50));
  console.log(`📈 総テスト数: ${results.total}`);
  console.log(`✅ 成功: ${results.passed}`);
  console.log(`❌ 失敗: ${results.failed}`);
  console.log(`⏱️  実行時間: ${results.duration}ms`);
  console.log(`📊 成功率: ${results.total > 0 ? Math.round((results.passed / results.total) * 100) : 0}%`);
  
  if (results.failed > 0) {
    console.log('\n🔍 失敗詳細:');
    results.errors.forEach((error, index) => {
      console.log(`${index + 1}. ${error}`);
    });
  }
  
  // TDD次ステップの提案
  if (results.failed === 0) {
    console.log('\n🎉 全テスト成功！');
    console.log('📝 TDD次ステップ: Refactorフェーズ（コード改善）');
  } else {
    console.log('\n🔧 TDD次ステップ: Green実装の修正が必要');
  }
  
  return results;
}

/**
 * WebApp機能のクイックテスト（TDD用）
 * @return {Object} テスト結果
 */
function runWebAppQuickTest() {
  console.log('⚡ WebApp.gs クイックテスト開始');
  console.log('-'.repeat(30));
  
  const startTime = new Date();
  let results = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: []
  };
  
  // 重要なテストのみ実行（TDD高速サイクル用）
  const quickTests = [
    testGenerateClockHTML_ValidUser_ReturnsHTML,
    testProcessClock_ClockIn_UpdatesLogRaw,
    testDoGet_AuthenticatedUser_ReturnsClockHTML
  ];
  
  quickTests.forEach(test => {
    results.total++;
    try {
      console.log(`⚡ ${test.name}`);
      test();
      results.passed++;
      console.log(`✅ PASS`);
    } catch (error) {
      results.failed++;
      results.errors.push(`${test.name}: ${error.message}`);
      console.log(`❌ FAIL: ${error.message}`);
    }
  });
  
  const endTime = new Date();
  const duration = endTime.getTime() - startTime.getTime();
  
  console.log('-'.repeat(30));
  console.log(`⚡ クイックテスト完了: ${results.passed}/${results.total} (${duration}ms)`);
  
  return results;
}

/**
 * WebApp統合テスト
 * @return {Object} テスト結果
 */
function runWebAppIntegrationTest() {
  console.log('🔗 WebApp.gs 統合テスト開始');
  console.log('-'.repeat(40));
  
  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: []
  };
  
  try {
    // 統合テスト1: doGet → generateClockHTML の連携
    results.total++;
    console.log('🔗 統合テスト1: doGet → generateClockHTML');
    
    const mockEvent = { parameter: {}, pathInfo: null };
    const htmlOutput = doGet(mockEvent);
    
    assert(htmlOutput !== null, 'doGetの戻り値がnull');
    assert(typeof htmlOutput === 'object', 'doGetの戻り値がオブジェクトでない');
    
    results.passed++;
    console.log('✅ 統合テスト1 成功');
    
  } catch (error) {
    results.failed++;
    results.errors.push(`統合テスト1: ${error.message}`);
    console.log(`❌ 統合テスト1 失敗: ${error.message}`);
  }
  
  try {
    // 統合テスト2: doPost → processClock の連携
    results.total++;
    console.log('🔗 統合テスト2: doPost → processClock');
    
    const mockPostEvent = {
      parameter: { action: 'IN' },
      postData: {
        contents: JSON.stringify({
          action: 'IN',
          timestamp: new Date().toISOString()
        })
      }
    };
    
    const jsonResponse = doPost(mockPostEvent);
    
    assert(jsonResponse !== null, 'doPostの戻り値がnull');
    assert(typeof jsonResponse === 'object', 'doPostの戻り値がオブジェクトでない');
    
    results.passed++;
    console.log('✅ 統合テスト2 成功');
    
  } catch (error) {
    results.failed++;
    results.errors.push(`統合テスト2: ${error.message}`);
    console.log(`❌ 統合テスト2 失敗: ${error.message}`);
  }
  
  console.log('-'.repeat(40));
  console.log(`🔗 統合テスト完了: ${results.passed}/${results.total}`);
  
  return results;
}

/**
 * WebApp機能の全包括テスト（完了確認用）
 * @return {Object} 最終テスト結果
 */
function runWebAppFinalTest() {
  console.log('🏁 WebApp.gs 最終テスト開始');
  console.log('='.repeat(60));
  
  const finalResults = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: [],
    modules: []
  };
  
  // 1. メインテスト
  console.log('1️⃣ メインテスト実行中...');
  const mainResults = runWebAppMainTests();
  finalResults.modules.push({ name: 'メインテスト', results: mainResults });
  
  // 2. 統合テスト
  console.log('\n2️⃣ 統合テスト実行中...');
  const integrationResults = runWebAppIntegrationTest();
  finalResults.modules.push({ name: '統合テスト', results: integrationResults });
  
  // 3. パフォーマンステスト
  console.log('\n3️⃣ パフォーマンステスト実行中...');
  const perfResults = runWebAppPerformanceTest();
  finalResults.modules.push({ name: 'パフォーマンステスト', results: perfResults });
  
  // 結果の集計
  finalResults.modules.forEach(module => {
    finalResults.total += module.results.total;
    finalResults.passed += module.results.passed;
    finalResults.failed += module.results.failed;
    finalResults.errors = finalResults.errors.concat(module.results.errors);
  });
  
  // 最終結果の表示
  console.log('\n' + '='.repeat(60));
  console.log('🏆 WebApp.gs 最終テスト結果');
  console.log('='.repeat(60));
  
  finalResults.modules.forEach(module => {
    const success = module.results.failed === 0 ? '✅' : '❌';
    console.log(`${success} ${module.name}: ${module.results.passed}/${module.results.total}`);
  });
  
  console.log('\n📊 総合結果:');
  console.log(`📈 総テスト数: ${finalResults.total}`);
  console.log(`✅ 成功: ${finalResults.passed}`);
  console.log(`❌ 失敗: ${finalResults.failed}`);
  console.log(`📊 成功率: ${finalResults.total > 0 ? Math.round((finalResults.passed / finalResults.total) * 100) : 0}%`);
  
  if (finalResults.failed === 0) {
    console.log('\n🎉 WebApp.gs 実装完了！');
    console.log('📝 次のマイルストーン: フェーズ3（集計機能のTDD実装）');
  } else {
    console.log('\n🔧 修正が必要な項目があります');
    console.log('📝 TDD次ステップ: 失敗テストの修正');
  }
  
  return finalResults;
}

/**
 * WebAppパフォーマンステスト
 * @return {Object} パフォーマンステスト結果
 */
function runWebAppPerformanceTest() {
  console.log('⚡ WebApp.gs パフォーマンステスト');
  console.log('-'.repeat(35));
  
  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: [],
    metrics: {}
  };
  
  try {
    // HTML生成のパフォーマンス
    results.total++;
    console.log('⚡ HTML生成速度テスト');
    
    const startTime = new Date();
    const userInfo = {
      email: 'test@example.com',
      name: 'テストユーザー',
      employeeId: 'EMP999'
    };
    
    const html = generateClockHTML(userInfo);
    const endTime = new Date();
    const duration = endTime.getTime() - startTime.getTime();
    
    results.metrics.htmlGeneration = duration;
    
    // 100ms以下であれば成功
    if (duration < 100) {
      results.passed++;
      console.log(`✅ HTML生成: ${duration}ms (OK)`);
    } else {
      results.failed++;
      results.errors.push(`HTML生成が遅い: ${duration}ms (>100ms)`);
      console.log(`❌ HTML生成: ${duration}ms (遅い)`);
    }
    
  } catch (error) {
    results.failed++;
    results.errors.push(`パフォーマンステストエラー: ${error.message}`);
  }
  
  console.log('-'.repeat(35));
  console.log(`⚡ パフォーマンステスト完了: ${results.passed}/${results.total}`);
  
  return results;
}

/**
 * WebApp機能のヘルスチェック（開発中の動作確認用）
 * @return {boolean} 正常動作フラグ
 */
function webAppHealthCheck() {
  console.log('🏥 WebApp.gs ヘルスチェック');
  
  try {
    // 基本関数の存在チェック
    const functions = [
      'doGet', 'doPost', 'generateClockHTML', 'processClock'
    ];
    
    functions.forEach(funcName => {
      if (typeof eval(funcName) !== 'function') {
        throw new Error(`関数 ${funcName} が定義されていません`);
      }
    });
    
    console.log('✅ 全ての関数が定義されています');
    
    // クイックテスト実行
    const quickResults = runWebAppQuickTest();
    const isHealthy = quickResults.failed === 0;
    
    console.log(isHealthy ? '✅ ヘルスチェック: 正常' : '❌ ヘルスチェック: 異常');
    return isHealthy;
    
  } catch (error) {
    console.log(`❌ ヘルスチェック: エラー - ${error.message}`);
    return false;
  }
} 