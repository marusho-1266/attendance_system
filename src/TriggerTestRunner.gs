/**
 * TriggerTestRunner.gs - Triggers機能専用テストランナー
 * フェーズ3 TDD実装用の高速テスト実行とレポート
 * 
 * 機能:
 * - Triggers関連のテスト実行と結果管理
 * - TDDサイクル用のクイックテスト
 * - Red-Green-Refactorサイクルの進捗追跡
 */

/**
 * Triggers機能のメインテストランナー
 * @return {Object} テスト結果
 */
function runTriggersMainTests() {
  console.log('🚀 Triggers.gs メインテスト開始');
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
    // Triggersテストの実行
    const triggerResults = runTriggersTestsDetailed();
    
    // 結果の集計
    results.total += triggerResults.total;
    results.passed += triggerResults.passed;
    results.failed += triggerResults.failed;
    results.errors = results.errors.concat(triggerResults.errors);
    
  } catch (error) {
    console.error('❌ Triggersテスト実行エラー:', error);
    results.failed++;
    results.errors.push(`Triggersテスト実行エラー: ${error.message}`);
  }
  
  // 実行時間の計算
  const endTime = new Date();
  results.duration = endTime.getTime() - startTime.getTime();
  
  // 結果サマリーの表示
  console.log('\n' + '='.repeat(50));
  console.log('📊 Triggers.gs テスト結果サマリー');
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
    console.log('📝 TDD次ステップ: Refactorフェーズ（バッチ処理の詳細実装）');
  } else {
    console.log('\n🔧 TDD次ステップ: Green実装の修正が必要');
  }
  
  return results;
}

/**
 * Triggers機能のクイックテスト（TDD用）
 * @return {Object} テスト結果
 */
function runTriggersQuickTest() {
  console.log('⚡ Triggers.gs クイックテスト開始');
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
    testOnOpen_SpreadsheetOpen_AddsManagementMenu,
    testDailyJob_ExecuteDaily_UpdatesDailySummary,
    testWeeklyOvertimeJob_Monday_CalculatesOvertimeHours,
    testMonthlyJob_FirstDayOfMonth_UpdatesMonthlySummary
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
 * Triggers詳細テスト実行（内部使用）
 * @return {Object} 詳細テスト結果
 */
function runTriggersTestsDetailed() {
  console.log('=== Triggers.gs 詳細テスト実行開始 ===');
  
  const testResults = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: []
  };
  
  // テスト関数リスト
  const testFunctions = [
    // onOpen関数のテスト
    'testOnOpen_SpreadsheetOpen_AddsManagementMenu',
    'testOnOpen_ErrorHandling_DoesNotThrow',
    
    // dailyJob関数のテスト
    'testDailyJob_ExecuteDaily_UpdatesDailySummary',
    'testDailyJob_HasUnfinishedClockOut_SendsEmail',
    'testDailyJob_QuotaCompliance_StaysWithinLimits',
    
    // weeklyOvertimeJob関数のテスト
    'testWeeklyOvertimeJob_Monday_CalculatesOvertimeHours',
    'testWeeklyOvertimeJob_HighOvertimeDetected_SendsWarningEmail',
    
    // monthlyJob関数のテスト
    'testMonthlyJob_FirstDayOfMonth_UpdatesMonthlySummary',
    'testMonthlyJob_MonthlyProcess_ExportsCSVToDrive',
    'testMonthlyJob_CSVGenerated_SendsLinkEmail',
    
    // 統合テスト
    'testTriggerIntegration_AllFunctions_ExecuteSuccessfully',
    'testTriggerErrorHandling_ExceptionOccurs_SystemContinues'
  ];
  
  console.log('実行予定テスト数: ' + testFunctions.length);
  
  // 各テスト実行
  testFunctions.forEach(function(testFunctionName) {
    testResults.total++;
    
    try {
      console.log('実行中: ' + testFunctionName);
      
      // テスト関数を実行
      var testFunction = eval(testFunctionName);
      testFunction();
      
      testResults.passed++;
      console.log('✓ PASS: ' + testFunctionName);
      
    } catch (error) {
      testResults.failed++;
      testResults.errors.push({
        test: testFunctionName,
        error: error.message,
        stack: error.stack
      });
      console.log('✗ FAIL: ' + testFunctionName + ' - ' + error.message);
    }
  });
  
  // 結果をコンソール用フォーマットに変換
  const consoleErrors = testResults.errors.map(e => `${e.test}: ${e.error}`);
  
  return {
    total: testResults.total,
    passed: testResults.passed,
    failed: testResults.failed,
    errors: consoleErrors
  };
}

/**
 * Triggers統合テスト
 * @return {Object} テスト結果
 */
function runTriggersIntegrationTest() {
  console.log('🔗 Triggers.gs 統合テスト開始');
  console.log('-'.repeat(30));
  
  const startTime = new Date();
  let results = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: []
  };
  
  // 統合テストシナリオ
  const integrationScenarios = [
    {
      name: 'onOpen実行→メニュー確認',
      test: () => {
        const result = onOpen();
        if (!result || !result.success) {
          throw new Error('onOpen実行失敗');
        }
      }
    },
    {
      name: 'dailyJob実行→結果確認',
      test: () => {
        const result = dailyJob();
        if (!result || !result.success) {
          throw new Error('dailyJob実行失敗');
        }
        if (result.duration > 60000) {
          throw new Error(`dailyJob実行時間超過: ${result.duration}ms`);
        }
      }
    },
    {
      name: 'weeklyOvertimeJob実行→残業チェック',
      test: () => {
        const result = weeklyOvertimeJob();
        if (!result || !result.success) {
          throw new Error('weeklyOvertimeJob実行失敗');
        }
        if (typeof result.highOvertimeCount !== 'number') {
          throw new Error('残業超過者数が不正');
        }
      }
    },
    {
      name: 'monthlyJob実行→CSV出力確認',
      test: () => {
        const result = monthlyJob();
        if (!result || !result.success) {
          throw new Error('monthlyJob実行失敗');
        }
        if (!result.csvResult || !result.csvResult.fileName) {
          throw new Error('CSV出力結果が不正');
        }
      }
    }
  ];
  
  integrationScenarios.forEach(scenario => {
    results.total++;
    try {
      console.log(`🔗 ${scenario.name}`);
      scenario.test();
      results.passed++;
      console.log(`✅ PASS`);
    } catch (error) {
      results.failed++;
      results.errors.push(`${scenario.name}: ${error.message}`);
      console.log(`❌ FAIL: ${error.message}`);
    }
  });
  
  const endTime = new Date();
  const duration = endTime.getTime() - startTime.getTime();
  
  console.log('-'.repeat(30));
  console.log(`🔗 統合テスト完了: ${results.passed}/${results.total} (${duration}ms)`);
  
  return results;
}

/**
 * Triggersパフォーマンステスト
 * クォータ制限の確認と実行時間測定
 * @return {Object} パフォーマンステスト結果
 */
function runTriggersPerformanceTest() {
  console.log('⚡ Triggers.gs パフォーマンステスト開始');
  console.log('-'.repeat(30));
  
  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: [],
    performance: {}
  };
  
  // パフォーマンステスト項目
  const performanceTests = [
    {
      name: 'dailyJob実行時間測定',
      target: 60000, // 1分以内目標
      test: () => {
        const startTime = new Date().getTime();
        const result = dailyJob();
        const endTime = new Date().getTime();
        const duration = endTime - startTime;
        
        results.performance.dailyJob = duration;
        
        if (duration > 60000) {
          throw new Error(`実行時間超過: ${duration}ms > 60000ms`);
        }
        
        return duration;
      }
    },
    {
      name: 'weeklyOvertimeJob実行時間測定',
      target: 120000, // 2分以内目標
      test: () => {
        const startTime = new Date().getTime();
        const result = weeklyOvertimeJob();
        const endTime = new Date().getTime();
        const duration = endTime - startTime;
        
        results.performance.weeklyOvertimeJob = duration;
        
        if (duration > 120000) {
          throw new Error(`実行時間超過: ${duration}ms > 120000ms`);
        }
        
        return duration;
      }
    },
    {
      name: 'monthlyJob実行時間測定',
      target: 300000, // 5分以内目標
      test: () => {
        const startTime = new Date().getTime();
        const result = monthlyJob();
        const endTime = new Date().getTime();
        const duration = endTime - startTime;
        
        results.performance.monthlyJob = duration;
        
        if (duration > 300000) {
          throw new Error(`実行時間超過: ${duration}ms > 300000ms`);
        }
        
        return duration;
      }
    }
  ];
  
  performanceTests.forEach(test => {
    results.total++;
    try {
      console.log(`⚡ ${test.name}`);
      const duration = test.test();
      results.passed++;
      console.log(`✅ PASS: ${duration}ms (目標: ${test.target}ms)`);
    } catch (error) {
      results.failed++;
      results.errors.push(`${test.name}: ${error.message}`);
      console.log(`❌ FAIL: ${error.message}`);
    }
  });
  
  // パフォーマンス結果サマリー
  console.log('\n📊 パフォーマンス結果:');
  Object.keys(results.performance).forEach(key => {
    console.log(`  ${key}: ${results.performance[key]}ms`);
  });
  
  return results;
}

/**
 * Triggers機能の最終テスト（完了確認用）
 * @return {Object} 最終テスト結果
 */
function runTriggersFinalTest() {
  console.log('🏁 Triggers.gs 最終テスト開始');
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
  const mainResults = runTriggersMainTests();
  finalResults.modules.push({ name: 'メインテスト', results: mainResults });
  
  // 2. 統合テスト
  console.log('\n2️⃣ 統合テスト実行中...');
  const integrationResults = runTriggersIntegrationTest();
  finalResults.modules.push({ name: '統合テスト', results: integrationResults });
  
  // 3. パフォーマンステスト
  console.log('\n3️⃣ パフォーマンステスト実行中...');
  const perfResults = runTriggersPerformanceTest();
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
  console.log('🏆 Triggers.gs 最終テスト結果');
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
  
  // TDD進捗レポート
  if (finalResults.failed === 0) {
    console.log('\n🎉 フェーズ3 triggers-1 TDD実装完了！');
    console.log('📝 次のマイルストーン: フェーズ3-form-1（フォーム連携）実装');
  } else {
    console.log('\n🔧 Refactorフェーズで改善が必要');
    console.log('失敗したテスト数: ' + finalResults.failed);
  }
  
  return finalResults;
}

/**
 * TDD状況レポート（Triggers専用）
 */
function showTriggersTDDStatus() {
  console.log('=== Triggers TDD進捗レポート ===');
  
  const tddStatus = {
    'フェーズ3: 集計機能': {
      'triggers-1': 'Green（実装完了）',
      'form-1': 'Red（未実装）',
      'mail-1': 'Red（未実装）'
    }
  };
  
  Object.keys(tddStatus).forEach(function(phase) {
    console.log('\n' + phase + ':');
    const tasks = tddStatus[phase];
    Object.keys(tasks).forEach(function(task) {
      console.log('  ' + task + ': ' + tasks[task]);
    });
  });
  
  console.log('\n現在のステップ: triggers-1 Greenフェーズ完了、Refactorフェーズ準備中');
  console.log('次のステップ: triggers-1のRefactor実装（詳細なバッチ処理ロジック）');
} 