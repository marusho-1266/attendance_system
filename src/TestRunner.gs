/**
 * 統一されたテストランナー
 * eval()を使用せず、関数参照による直接呼び出しを実装
 * セキュリティリスクを排除し、コード重複を削減
 */

/**
 * テスト実行結果の構造
 */
function createTestResult(name, success, duration, error) {
  return {
    name: name,
    success: success,
    duration: duration,
    error: error || null
  };
}

/**
 * 統一されたテスト実行関数
 * @param {Array} testFunctions - テスト関数の配列
 * @param {string} testSuiteName - テストスイート名
 * @param {Object} options - 実行オプション
 * @return {Object} テスト実行結果
 */
function runTestSuite(testFunctions, testSuiteName, options) {
  options = options || {};
  var iterations = options.iterations || 1;
  var showProgress = options.showProgress !== false;
  
  if (showProgress) {
    console.log('⚡ ' + testSuiteName + ' 開始');
    console.log('------------------------------');
  }
  
  var startTime = new Date().getTime();
  var testResults = [];
  var passedTests = 0;
  var failedTests = 0;
  
  testFunctions.forEach(function(testFunction, index) {
    var testName = typeof testFunction === 'string' ? testFunction : testFunction.name;
    var totalDuration = 0;
    var success = true;
    var error = null;
    
    // 複数回実行（パフォーマンステスト用）
    for (var i = 0; i < iterations; i++) {
      try {
        var testStartTime = new Date().getTime();
        
        if (typeof testFunction === 'function') {
          testFunction();
        } else {
          // 文字列の場合は関数を取得して実行
          var func = window[testFunction];
          if (typeof func === 'function') {
            func();
          } else {
            throw new Error('Test function not found: ' + testFunction);
          }
        }
        
        var testEndTime = new Date().getTime();
        totalDuration += (testEndTime - testStartTime);
        
      } catch (err) {
        success = false;
        error = err.message;
        if (showProgress) {
          console.log('❌ ' + testName + ' - 失敗: ' + error);
        }
        break; // エラーが発生したら残りの実行をスキップ
      }
    }
    
    var averageDuration = totalDuration / iterations;
    var result = createTestResult(testName, success, averageDuration, error);
    testResults.push(result);
    
    if (success) {
      passedTests++;
      if (showProgress) {
        if (iterations > 1) {
          console.log('⚡ ' + testName + ' - 平均: ' + averageDuration.toFixed(2) + 'ms (' + iterations + '回実行)');
        } else {
          console.log('✅ ' + testName + ' - 成功 (' + averageDuration + 'ms)');
        }
      }
    } else {
      failedTests++;
    }
  });
  
  var endTime = new Date().getTime();
  var totalDuration = endTime - startTime;
  var successRate = testResults.length > 0 ? (passedTests / testResults.length * 100).toFixed(1) : '0.0';
  
  if (showProgress) {
    console.log('------------------------------');
    console.log('⚡ ' + testSuiteName + ' 完了: ' + passedTests + '/' + testResults.length + ' (' + totalDuration + 'ms)');
    console.log('成功率: ' + successRate + '%');
    
    if (failedTests > 0) {
      console.log('失敗したテスト:');
      testResults.filter(function(result) { return !result.success; }).forEach(function(result) {
        console.log('  - ' + result.name + ': ' + result.error);
      });
    }
  }
  
  return {
    testSuiteName: testSuiteName,
    totalTests: testResults.length,
    passedTests: passedTests,
    failedTests: failedTests,
    successRate: successRate,
    totalDuration: totalDuration,
    results: testResults
  };
}

/**
 * クイックテスト実行
 * @param {Array} testFunctions - テスト関数の配列
 * @param {string} testSuiteName - テストスイート名
 * @return {Object} テスト実行結果
 */
function runQuickTests(testFunctions, testSuiteName) {
  return runTestSuite(testFunctions, testSuiteName, {
    iterations: 1,
    showProgress: true
  });
}

/**
 * パフォーマンステスト実行
 * @param {Array} testFunctions - テスト関数の配列
 * @param {string} testSuiteName - テストスイート名
 * @param {number} iterations - 実行回数（デフォルト: 10）
 * @return {Object} テスト実行結果
 */
function runPerformanceTests(testFunctions, testSuiteName, iterations) {
  return runTestSuite(testFunctions, testSuiteName, {
    iterations: iterations || 10,
    showProgress: true
  });
}

/**
 * 統合テスト実行
 * @param {Array} testFunctions - テスト関数の配列
 * @param {string} testSuiteName - テストスイート名
 * @return {Object} テスト実行結果
 */
function runIntegrationTests(testFunctions, testSuiteName) {
  return runTestSuite(testFunctions, testSuiteName, {
    iterations: 1,
    showProgress: true
  });
}

/**
 * 全テスト実行
 * @param {Array} testFunctions - テスト関数の配列
 * @param {string} testSuiteName - テストスイート名
 * @return {Object} テスト実行結果
 */
function runAllTests(testFunctions, testSuiteName) {
  return runTestSuite(testFunctions, testSuiteName, {
    iterations: 1,
    showProgress: false
  });
}

/**
 * テスト関数の存在確認
 * @param {string} functionName - 関数名
 * @return {boolean} 関数が存在するかどうか
 */
function testFunctionExists(functionName) {
  return typeof window[functionName] === 'function';
}

/**
 * テスト関数の取得
 * @param {string} functionName - 関数名
 * @return {Function|null} テスト関数またはnull
 */
function getTestFunction(functionName) {
  return window[functionName] || null;
} 