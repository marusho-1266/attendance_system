# 設計書

## 概要

本設計書は、Google Apps ScriptとGoogle スプレッドシートを基盤とした出勤管理システムのアーキテクチャと実装詳細を定義します。システムは50名以下の小規模組織向けに設計され、Googleの無料枠制限内で動作し、堅牢なエラーハンドリングと効率的なクォータ管理を実現します。

## アーキテクチャ

### システム全体構成

```mermaid
graph TB
    A[従業員] --> B[打刻WebApp]
    A --> C[Google Forms]
    B --> D[Apps Script]
    C --> D
    D --> E[Google Sheets]
    D --> F[Gmail通知]
    D --> G[Google Drive]
    
    E --> H[Master_Employee]
    E --> I[Master_Holiday]
    E --> J[Log_Raw]
    E --> K[Daily_Summary]
    E --> L[Monthly_Summary]
    E --> M[Request_Responses]
    
    N[時間トリガー] --> D
    O[フォームトリガー] --> D
```

### レイヤー構成

1. **プレゼンテーション層**
   - 打刻WebApp（HTML/JavaScript）
   - Google Forms（申請用）

2. **ビジネスロジック層**
   - Apps Script（認証、計算、ワークフロー）
   - 自動トリガー処理

3. **データ層**
   - Google Sheets（6つのシート）
   - Google Drive（CSVエクスポート）

## コンポーネントと インターフェース

### 1. 認証・セキュリティコンポーネント

#### Authentication.gs
```javascript
// 主要関数
function authenticateUser()
function validateEmployeeAccess(email)
function getEmployeeInfo(email)
```

**責任:**
- `Session.getActiveUser().getEmail()`による認証
- 従業員マスタとの照合
- アクセス権限の検証

### 2. 打刻処理コンポーネント

#### WebApp.gs
```javascript
// 主要関数
function doGet(e)
function doPost(e)
function recordTimeEntry(employeeId, action)
function validateTimeEntry(employeeId, action, date)
```

**責任:**
- 打刻WebAppのUI提供
- 打刻データの受付と検証
- 重複チェック・整合性チェック
- Log_Rawシートへの記録

### 3. 業務ロジックコンポーネント

#### BusinessLogic.gs
```javascript
// 主要関数
function calculateDailyWorkTime(employeeId, date)
function calculateOvertime(workTime, standardTime)
function applyBreakDeduction(workTime)
function checkHoliday(date)
function updateDailySummary(employeeId, date)
```

**責任:**
- 労働時間計算（実働、残業、深夜）
- 休憩時間の自動控除
- 祝日・休日判定
- 遅刻・早退の判定

### 4. 承認ワークフローコンポーネント

#### FormManager.gs
```javascript
// 主要関数
function onRequestSubmit(e)
function processApprovalRequest(requestData)
function batchApprovalNotifications()
function updateRequestStatus(requestId, status)
```

**責任:**
- フォーム申請の受付
- 承認者の特定
- 承認通知のバッチ処理
- ステータス更新の処理

### 5. 自動処理コンポーネント

#### Triggers.gs
```javascript
// 主要関数
function dailyJob()
function weeklyOvertimeJob()
function monthlyJob()
function onOpen()
```

**責任:**
- 日次サマリー更新（02:00）
- 週次残業アラート（月曜07:00）
- 月次レポート生成（1日02:30）
- 未退勤者通知

### 6. 通知・メール管理コンポーネント

#### MailManager.gs
```javascript
// 主要関数
function sendBatchNotification(recipients, subject, body)
function sendErrorAlert(error, context)
function sendOvertimeWarning(employees)
function sendMonthlyReport(csvLinks)
```

**責任:**
- バッチメール送信（クォータ節約）
- エラー通知
- 各種アラート送信
- レポート配信

### 7. ユーティリティコンポーネント

#### Utils.gs
```javascript
// 主要関数
function isHoliday(date)
function formatTime(minutes)
function getSheetData(sheetName, range)
function setSheetData(sheetName, range, values)
function exportToCsv(sheetName, fileName)
```

**責任:**
- 共通ユーティリティ関数
- データアクセス抽象化
- 日付・時間処理
- CSVエクスポート

## データモデル

### スプレッドシート設計

#### 1. Master_Employee
```
列A: 社員ID (文字列, 主キー)
列B: 氏名 (文字列)
列C: Gmail (文字列, 認証用)
列D: 所属 (文字列)
列E: 雇用区分 (文字列)
列F: 上司Gmail (文字列, 承認者特定用)
列G: 基準始業 (時刻)
列H: 基準終業 (時刻)
```

#### 2. Master_Holiday
```
列A: 日付 (日付型, 主キー)
列B: 名称 (文字列)
列C: 法定休日フラグ (ブール)
列D: 会社休日フラグ (ブール)
```

#### 3. Log_Raw（保護対象）
```
列A: タイムスタンプ (日時)
列B: 社員ID (文字列)
列C: 氏名 (文字列)
列D: Action (IN/OUT/BRK_IN/BRK_OUT)
列E: 端末IP (文字列)
列F: 備考 (文字列)
```

#### 4. Daily_Summary
```
列A: 日付 (日付型)
列B: 社員ID (文字列)
列C: 出勤時刻 (時刻)
列D: 退勤時刻 (時刻)
列E: 休憩時間 (分)
列F: 実働時間 (分)
列G: 残業時間 (分)
列H: 遅刻早退 (分)
列I: 承認状態 (文字列)
```

#### 5. Monthly_Summary
```
列A: 年月 (文字列, YYYY-MM)
列B: 社員ID (文字列)
列C: 勤務日数 (数値)
列D: 総労働時間 (分)
列E: 残業時間 (分)
列F: 有給日数 (数値)
列G: 備考 (文字列)
```

#### 6. Request_Responses
```
列A: タイムスタンプ (日時)
列B: 社員ID (文字列)
列C: 種別 (修正/残業/休暇)
列D: 詳細 (文字列)
列E: 希望値 (文字列)
列F: 承認者 (文字列)
列G: Status (Pending/Approved/Rejected, プルダウン)
```

### データアクセスパターン

#### バッチ処理最適化
```javascript
// 効率的なデータ読み書き
const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
// 処理...
sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
```

#### 差分更新戦略
```javascript
// 更新対象の従業員のみ処理
const updatedEmployees = getUpdatedEmployees(date);
updatedEmployees.forEach(emp => updateDailySummary(emp.id, date));
```

## エラーハンドリング

### エラー処理戦略

#### 1. 包括的try-catch実装
```javascript
function dailyJob() {
  try {
    // メイン処理
    updateDailySummaries();
    sendNotifications();
  } catch (error) {
    Logger.log('Daily job failed: ' + error.toString());
    MailManager.sendErrorAlert(error, 'dailyJob');
    throw error; // 再スロー for monitoring
  }
}
```

#### 2. エラー分類と対応

**システムエラー:**
- Apps Script実行時間超過
- クォータ制限到達
- シートアクセス権限エラー

**ビジネスロジックエラー:**
- データ整合性エラー
- 計算エラー
- 承認フローエラー

**外部サービスエラー:**
- Gmail送信エラー
- Drive書き込みエラー

#### 3. エラー通知システム
```javascript
function sendErrorAlert(error, context) {
  const subject = `[緊急] 出勤管理システムエラー: ${context}`;
  const body = `
エラー詳細: ${error.message}
発生時刻: ${new Date()}
コンテキスト: ${context}
スタックトレース: ${error.stack}
  `;
  GmailApp.sendEmail(ADMIN_EMAIL, subject, body);
}
```

### 復旧戦略

#### 1. 自動復旧機能
- 一時的なエラーに対する再試行機能
- 部分的な処理継続機能

#### 2. 手動復旧支援
- エラー状況の詳細ログ
- 復旧手順の文書化
- データ整合性チェック機能

## テスト戦略

### テスト構成

#### 1. 単体テスト
```javascript
// BusinessLogicTest.gs
function testCalculateWorkTime() {
  const result = calculateDailyWorkTime('EMP001', new Date('2024-01-15'));
  assertEquals(480, result.workMinutes); // 8時間
}
```

#### 2. 統合テスト
```javascript
// IntegrationTest.gs
function testEndToEndTimeEntry() {
  // 打刻 → 計算 → サマリー更新の一連の流れをテスト
}
```

#### 3. テストカバレッジ管理
```javascript
// Test.gs
function showTestCoverage() {
  const moduleInfo = {
    'BusinessLogic.gs': ['calculateWorkTime', 'applyBreakDeduction'],
    'Authentication.gs': ['authenticateUser', 'validateAccess']
  };
  // カバレッジレポート生成
}
```

### テスト実行戦略

#### 1. 開発時テスト
- 関数単位での単体テスト
- モックデータを使用した統合テスト

#### 2. デプロイ前テスト
- 本番データのコピーを使用したテスト
- パフォーマンステスト

#### 3. 本番監視
- 日次処理の成功/失敗監視
- エラーログの定期確認

## セキュリティ設計

### アクセス制御

#### 1. 認証方式
```javascript
function authenticateUser() {
  const userEmail = Session.getActiveUser().getEmail();
  const employee = getEmployeeByEmail(userEmail);
  if (!employee) {
    throw new Error('認証失敗: 従業員マスタに存在しません');
  }
  return employee;
}
```

#### 2. シート保護設定

**Log_Raw:** 完全保護（管理者のみアクセス）
**Summary系:** 全員読み取り専用
**Request_Responses:** 承認者は担当行のステータス列のみ編集可

#### 3. スクリプト権限管理
- 共有組織アカウントでの所有
- 編集権限は管理者のみ
- 実行権限は必要最小限

### データ保護

#### 1. 個人情報保護
- 必要最小限のデータ収集
- ログの定期アーカイブ
- 退職者データの適切な処理

#### 2. 監査ログ
```javascript
function logAccess(action, employeeId, details) {
  const logEntry = [
    new Date(),
    employeeId,
    action,
    details,
    Session.getActiveUser().getEmail()
  ];
  getSheet('Access_Log').appendRow(logEntry);
}
```

## パフォーマンス最適化

### クォータ管理戦略

#### 1. メール送信最適化
```javascript
// バッチ送信でクォータ節約
function sendBatchNotifications(notifications) {
  const groupedByRecipient = groupNotifications(notifications);
  Object.keys(groupedByRecipient).forEach(recipient => {
    const messages = groupedByRecipient[recipient];
    const combinedMessage = combineMessages(messages);
    GmailApp.sendEmail(recipient, combinedMessage.subject, combinedMessage.body);
  });
}
```

#### 2. 実行時間最適化
```javascript
// チャンク処理で実行時間制限回避
function processLargeDataset(data) {
  const CHUNK_SIZE = 100;
  for (let i = 0; i < data.length; i += CHUNK_SIZE) {
    const chunk = data.slice(i, i + CHUNK_SIZE);
    processChunk(chunk);
    
    // 実行時間チェック
    if (new Date() - startTime > MAX_EXECUTION_TIME) {
      scheduleNextExecution(i + CHUNK_SIZE);
      break;
    }
  }
}
```

#### 3. データアクセス最適化
```javascript
// 一括読み書きでAPI呼び出し削減
function optimizedDataUpdate(updates) {
  const sheet = getSheet('Daily_Summary');
  const range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  const values = range.getValues();
  
  // メモリ上で更新処理
  updates.forEach(update => {
    values[update.row][update.col] = update.value;
  });
  
  // 一括書き込み
  range.setValues(values);
}
```

### スケーラビリティ考慮

#### 1. データ量増加対応
- 月次アーカイブ機能
- 古いデータの別シート移動
- 検索範囲の最適化

#### 2. ユーザー数増加対応
- 部門別処理の並列化
- 処理優先度の設定
- 負荷分散の検討

## 運用・保守設計

### 監視機能

#### 1. システム稼働監視
```javascript
function healthCheck() {
  const checks = [
    checkSheetAccess(),
    checkTriggerStatus(),
    checkQuotaUsage()
  ];
  
  const failures = checks.filter(check => !check.success);
  if (failures.length > 0) {
    sendHealthAlert(failures);
  }
}
```

#### 2. データ整合性チェック
```javascript
function validateDataIntegrity() {
  const issues = [
    checkMissingClockOuts(),
    checkCalculationErrors(),
    checkApprovalConsistency()
  ];
  
  return issues.filter(issue => issue.severity === 'HIGH');
}
```

### バックアップ・復旧

#### 1. 自動バックアップ
```javascript
function createBackup() {
  const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
  const backupName = `出勤管理_バックアップ_${timestamp}`;
  
  const originalFile = DriveApp.getFileById(SPREADSHEET_ID);
  const backup = originalFile.makeCopy(backupName);
  
  // バックアップフォルダに移動
  DriveApp.getFolderById(BACKUP_FOLDER_ID).addFile(backup);
}
```

#### 2. 復旧手順
- バックアップからの復元手順
- データ整合性の確認方法
- 段階的復旧プロセス