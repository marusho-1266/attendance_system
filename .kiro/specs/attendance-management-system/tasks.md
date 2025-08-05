# 実装計画

- [x] 1. プロジェクト構造とコア設定の初期化
  - Google Apps Scriptプロジェクトの基本構造を作成
  - appsscript.jsonファイルで必要な権限とライブラリを設定
  - 基本的な定数とグローバル設定を定義
  - _要件: 1.1, 7.5_

- [x] 2. Google スプレッドシートの構造とマスターデータの設定
  - 6つのシート（Master_Employee, Master_Holiday, Log_Raw, Daily_Summary, Monthly_Summary, Request_Responses）を作成
  - 各シートのヘッダー行と列構造を定義
  - シート保護設定とアクセス権限を実装
  - 入力規則（プルダウン）をRequest_ResponsesのStatus列に設定
  - _要件: 7.2, 7.3, 7.4, 3.3_

- [x] 3. 認証・セキュリティモジュールの実装
  - Authentication.gsファイルを作成
  - Session.getActiveUser().getEmail()を使用した認証機能を実装
  - 従業員マスターとの照合機能を実装
  - アクセス権限検証機能を実装
  - 認証失敗時のエラーハンドリングを実装
  - _要件: 1.1, 7.1_

- [x] 4. ユーティリティ関数の実装
  - Utils.gsファイルを作成
  - 祝日判定機能（isHoliday）を実装
  - 時間計算ユーティリティ関数を実装
  - データアクセス抽象化関数（getSheetData, setSheetData）を実装
  - 日付・時刻フォーマット関数を実装
  - _要件: 5.3, 5.4_

- [x] 5. 業務ロジックモジュールの実装
  - BusinessLogic.gsファイルを作成
  - 日次労働時間計算機能（calculateDailyWorkTime）を実装
  - 残業時間計算機能（calculateOvertime）を実装
  - 休憩時間自動控除機能（applyBreakDeduction）を実装
  - 遅刻・早退判定機能を実装
  - 深夜勤務時間計算機能を実装
  - _要件: 5.1, 5.2, 5.5_

- [x] 6. 打刻WebAppの実装
  - WebApp.gsファイルを作成
  - doGet関数で打刻画面のHTMLを提供
  - doPost関数で打刻データを受信・処理
  - 打刻データの検証機能（重複チェック、整合性チェック）を実装
  - Log_Rawシートへのデータ記録機能を実装
  - 打刻完了画面の表示機能を実装
  - 同時実行制御機能の実装
    - LockServiceを使用した書き込み競合防止機能を実装
    - バッチ書き込み機能による複数データの一括処理を実装
    - 同時送信時のキューイング機能を実装
    - 書き込み失敗時のリトライ機能を実装
    - データ整合性保証機能（GAS制約対応版）を実装
      - 書き込み前のバックアップ行保持による疑似ロールバック機能を実装
      - バージョン番号またはタイムスタンプチェックによる重複書き込み防止機能を実装
      - エラーログ記録と管理者通知機能を実装
      - 手動復旧プロセスとデータ整合性チェック機能を実装
  - _要件: 1.2, 1.3, 1.4, 1.5, 6.1, 6.3_

- [x] 7. フォーム管理・承認ワークフローの実装
  - FormManager.gsファイルを作成
  - onRequestSubmit関数でフォーム申請を受信
  - 申請データをRequest_Responsesシートに記録（ステータス"Pending"）
  - 承認者特定機能（Master_Employeeから上司情報を取得）
  - 申請データの検証機能を実装
  - CSRF保護機能の実装
    - 外部フォーム送信用のCSRFトークン生成・検証機能を実装
    - フォーム送信時のCSRFトークン検証機能を実装
    - PropertiesService.getScriptProperties()を使用したサーバーサイド専用トークン管理機能を実装
    - トークンTTL（1時間）設定と自動期限切れ機能を実装
    - 期限切れトークンの自動クリーンアップ機能を実装
  - 重複送信防止機能の実装
    - フォーム送信時のユニークキー生成機能を実装
    - 重複送信検出機能（同一申請者の同一日付・同一申請種別の重複チェック）を実装
    - 重複送信時のエラーハンドリングとユーザー通知機能を実装
    - 申請IDの自動生成と一意性保証機能を実装
  - _要件: 2.1, 2.2, 2.4, 7.1, 7.2_

- [x] 8. メール管理・通知システムの実装
  - MailManager.gsファイルを作成
  - バッチメール送信機能（クォータ節約）を実装
  - 承認依頼メールのバッチ処理機能を実装
  - エラー通知メール機能を実装
  - 未退勤者通知メール機能を実装
  - _要件: 2.3, 4.2, 6.2_

- [x] 9. 自動処理トリガーの実装
  - Triggers.gsファイルを作成
  - dailyJob関数（02:00実行）を実装してDaily_Summary更新機能を作成
  - weeklyOvertimeJob関数（月曜07:00実行）を実装して残業警告機能を作成
  - monthlyJob関数（1日02:30実行）を実装してMonthly_Summary更新とCSVエクスポート機能を作成
  - onOpen関数で管理メニューを追加
  - _要件: 4.1, 4.3, 4.4_

- [x] 10. エラーハンドリングシステムの実装
  - 全ての主要関数にtry-catch文を実装
  - エラーログ記録機能を実装
  - 管理者への即時エラー通知機能を実装
  - クォータ制限対応のバッチ処理機能を実装
  - 実行時間制限対応のチャンク処理機能を実装
  - _要件: 6.1, 6.3, 6.4, 6.5_

- [x] 11. 承認処理とステータス更新の実装
  - 承認者によるステータス変更の検出機能を実装
  - ステータス変更時の自動処理（再計算トリガー）を実装
  - 申請者への処理完了通知機能を実装
  - 承認履歴の記録機能を実装
  - _要件: 3.3, 3.4, 3.5_

- [x] 12. レポート生成・エクスポート機能の実装
  - 日次サマリーレポート生成機能を実装
  - 月次サマリーレポート生成機能を実装
  - CSVエクスポート機能（Google Driveへの保存）を実装
  - 残業時間監視・アラート機能を実装
  - レポートへの承認ステータス表示機能を実装
  - _要件: 8.1, 8.2, 8.3, 8.4, 8.5_

- [x] 13. テスト機能の実装
  - Test.gsファイルを作成
  - 各モジュールの単体テスト関数を実装
  - テストカバレッジ表示機能を実装
  - データ整合性チェック機能を実装
  - システム稼働状況チェック機能を実装
  - _要件: 6.1_

- [x] 14. 時間トリガーの設定と最終統合
  - Google Apps Scriptで時間ベースのトリガーを設定
  - dailyJob, weeklyOvertimeJob, monthlyJobのスケジュール設定
  - フォームトリガーの設定
  - 全機能の統合テストを実行
  - エラーハンドリングの動作確認
  - _要件: 4.1, 4.3, 4.4, 4.5_

- [ ] 15. セキュリティ設定と権限管理の最終確認
  - WebAppの公開設定（「Googleアカウントを持つ全員」）を設定
  - 各シートの保護設定を最終確認
  - Request_Responsesシートの承認者権限設定を確認
  - スクリプト実行権限の設定を確認
  - セキュリティテストを実行
  - _要件: 7.1, 7.2, 7.3, 7.4, 7.5_

- [ ] 16. 包括的セキュリティ監査の実施
  - OAuthスコープの最小化確認
    - spreadsheets.readonlyスコープの使用可能性を検証
    - 必要最小限のスコープのみを要求するよう設定を最適化
    - 不要な権限の削除と権限の段階的付与の実装
  - アドオン・外部API呼び出しの検証
    - 使用中のアドオンの一覧作成と必要性の確認
    - 外部API呼び出しの有無とセキュリティリスクの評価
    - 不要な外部依存関係の特定と削除
  - STRIDE脅威モデリングの実行
    - Spoofing（なりすまし）対策の確認
    - Tampering（改ざん）対策の確認
    - Repudiation（否認）対策の確認
    - Information Disclosure（情報漏洩）対策の確認
    - Denial of Service（サービス拒否）対策の確認
    - Elevation of Privilege（権限昇格）対策の確認
  - セキュリティ監査レポートの作成
    - 発見された脆弱性の文書化
    - リスク評価と優先度の設定
    - 是正措置の提案と実装計画
  - _要件: 7.1, 7.2, 7.3, 7.4, 7.5_

---

# データ整合性保証の実装ガイドライン（GAS制約対応版）

## 概要
Google Apps Script（GAS）はACIDトランザクションをサポートしていないため、従来のトランザクション的処理の代わりに、以下の代替戦略を採用します。

## 実装戦略

### 1. バックアップ行保持による疑似ロールバック
- **目的**: 書き込み失敗時のデータ復旧
- **実装方法**:
  - 書き込み前に既存データをバックアップ行にコピー
  - 書き込み処理の成功/失敗を記録
  - 失敗時はバックアップ行から復旧処理を実行
- **適用箇所**: Log_Raw、Request_Responses、Daily_Summary、Monthly_Summaryシート

### 2. バージョン番号/タイムスタンプチェック
- **目的**: 重複書き込みと競合状態の防止
- **実装方法**:
  - 各レコードにバージョン番号またはタイムスタンプを追加
  - 書き込み時にバージョン番号をチェック
  - 不一致時は書き込みを拒否し、エラーログを記録
- **適用箇所**: 全シートの更新処理

### 3. エラーログ記録と管理者通知
- **目的**: 問題の早期発見と対応
- **実装方法**:
  - 専用のエラーログシートを作成
  - エラー発生時に詳細情報を記録
  - 管理者への即時メール通知
- **記録項目**: エラー時刻、エラー内容、影響範囲、復旧状況

### 4. 手動復旧プロセス
- **目的**: 自動復旧不可能な場合の対応
- **実装方法**:
  - 復旧用の管理画面を提供
  - データ整合性チェック機能
  - 手動復旧操作の実行機能
- **機能**: バックアップからの復旧、データ整合性検証、エラー修正

## 技術的制約と対策

### GASの制約
- ACIDトランザクション非対応
- 同時実行制限（6分間の実行時間制限）
- スプレッドシートの同時編集制限

### 対策
- LockServiceによる同時実行制御
- バッチ処理による効率化
- エラーハンドリングの強化
- 定期的なデータ整合性チェック

## 実装優先度
1. エラーログ記録システム
2. バックアップ行保持機能
3. バージョン番号チェック機能
4. 手動復旧プロセス
5. 管理者通知機能

---

# CSRFトークン管理の実装ガイドライン（PropertiesService対応版）

## 概要
従来のGoogle SheetでのCSRFトークン管理から、`PropertiesService.getScriptProperties()`を使用したサーバーサイド専用ストレージに変更し、セキュリティリスクを軽減します。

## 設計変更の理由

### 従来の課題
- **セキュリティリスク**: Google Sheetの広いアクセス権限によるトークン漏洩リスク
- **パフォーマンス**: スプレッドシートアクセスによる処理遅延
- **管理複雑性**: シートの作成・削除・クリーンアップ処理の複雑さ

### 新しい設計の利点
- **セキュリティ向上**: スクリプト専用のサーバーサイドストレージ
- **パフォーマンス向上**: 高速なキー・バリューアクセス
- **管理簡素化**: 自動的なTTL管理とクリーンアップ

## 実装仕様

### 1. トークンストレージ設計
- **ストレージ**: `PropertiesService.getScriptProperties()`
- **キー形式**: `csrf_token_{userEmail}_{timestamp}`
- **値形式**: `{token}_{expiryTimestamp}`
- **TTL**: 1時間（3600秒）

### 2. 主要機能

#### トークン生成・保存
```javascript
function saveCsrfToken(userEmail, token, expiryMinutes = 60) {
  const props = PropertiesService.getScriptProperties();
  const key = `csrf_token_${userEmail}_${Date.now()}`;
  const expiryTime = Date.now() + (expiryMinutes * 60 * 1000);
  const value = `${token}_${expiryTime}`;
  props.setProperty(key, value);
}
```

#### トークン検証
```javascript
function validateCsrfToken(userEmail, token) {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();
  const userKeys = keys.filter(key => key.startsWith(`csrf_token_${userEmail}_`));
  
  for (const key of userKeys) {
    const value = props.getProperty(key);
    const [storedToken, expiryTime] = value.split('_');
    
    if (Date.now() > parseInt(expiryTime)) {
      props.deleteProperty(key); // 期限切れトークンを削除
      continue;
    }
    
    if (storedToken === token) {
      props.deleteProperty(key); // 使用済みトークンを削除
      return true;
    }
  }
  
  return false;
}
```

#### 自動クリーンアップ
```javascript
function cleanupExpiredCsrfTokens() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();
  const csrfKeys = keys.filter(key => key.startsWith('csrf_token_'));
  
  for (const key of csrfKeys) {
    const value = props.getProperty(key);
    const expiryTime = parseInt(value.split('_')[1]);
    
    if (Date.now() > expiryTime) {
      props.deleteProperty(key);
    }
  }
}
```

### 3. テストケース設計

#### 基本機能テスト
- トークン生成・保存の正常動作
- 有効期限内でのトークン検証成功
- 期限切れトークンの検証失敗
- 無効トークンの検証失敗

#### TTL機能テスト
- 1時間後の自動期限切れ
- 期限切れトークンの自動削除
- 複数トークンの同時管理
- 大量トークン処理時のパフォーマンス

#### セキュリティテスト
- 異なるユーザー間のトークン分離
- トークンの再利用防止
- 無効なトークン形式の処理
- プロパティストレージの容量制限対応

### 4. 移行計画

#### フェーズ1: 新実装の準備
- PropertiesService版の関数実装
- テストケースの作成と実行
- パフォーマンステスト

#### フェーズ2: 並行運用
- 新旧両方のシステムを並行運用
- 段階的なトラフィック移行
- エラー監視とロールバック準備

#### フェーズ3: 完全移行
- 旧システムの無効化
- CSRF_Tokensシートの削除
- 最終動作確認

### 5. 制約と対策

#### PropertiesServiceの制約
- **値サイズ制限**: 9KB
- **キー数制限**: 500個
- **文字エンコーディング**: UTF-8

#### 対策
- トークン値の最小化（UUID使用）
- 定期的なクリーンアップ処理
- エラー時のフォールバック処理

## 実装優先度
1. PropertiesService版の基本機能実装
2. TTL機能と自動クリーンアップ実装
3. テストケースの作成と実行
4. 既存システムとの並行運用
5. 完全移行と旧システム削除