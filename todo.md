# 出勤管理システム実装計画（TDD版）

## プロジェクト概要

**プロジェクト名**: 出勤管理システム（Google Workspace 非利用版）  
**開発手法**: Test-Driven Development (TDD) - Red-Green-Refactorサイクル  
**技術スタック**: Google Apps Script（GAS）+ Google スプレッドシート  
**対象規模**: 従業員50名以下  
**作成日**: 2025-07-10  

## 開発原則（t-wada準拠）

### TDDサイクル
1. **Red**: テストを先に書き、失敗させる（1-3分）
2. **Green**: テストを通すための最小限のコードを書く（2-5分）
3. **Refactor**: 動作を変えずにコードを改善する（1-3分）

### 開発ルール
- **テストファースト**: コードより先にテストを書く
- **小さなステップ**: 1つの機能を数行〜十数行で実装
- **5-10分サイクル**: 1回のRed-Green-Refactorを短時間で完了
- **F.I.R.S.T.原則**: Fast, Independent, Repeatable, Self-Validating, Timely
- **継続的リファクタリング**: 毎サイクルでコード品質を向上

## 制約条件

### 無料枠クォータ制限
- Apps Script実行時間：90分/日
- メール送信：100通/日
- URLFetch：20,000呼/日

### GAS固有の制約
- フォルダ構造は使用不可（フラット構造）
- ファイル名による機能分類
- claspによるローカル開発とGAS同期

## TDD実装計画

### 🔴 フェーズ1: 基盤構築（TDD基礎固め）

- [ ] **setup-1**: プロジェクト基本設定とテスト環境の構築
  - **TDDアプローチ**: まずテスト用フレームワークの導入と動作確認
  - GASUnitテストフレームワークの設定
  - 最初のダミーテスト作成（assert(true)から開始）
  - 見積時間: 2時間（テストファースト環境構築）
  - 依存関係: なし

- [ ] **constants-1**: Constants.gs の TDD実装
  - **Red**: 定数取得テストを先に作成
  - **Green**: 最小限の定数定義で通す
  - **Refactor**: 定数の整理と命名改善
  - テスト例: `testGetColumnIndex_Employee_Name_ReturnsCorrectIndex()`
  - 見積時間: 3時間（TDD小刻み実装）
  - 依存関係: setup-1

- [ ] **utils-1**: CommonUtils.gs の TDD実装
  - **アプローチ**: 1つの関数ずつTDDサイクルで実装
  - まずは簡単な関数（formatDate等）から開始
  - 各関数に対してRed-Green-Refactorを適用
  - テスト例: `testFormatDate_ValidDate_ReturnsFormattedString()`
  - 見積時間: 4時間（継続的リファクタリング含む）
  - 依存関係: constants-1

- [ ] **spreadsheet-1**: スプレッドシート操作の TDD実装
  - **Red**: シート作成・取得のテストを先に書く
  - **Green**: SpreadsheetApp.getActiveSpreadsheet()による最小実装
  - **Refactor**: エラーハンドリングと共通化
  - 段階的に6枚のシートを1枚ずつ追加
  - 見積時間: 4時間（小さなステップでの積み重ね）
  - 依存関係: setup-1

- [ ] **test-framework**: Test.gs - テストランナーの実装
  - **目的**: GAS独自のテスト実行環境構築
  - 各モジュールのテスト実行
  - テスト結果のログ出力
  - 見積時間: 3時間（テスト環境の充実）
  - 依存関係: utils-1

### 🔴 フェーズ2: コア機能のTDD実装

- [ ] **utils-2**: 業務ロジック関数の TDD実装
  - **アプローチ**: 最も重要な関数から優先実装
  - `isHoliday()`: 休日判定ロジック
    - Red: 様々な日付パターンのテスト
    - Green: 最小限の判定ロジック
    - Refactor: 効率的な判定アルゴリズム
  - `calcWorkTime()`: 勤務時間計算
    - Red: 正常・異常パターンのテスト
    - Green: 基本的な時間計算
    - Refactor: エッジケース対応
  - 見積時間: 6時間（小刻みTDDサイクル）
  - 依存関係: utils-1, constants-1

- [ ] **auth-1**: 認証機能の TDD実装
  - **Red**: Gmail認証のテストケース作成
  - **Green**: Session.getActiveUser().getEmail()による基本実装
  - **Refactor**: ホワイトリスト照合とセキュリティ強化
  - テスト駆動でのセキュリティ実装
  - 見積時間: 4時間（セキュリティTDD）
  - 依存関係: utils-1

- [ ] **webapp-1**: WebApp.gs の TDD実装
  - **アプローチ**: doGet/doPostを分離してTDD
  - doGet実装:
    - Red: GET リクエストのテスト
    - Green: 基本的なHTML返却
    - Refactor: テンプレート化
  - doPost実装:
    - Red: POST データ処理のテスト
    - Green: 最小限のデータ受信
    - Refactor: バリデーション追加
  - 見積時間: 6時間（WebAPI設計TDD）
  - 依存関係: utils-2, auth-1

### 🟡 フェーズ3: 集計機能のTDD実装

- [ ] **triggers-1**: トリガー関数の TDD実装
  - **Red**: 各種トリガーのテストケース
  - **Green**: 最小限のトリガー処理
  - **Refactor**: エラーハンドリングと最適化
  - onOpen(), dailyJob(), weeklyJob(), monthlyJob()を順次実装
  - 見積時間: 6時間（バッチ処理TDD）
  - 依存関係: utils-2

- [ ] **form-1**: Google フォーム連携の TDD実装
  - **Red**: フォーム応答処理のテスト
  - **Green**: 基本的な応答受信
  - **Refactor**: データ変換と保存処理
  - 見積時間: 4時間（外部連携TDD）
  - 依存関係: spreadsheet-1

- [ ] **mail-1**: メール機能の TDD実装
  - **Red**: メール送信のテスト（モック使用）
  - **Green**: GmailApp.sendEmailによる基本実装
  - **Refactor**: テンプレート化とクォータ管理
  - 見積時間: 5時間（外部依存TDD）
  - 依存関係: utils-1

### 🟢 フェーズ4: TDDテスト品質向上

- [ ] **coverage-1**: テストカバレッジの計測
  - **目的**: t-wada推奨のカバレッジ活用
  - 未テスト箇所の特定
  - カバレッジレポート生成
  - 見積時間: 3時間（品質可視化）
  - 依存関係: 全機能実装完了

- [ ] **test-enhancement**: テスト品質向上
  - **F.I.R.S.T.原則** の適用
    - Fast: テスト実行時間の最適化
    - Independent: テスト間の独立性確保
    - Repeatable: どの環境でも同じ結果
    - Self-Validating: 明確なpass/fail判定
    - Timely: 適切なタイミングでのテスト作成
  - テストの名前付け改善
  - 見積時間: 4時間（テスト品質向上）
  - 依存関係: coverage-1

- [ ] **integration-test**: 統合テストの TDD実装
  - **Red**: エンドツーエンドのテストシナリオ作成
  - **Green**: 実際のワークフロー実行
  - **Refactor**: テストデータとモックの整理
  - 見積時間: 6時間（結合テストTDD）
  - 依存関係: test-enhancement

### 🟢 フェーズ5: ドキュメント・デプロイ

- [ ] **doc-1**: 管理者向けドキュメント（/doc/admin-guide.md）
  - TDDで実装された機能の解説
  - テストケースを活用した動作説明
  - 見積時間: 3時間
  - 依存関係: integration-test

- [ ] **doc-2**: 従業員向け利用マニュアル（/doc/user-manual.md）
  - ユーザーストーリーベースの説明
  - FAQ（テストケースベース）
  - 見積時間: 2時間
  - 依存関係: integration-test

- [ ] **deploy-1**: Webアプリのデプロイ設定
  - 本番環境テストの実施
  - テスト付きデプロイメント
  - 見積時間: 2時間
  - 依存関係: doc-2

### ⚪ フェーズ6: オプション機能（TDD継続）

- [ ] **webhook-1**: Slack/LINE Webhook連携（TDD実装）
  - 外部API連携もTDDで実装
  - モックサーバーを活用
  - 見積時間: 4時間
  - 依存関係: mail-1

- [ ] **refactor-continuous**: 継続的リファクタリング
  - 全体のコード品質見直し
  - TDDサイクルで段階的改善
  - 見積時間: 4時間
  - 依存関係: deploy-1

## TDD実装スケジュール

### 週次スケジュール（TDD重視）

**第1週**: TDD基盤構築
- 月: setup-1（テスト環境）
- 火: constants-1（定数TDD）
- 水: utils-1（共通関数TDD）
- 木: spreadsheet-1（データアクセスTDD）
- 金: test-framework（テスト整備）

**第2週**: コア機能TDD
- 月: utils-2（業務ロジックTDD）
- 火: auth-1（認証TDD）
- 水-木: webapp-1（WebAPIのTDD）
- 金: リファクタリング・テスト見直し

**第3週**: 機能拡張TDD
- 月: triggers-1（バッチ処理TDD）
- 火: form-1（フォーム連携TDD）
- 水: mail-1（メール機能TDD）
- 木-金: coverage-1, test-enhancement

**第4週**: 統合・完成
- 月-火: integration-test（統合テストTDD）
- 水: doc-1, doc-2（ドキュメント）
- 木: deploy-1（デプロイ）
- 金: refactor-continuous（継続改善）

## TDD品質基準

### テスト品質（F.I.R.S.T.原則）
- **Fast**: 単体テスト1件 < 100ms、全体 < 30秒
- **Independent**: テスト間の依存関係なし
- **Repeatable**: どの環境でも同じ結果
- **Self-Validating**: 明確なpass/fail判定
- **Timely**: 実装前にテスト作成

### コード品質（TDD成果物）
- テストファーストによる設計駆動
- 小さなステップでの継続的リファクタリング
- 機能ごとの独立性確保
- Red-Green-Refactorサイクルでの品質担保

### カバレッジ活用（t-wada方式）
- 管理目標ではなく開発支援ツール
- 未テスト部分の可視化
- 自分のためのフィードバック機能
- 90%目標、100%は過度な追求なし

## リスク管理（TDD対応）

### 技術的リスク
1. **テストの複雑化**: 小さなステップで回避
2. **外部依存**: モック・スタブで制御
3. **GAS特有の制約**: 実際の環境でのテスト重視

### TDD実践リスク
1. **初期速度の低下**: 学習投資として受容
2. **テスト設計の困難**: ペアプログラミングで解決
3. **リファクタリング不足**: 毎サイクル必須実施

## 進捗管理（TDD版）

### ステータス定義
- [ ] 未着手
- [🔴] Red実装中（テスト作成中）
- [🟢] Green実装中（コード作成中）
- [♻️] Refactor実行中（改善中）
- [x] 完了（TDDサイクル完了）
- [!] 問題あり（サイクル中断）

### 更新ルール
1. 各TDDサイクル開始時にステータス更新
2. Red-Green-Refactor完了時に [x] に変更
3. 5-10分ルール違反時は [!] でサイクル見直し
4. 日次でカバレッジとテスト結果確認

---

**最終更新**: 2025-07-10  
**TDD手法**: t-wada（和田卓人）推奨方式  
**次回レビュー予定**: 第1週TDD基盤完了時 