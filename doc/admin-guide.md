# 出勤管理システム 管理者ガイド

## 1. はじめに
本ドキュメントは、GoogleスプレッドシートとGoogle Apps Script（GAS）を用いた出勤管理システムの管理者向け運用・設定・保守方法をまとめたものです。

## 2. システム構成
- Googleスプレッドシート（データ管理・UI）
- Google Apps Script（業務ロジック・自動化）
- 主要スクリプトファイル（src/配下）

## 3. 初期セットアップ手順（共有コピー方式）

1. **管理者から共有された「出勤管理システム」用Googleスプレッドシートを開きます。**
2. スプレッドシートの「ファイル」→「コピーを作成」を選択します。
   - この操作で、**スプレッドシート本体と紐付いたGAS（Apps Script）プロジェクトも同時にコピーされます。**
3. コピー後のスプレッドシートで「拡張機能」→「Apps Script」を開き、スクリプトが正しく複製されていることを確認します。
4. 必要に応じて、**「SYSTEM_CONFIG」シートの「MANAGER_EMAILS」行（または「ADMIN_EMAILS」行）や他の設定値を自組織用に編集してください。コードの直接編集やハードコーディングは不要です。**
5. スクリプトの初回実行時に権限付与ダイアログが表示されるので、指示に従い許可します。
6. スプレッドシートの「共有」から、従業員や他管理者のメールアドレスを「閲覧者」または「編集者」として追加します。

## 4. 管理者向け主要機能

- **ユーザー管理**
  - 「MASTER_EMPLOYEE」シートで従業員情報（氏名・メールアドレス・所属・雇用区分・上司Gmail等）を編集できます。
  - 新規従業員の追加や、退職者の削除・編集もこのシートで行います。
  - 上司Gmail欄が空欄の従業員は「管理者」として扱われます。

- **権限管理**
  - 「SYSTEM_CONFIG」シートの「MANAGER_EMAILS」または「ADMIN_EMAILS」行で管理者メールアドレスを設定できます。
  - 管理者は全データの閲覧・編集、集計・レポート出力、設定変更が可能です。
  - 一般ユーザーは自身の打刻・申請のみ可能です。

- **システム設定**
  - 「SYSTEM_CONFIG」シートで各種システム設定（セッションタイムアウト、メール送信モード、残業閾値など）を編集できます。
  - 設定変更後は、必要に応じて再読み込みや再実行を行ってください。

- **メール通知設定**
  - 未退勤者通知や月次レポートの送信先は「MANAGER_EMAILS」または「ADMIN_EMAILS」で管理します。
  - メール送信モード（テスト/本番）は「EMAIL_MOCK_ENABLED」「EMAIL_ACTUAL_SEND」で切り替え可能です。

- **Googleフォーム連携設定**
  - 「FormManager」機能により、Googleフォームからの応答を自動で取り込み・保存できます。
  - フォームの設計変更時は、対応するスクリプトの項目名・バリデーションも確認してください。


## 5. 日常運用ガイド

- **出勤・退勤打刻の確認**
  - 「Log_Raw」シートに全従業員の打刻データが記録されます。
  - 管理者はこのシートで打刻状況を確認できます。

- **集計・レポート出力**
  - 日次・週次・月次の集計は自動トリガーまたは手動実行で行われ、「Daily_Summary」「Weekly_Overtime」「Monthly_Report」等のシートに出力されます。
  - 月次レポートはCSV形式でエクスポートも可能です。

- **未退勤者チェック・メール送信**
  - トリガー実行時に未退勤者リストが自動生成され、管理者宛にメール通知されます。
  - 通知先は「SYSTEM_CONFIG」シートの管理者メール設定に従います。

- **フォーム応答の確認・一括処理**
  - Googleフォームからの申請は「FormManager」経由で自動保存されます。
  - 一括処理やエラー時の再処理も可能です。

- **バッチ処理（トリガー）の手動実行**
  - 「拡張機能」→「Apps Script」→「関数を選択」から、`dailyJob`や`monthlyJob`等のバッチ処理を手動実行できます。
  - 定期実行トリガーの設定・確認も管理者が行えます。 

## 6. テストと品質管理

- **テスト実行方法**
  - 「拡張機能」→「Apps Script」→「関数を選択」から、`runAllTests`や各TestRunner（例：`WebAppTestRunner`、`FormTestRunner`）を実行できます。
  - テスト結果はログまたはスプレッドシート上で確認できます。
  - テストカバレッジや実行時間も自動で記録されます。

- **テストケース例と期待動作**
  - 主要な業務ロジック・認証・WebAPI・トリガー・フォーム連携・メール送信など、各機能ごとにテストケースが用意されています。
  - テストケースはTDD方式で設計されており、失敗時は詳細なエラーメッセージが表示されます。

- **テストカバレッジ確認方法**
  - Test.gsの`showTestCoverage()`関数で、各モジュールのテスト済み関数・未テスト関数を一覧表示できます。
  - カバレッジレポートは開発・保守時の品質管理に活用してください。

- **テスト失敗時の対応**
  - エラーメッセージやログを確認し、該当するシート・設定・スクリプトを見直してください。
  - テストデータの初期化や再投入も有効です。


## 7. トラブルシューティング

- **よくあるエラーと対処法**
  - 「権限がありません」→ スクリプトの権限付与を再実行、Googleアカウントの切り替えを確認
  - 「管理者メールアドレスが未設定」→ SYSTEM_CONFIGシートの「MANAGER_EMAILS」または「ADMIN_EMAILS」を確認・修正
  - 「メール送信に失敗」→ メール送信モードやクォータ設定、Gmailの利用制限を確認
  - 「データが反映されない」→ シート名・列名・設定値のスペルミスやシート構成を確認

- **クォータ制限への対応**
  - GASの無料枠制限（実行時間・メール送信数等）に達した場合は、翌日まで待つか、処理件数を減らしてください。

- **権限・認証エラー**
  - Googleアカウントの権限設定や、スプレッドシートの共有範囲を見直してください。

- **データ不整合時の対応**
  - シートのバックアップを取得し、必要に応じて手動修正や再投入を行ってください。

- **ログ・デバッグ方法**
  - Apps Scriptの「表示」→「ログ」や、スプレッドシートのエラーログシートを活用してください。 