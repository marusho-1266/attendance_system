# 出勤管理システム仕様書（Google Workspace 非利用版）改訂版

## 1. 想定環境

### 基本環境
- 各従業員は個人または共通 Gmail アカウントを保有
- Google Apps Script（無料枠）＋ Google スプレッドシート
- 対象従業員：～50名
- チャットボット／Directory API／組織ドメイン限定公開は利用不可

### 無料枠クォータ（目安）
- Apps Script 実行時間：90分／日
- メール送信：100通／日
- URLFetch：20,000呼／日

## 2. システム全体構成

### 入力
1. **打刻 Web アプリ**（スマホ／PC ブラウザ、QR コード起動）
2. **Google フォーム**（修正・残業・休暇申請）

### 保存
- Google スプレッドシート 6枚（後述）

### 処理
- **Apps Script**
  - 打刻受付
  - 日次／月次自動集計
  - クォータ節約型メール通知
  - **プルダウン選択式**の簡易承認ワークフロー
  - **堅牢なエラーハンドリング**

### 通知
- Gmail（まとめメール方式）
- Slack／LINE Webhook（任意、メール節約用）

## 3. 機能一覧

### A. 打刻
- IN（出勤）／OUT（退勤）／BRK_IN（休憩開始）／BRK_OUT（休憩終了）
- **`Session.getActiveUser().getEmail()` を利用し、現在ログイン中の Google アカウント情報で本人を直接認証**
- 同日二重打刻チェック／退勤未入力チェック

### B. 自動計算
- 実働時間、所定外（残業）時間、深夜時間、遅刻・早退
- 休憩自動補正（例：実働6h以上で45分強制控除 等）
- 法定休日・会社休日判定（Holiday マスタ）

### C. 承認ワークフロー
- 申請フォーム → ステータス列"Pending"
- トリガーで1通にまとめて承認者へメール
- **承認者がシート保護された範囲のプルダウンから"Approved/Rejected"を選択 → スクリプトで自動反映。入力ミスを防止**

### D. アラート／リマインド
- 前日退勤漏れ一覧メール（全員分まとめ）
- 月80hを超えそうな残業者リストメール（管理者向け週次）

### E. レポート
- Daily_Summary／Monthly_Summary シート
- CSV エクスポート（Drive に保存）

### F. 権限管理
- Raw ログ＝閲覧／編集とも禁止（シート保護）
- Summary シート＝全員「閲覧のみ」共有
- **承認申請シート＝承認者自身が担当する行の「ステータス列」のみ編集可能（プルダウン選択）とし、他はすべて保護**
- システム管理者のみスクリプト編集可

## 4. スプレッドシート構成

### ① Master_Employee
| 列 | 項目 |
|---|---|
| A | 社員ID |
| B | 氏名 |
| C | Gmail |
| D | 所属 |
| E | 雇用区分 |
| F | 上司 Gmail |
| G | 基準始業 |
| H | 基準終業 |

### ② Master_Holiday
| 列 | 項目 |
|---|---|
| A | 日付 |
| B | 名称 |
| C | 法定休日? |
| D | 会社休日? |

### ③ Log_Raw（編集禁止）
| 列 | 項目 |
|---|---|
| A | タイムスタンプ |
| B | 社員ID |
| C | 氏名 |
| D | Action |
| E | 端末IP |
| F | 備考 |

### ④ Daily_Summary
| 列 | 項目 |
|---|---|
| A | 日付 |
| B | 社員ID |
| C | 出勤 |
| D | 退勤 |
| E | 休憩 |
| F | 実働 |
| G | 残業 |
| H | 遅刻/早退 |
| I | 承認状態 |

### ⑤ Monthly_Summary
| 列 | 項目 |
|---|---|
| A | 年月 |
| B | 社員ID |
| C | 勤務日数 |
| D | 総労働 |
| E | 残業 |
| F | 有給 |
| G | 備考 |

### ⑥ Request_Responses（フォーム連携）
| 列 | 項目 |
|---|---|
| A | タイムスタンプ |
| B | 社員ID |
| C | 種別(修正/残業/休暇) |
| D | 詳細 |
| E | 希望値 |
| F | 承認者 |
| G | Status **(入力規則によるプルダウン化)** |

## 5. Apps Script モジュール

### WebApp.gs
#### doGet(e)/doPost(e)
- **`Session.getActiveUser().getEmail()`でログイン中のGoogleアカウントを取得 → マスタ照合**
- ログ追記 → 重複チェック
- 打刻後のレスポンス HTML

### Triggers.gs
#### onOpen()
- 管理メニュー追加

#### dailyJob()
- **実行時間：** 02:00
- **`try...catch` でエラーを捕捉し、失敗時は管理者に即時メール通知**
- 前日分を Daily_Summary 更新
- 未退勤者一覧メール（1通、To:本人 / Cc:管理）

#### weeklyOvertimeJob()
- **実行時間：** 毎週月曜07:00
- **`try...catch` でエラーを捕捉し、失敗時は管理者に即時メール通知**
- 直近4週の残業集計 → 管理者へ警告メール

#### monthlyJob()
- **実行時間：** 毎月1日 02:30
- **`try...catch` でエラーを捕捉し、失敗時は管理者に即時メール通知**
- Monthly_Summary 転記 → 前月シート保護
- CSV を Drive に出力 → 管理者へリンク1通

### FormTrigger.gs
#### onRequestSubmit(e)
- **`try...catch` でエラーを捕捉し、失敗時は管理者に即時メール通知**
- Status を "Pending" で初期化
- 同一承認者の申請を1日1通にまとめるバッファ配列 → cronMail()

### Utils.gs
- isHoliday(d) / calcWorkTime(in,out,breaks) / getEmployee(gmail) ほか

## 6. アクセス・セキュリティ設計

### Web アプリ公開設定
- **「Google アカウントを持つ全員」に設定し、スクリプト実行時に `Session.getActiveUser().getEmail()` で取得したアカウントが従業員マスタに存在するか検証することで、なりすましを防止**

### スクリプト所有者
- 組織共通 Gmail（退職等のリスク回避）

### シート共有
- **Master/Log_Raw：** 管理者のみ
- **Daily/Monthly_Summary：** 全社員「閲覧のみ」
- **Request_Responses:** **原則閲覧のみ。ただし「範囲の保護」機能で、承認者は担当行のステータス列のみ編集可とする**

## 7. クォータ対策ポリシー

### ① メール
- 個別リマインドは禁止。件名 "未退勤者まとめ (04/10)" の1通に集約
- 承認依頼は申請者/承認者単位で1日1通上限
- Slack/LINE Webhook 切り替え可（メール0通運用も可能）

### ② スクリプト実行時間
- `getValues` / `setValues` はブロック単位処理
- 集計ターゲットは「更新があった社員のみ」抽出で再計算

## 8. 代表的な業務フロー

### 出勤
1. 従業員 → スマホで QR 読取
2. Web アプリ → **ログイン中のGoogleアカウントで自動認証** → 「出勤」
3. "打刻完了" 画面

### 退勤漏れ
1. 02:00 バッチ → 未退勤者があれば「未退勤者まとめ」メール送信
2. 当人が翌朝フォームで修正

### 残業申請
1. フォーム送信 → `onRequestSubmit` で Pending
2. 当日17:30 承認者へまとめメール
3. **承認者がシートを開き、担当行のステータス列のプルダウンから "Approved" を選択** → 次回 `dailyJob` で再計算＆本人へ自動通知

### 月次
1. 毎月1日 02:30 → `Monthly_Summary` 完成
2. 管理者へ CSV リンクメール
3. 経理担当が給与システムにインポート

## 9. 導入・運用ステップ

| Step | 作業内容 | 所要時間 |
|------|----------|----------|
| Step0 | 共通 Gmail 取得・管理者権限付与 | - |
| Step1 | スプレッドシートひな形作成・**入力規則と保護設定** | 1時間 |
| Step2 | スクリプト初期導入（**エラーハンドリング含む**） | 1.5週間 |
| Step3 | 社員マスタ登録＆QR配布 | 半日 |
| Step4 | テスト運用（実データ） | 2週間 |
| Step5 | ルール微調整／クォータ監視 | - |
| Step6 | 本番運用開始・マニュアル配布 | - |
| **Step7** | **運用：ログの定期アーカイブ（年度ごと等）** | **年1回** |

## 10. 拡張余地

- Looker Studio で BI ダッシュボード（無料）
- CSV 自動アップロード → freee/マネフォークラウド給与 API 連携
- Firebase Hosting + GAS API で UX 高度化
- BigQuery 移行（50名超・数年分データ蓄積時）

## まとめ

本システムの特徴は以下の3点です：

1.  **認証は `Session.getActiveUser()` を利用し、ログイン中のGoogleアカウントで直接行うことでなりすましを防止**
2.  **メール／実行時間クォータを「まとめ送信」と「差分計算」で節約**
3.  **スクリプト所有者を共用 Gmail に固定し、堅牢なエラーハンドリングを組み込むことで運用リスクを最小化**

これにより、Google Workspace を利用せずとも、**より安全で、堅牢かつ効率的な**出勤管理システムを構築できます。

## GAS関数リスト自動抽出・カバレッジ半自動化運用例

### 1. 関数リスト自動抽出用Node.jsスクリプト例
```javascript
// extract-functions.js
const fs = require('fs');
const path = require('path');

const targetFiles = [
  'src/BusinessLogic.gs',
  'src/Authentication.gs',
  'src/WebApp.gs',
  'src/Triggers.gs',
  'src/FormManager.gs',
  'src/MailManager.gs',
];

targetFiles.forEach(file => {
  const content = fs.readFileSync(file, 'utf8');
  const matches = [...content.matchAll(/^function\s+([a-zA-Z0-9_]+)/gm)];
  const functions = matches.map(m => m);
  console.log(`${path.basename(file)}: [${functions.map(f => `'${f}'`).join(', ')}]`);
});
Use code with caution.
Md
2. テスト済み関数リスト抽出例```javascript
// extract-tested.js
const fs = require('fs');
const path = require('path');
const testFiles = [
'src/BusinessLogicTest.gs',
'src/AuthenticationTest.gs',
'src/WebAppTest.gs',
'src/TriggersTest.gs',
'src/FormManagerTest.gs',
'src/MailManagerTest.gs',
];
testFiles.forEach(file => {
const content = fs.readFileSync(file, 'utf8');
const matches = [...content.matchAll(/^function\s+test([A-Z][a-zA-Z0-9_])/gm)];
const tested = matches.map(m => m.replace(/_./, ''));
console.log(${path.basename(file)}: [${[...new Set(tested)].map(f =>'${f}').join(', ')}]);
});
Generated code
### 3. 運用手順
1. 上記スクリプトをローカルで実行し、各モジュールの関数リスト・テスト済みリストを抽出。
2. 抽出結果をTest.gsのshowTestCoverage()のmoduleInfoに貼り付け。
3. 必要に応じて手動で補正。
4. カバレッジレポートを実行し、未テスト関数を可視化。

---
この運用により、カバレッジ管理の半自動化・メンテナンス性向上が可能です。