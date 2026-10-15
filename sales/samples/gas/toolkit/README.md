# GASツールキット(出品B用の部品)
買い手の依頼133件の分析(ops/customer-needs-gas-2026-10-16.md)で多かった困りごと向けの部品。サンプル値(example@example.com、シート名)は納品先に合わせて書き換える。

| ファイル | 役割 | 主な関数(`_`付きはテスト済みの純粋関数) |
|---|---|---|
| Core.gs | 実行ログ、失敗時メール通知(同じ失敗は1時間に1通)、二重実行防止 | `runSafely_` `logEvent_` `shouldNotify_` `rowsToObjects_` |
| MailToSheet.gs | 予約/キャンセル等のメールを解析して一覧へ。二重処理防止、キャンセル反映、要確認 | `parseMail_` `applyMailEvent_` `filterNewIds_` `processMails` |
| FormRouter.gs | フォーム回答を日程別シートへ振り分け、表記ゆれ吸収、アンケート照合 | `dateSheetName_` `routeByDate_` `markSurveyDone_` `routeFormResponses` |
| CsvReconcile.gs | CSVの列名変更に強い読み取り、予定と実績の突き合わせ | `parseCsv_` `mapHeaders_` `reconcile_` `csvToRecords_` |
| Diagnose.gs | 必要なシート/トリガー/最後の成功/失敗/メール残数を診断シートに出力 | `evaluateDiagnosis_` `diagnose` |

## 検証状況(2026-10-16)
- **検証済み**: 純粋関数(`_`付き)を `TZ=Asia/Tokyo node ../tools/test_toolkit.js` で検証(16テスト・PASS)。コードを意図的に壊すとテストが失敗することも確認。
- **未検証**: Apps Script上での実行。`SpreadsheetApp` `GmailApp` `MailApp` `LockService` `PropertiesService` `ScriptApp` を使う部分(`processMails` `routeFormResponses` `diagnose` `runSafely_` など)、フォーム連携、トリガー、メール通知、権限承認の画面。ここの環境ではGoogleアカウントに接続できないため。
- 納品前は必ず `sales/gas-delivery-kit/delivery-check.md` で、自分のGoogleアカウント・ダミーデータで通して確認する。
- メールの正規表現(`MAIL_CFG.FIELDS`)は、実際のメール文面に合わせて調整が必要。

## 使い方(出品者が自分で試す)
1. 新しいスプレッドシートを作り、拡張機能→Apps Script を開く。
2. ファイルを5つ(Core / MailToSheet / FormRouter / CsvReconcile / Diagnose)作り、内容を貼る。
3. `TK.NOTIFY_TO` を自分のメールアドレスにする。
4. 「フォーム回答」シートを作り(1行目に見出し、3列目に「参加希望日」)、ダミー回答を数行入れて `routeFormResponses` を実行。日程別シートと「要確認」シートができ、「実行ログ」に成功が残ることを確認。
5. わざとシート名を変えて失敗させ、通知メールが届くか、`diagnose` が「要対応」を出すかを確認。
