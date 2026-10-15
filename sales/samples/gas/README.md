# 出品B用サンプル: フォーム回答の集計とメール通知(GAS)
すべてダミーデータです。実在の氏名・メールアドレスは含みません。

## ファイル
- `Code.gs`: Google Apps Script。集計(`aggregateOrders_`)、集計シートへの書き出し(`updateSummary`)、メール通知(`sendSummaryMail`)、トリガー設定(`setupTriggers`)。
- `sample-form-summary.xlsx`: ダミーの回答36行と、集計イメージ(SUMIFS)。
- `fixture.json`: テスト用データと期待値(Pythonで計算)。
- 画像: `../images/gas-sample-summary.png`(集計イメージ)、`../images/thumbnail-B.png`(サムネイル)。

## 検証(2026-10-16実施)
- `TZ=Asia/Tokyo node ../tools/test_gas.js` → PASS(集計9グループ、不正行3件を除外、空データ・数量0・単価負・月またぎも確認)。
- LibreOfficeでxlsxを再計算し、月次集計の9行と合計(78個、251,100円)がPythonの期待値と一致。
- **未検証**: Google Apps Script上での実行(フォーム連携、MailApp、トリガー)。ここの環境ではGoogleアカウントに接続できないため、ロジック部分(集計)のみNodeで検証した。納品前は必ず実環境で動作確認すること。

## 使い方(出品者が自分のGoogleアカウントで試す)
1. Googleフォームを作り、「日時・氏名・商品・数量・単価」の列順で回答先スプレッドシートを作る(回答シート名は `フォーム回答`)。
2. 拡張機能 → Apps Script に `Code.gs` を貼る。`CONFIG.NOTIFY_TO` を自分のメールアドレスにする。
3. `updateSummary` を実行して権限を承認し、「月次集計」シートができることを確認する。
4. `setupTriggers` を1回実行する(フォーム送信時の更新と、毎月1日9時のメール)。
5. 出品ページの画像には、自分の環境で動かした画面のスクリーンショットを使うと信頼されやすい(ダミーデータで)。
