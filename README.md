# 銀革屋 革在庫管理 Apps Script

## 目的

Gmailの「革仕入」ラベル付きメールから、革・資材の入荷予定を整理し、スプレッドシートで在庫管理するためのコードです。

## 基本方針

- 非エンジニアでも使いやすい運用にする
- 在庫数は勝手に増やさない
- 革の在庫は「枚数」ではなく「ds数」で管理する
- 画像PDFから読んだ内容は候補として出し、必ず確認してから反映する
- 1マス＝1行表示を前提にする

## Google側で必要な作業

1. スプレッドシートを開く
2. 上メニュー「拡張機能」→「Apps Script」
3. `src/Code.gs` の中身を全部コピーして貼り付ける
4. 保存する
5. 左側の「サービス」から Drive API を追加する
6. 最初に `testConnection` を実行する

## よく使う実行名

- `runLeatherInventoryAutomation`  
  メール取込・整理を1回実行します。

- `createFullAutomationTrigger`  
  1時間ごとの自動実行を設定します。

- `reflectConfirmedArrivals`  
  入荷予定で「確認済み」かつ「反映OK」のものを在庫マスターへ反映します。

- `sendMonthlyInventoryPdfTest`  
  棚卸PDFメールをテスト送信します。

- `createMonthlyInventoryPdfTrigger`  
  毎月1日朝9時ごろに棚卸PDFメールを自動送信します。

- `importWinsFactoryPdfCandidates`  
  ウインズファクトリーの画像PDF請求書を読み取り、候補を「PDF読取候補」に入れます。

- `copyConfirmedPdfCandidatesToArrivalPlan`  
  PDF読取候補で「確認済み」にしたものを入荷予定へコピーします。

## 安全設定

初期状態では、在庫マスターへの自動反映はOFFです。

```js
const AUTO_REFLECT_STOCK = false;
```

このままなら、在庫数が勝手に増えることはありません。

## 注意

PDF画像読み取りは100%正確ではありません。  
特に、ds数・金額・単価は必ず人が確認してから反映してください。
