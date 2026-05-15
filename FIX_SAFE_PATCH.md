# 銀革屋 在庫管理：安全修正パッチ

Codex差分で残っていた危険箇所を直すための置き換え指示です。

## 修正1：Drive APIエラー文を1行にする

`ocrPdfByDrive_` の中にある次の部分を探します。

```js
if (typeof Drive === 'undefined') {
  throw new Error('Drive APIが有効になっていません。Apps Script左側の「サービス」から Drive API を追加してください。');
  throw new Error('Apps Script左側の「サービス」から Drive API を追加してください。');
}
```

これを下の内容に置き換えます。

```js
if (typeof Drive === 'undefined') {
  throw new Error('Drive APIが有効になっていません。Apps Script左側の「サービス」→「＋」→「Drive API」を追加してから、もう一度実行してください。');
}
```

## 修正2：PDF読取候補を勝手に反映OKにしない

`copyConfirmedPdfCandidatesToArrivalPlan` の `arrivalSheet.appendRow` の中で、次の並びを探します。

```js
'PDF読取',
sourceText || '',
'反映OK',
'未反映',
'',
'革',
'PDF確認済み'
```

これを下の内容に置き換えます。

```js
'PDF読取',
sourceText || '',
'未反映',
'',
'革',
'PDF確認済み'
```

## 理由

PDF読取候補は、まだ候補です。  
確認済みであっても、すぐに在庫反映OKにすると危険です。

正しい流れは以下です。

1. PDF読取候補を作成
2. 内容を人が確認
3. 入荷予定へコピー
4. 入荷予定では「未反映」のまま
5. 実物確認後に「反映OK」にする
6. 在庫へ反映

## Codexへの依頼文

```text
ginkawaya/base の src/Code.gs に対して、FIX_SAFE_PATCH.md の2点をそのまま反映してください。
特に copyConfirmedPdfCandidatesToArrivalPlan では、PDF読取候補から入荷予定へコピーする時点で、P列の在庫反映は必ず「未反映」にしてください。
「反映OK」は入れないでください。
修正後に差分を出してください。
```
