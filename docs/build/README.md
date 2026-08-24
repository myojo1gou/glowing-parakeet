# ドキュメント生成

`build-purchase-justification.js` は `docs/物品購入理由書_UGREEN_DXP4800Plus.docx` を生成するスクリプトである。
本文を修正した場合はスクリプトを編集して再生成する。

```bash
npm install docx
node build-purchase-justification.js ../物品購入理由書_UGREEN_DXP4800Plus.docx
```

本文は `ＭＳ 明朝`（本文）と `ＭＳ ゴシック`（見出し・表頭）を指定している。
これらのフォントが無い環境で開いた場合は代替フォントで表示される。
