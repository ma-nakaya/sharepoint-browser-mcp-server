# Phase 5: SharePointサイト横断検索

## 目的

SharePoint自身の検索インデックスをサイト横断検索の入口として使い、候補ページ・候補文書を安全に絞り込んでから、ページ本文取得または文書本文抽出へ接続する。

```text
SharePoint検索インデックス
          ↓
サイト・フォルダー・種類・日付で候補を絞る
          ↓
ページ本文取得 / PDF・Office文書抽出
          ↓
必要な内容だけをエージェントが読む
```

文書単体の構造化検索だけでなく、設定済みSharePointサイト全体の検索品質と探索効率を改善対象とする。

## `sharepoint_search`の拡張

既存の`query`と`maxResults`に加え、次の任意入力へ対応する。

| 入力 | 内容 |
| --- | --- |
| `siteUrl` | 複数サイト構成時に検索対象を1つの設定サイトへ限定 |
| `startRow` | 検索結果の開始位置。前回結果の`nextStartRow`を使用できる |
| `scope` | `all`、`pages`、`documents`から対象種類を選択 |
| `folderUrl` | 設定サイト内のフォルダー配下に検索範囲を限定 |
| `fileExtensions` | PDF、DOCXなどの拡張子で限定。最大10種類 |
| `modifiedAfter` | 指定日以降に更新された項目へ限定 |
| `modifiedBefore` | 指定日以前に更新された項目へ限定 |
| `sort` | SharePointの関連度順または更新日時の降順 |

検索結果には、従来のタイトル、URL、更新日時、要約に加え、次の探索情報を含める。

- `kind`: `page`、`document`、`file`、`other`
- `parentUrl`: SharePointが返した親URLをサイト境界の検証後に返却
- `author`、`sizeBytes`、`rank`: SharePoint検索インデックスに値がある場合だけ返却
- `startRow`、`hasMore`、`nextStartRow`: 続きの検索に必要なページング情報
- 実際に適用したスコープ、拡張子、更新日、並び順

## SharePoint検索クエリ

SharePoint Search REST APIのGETエンドポイントを使用する。検索語自体は`querytext`へ渡し、サーバーが管理する制約を`querytemplate`へ追加する。

- 設定サイトまたは検証済みフォルダーの`Path`制約
- ページ検索では`FileType:aspx`
- 文書検索では`IsDocument:True`かつASPXを除外
- 拡張子は`FileType`のOR条件
- 更新日は`LastModifiedTime`の範囲条件
- 更新日時順では`sortlist=LastModifiedTime:descending`
- `rowlimit`、`rowsperpage`、`startrow`によるページング
- 語形変化を有効化し、重複結果を除外

## PageIndex型検索との関係

このフェーズは、将来の文書構造ツリー検索に対するサイト側の第1段階である。

1. SharePoint自身の権限・検索インデックスを使って候補ファイルを探す。
2. `kind=page`なら`sharepoint_get_page`で本文を取得する。
3. `kind=document`で対応形式なら`sharepoint_extract_document_text`へ渡す。
4. 今後、文書内の章・ページ・スライド・シート単位で構造を辿る。

サイト全体を外部サービスへ複製したり、ローカルへ無断で永続インデックスを作成したりしない。外部LLMによる要約やツリー生成もこのフェーズには含めない。

## セキュリティ境界

- 検索スコープは設定済みサイトまたはその配下のフォルダーだけ。
- `folderUrl`はHTTPS、オリジン、サイトパス、資格情報、パス表現を検証する。
- 拡張子は英数字だけを許可し、KQL構文の注入を防ぐ。
- 更新日は実在する`YYYY-MM-DD`だけを許可する。
- 検索結果URLと親URLを応答解析時にもサイト境界と照合する。
- SharePointへの要求はGETだけを使用する。
- Cookie、トークン、Authorizationヘッダー、検索応答本文をログへ記録しない。

## 制限事項

- 複数サイトを同時検索する場合、`startRow`は0だけを許可する。ページング時は`siteUrl`または`folderUrl`で1サイトへ限定する。
- SharePoint検索インデックスへ未反映の更新は検索できない。
- SharePointの検索スキーマ、権限、言語処理、ランキング設定に結果が依存する。
- `author`、`sizeBytes`、`rank`、`parentUrl`は検索スキーマに値がない場合は省略される。
- サイト全体の意味要約、ベクトル検索、PageIndexツリーの永続生成は未対応。
- 文書内の構造化検索はPhase 6で対応した。詳細は[`phase-6-document-structure.md`](phase-6-document-structure.md)を参照。
