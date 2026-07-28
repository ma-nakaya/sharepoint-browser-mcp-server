# Phase 8: SharePointリスト項目の本文取得

## 背景

SharePoint検索では、現行の規程や社内情報が文書ファイルではなく、次のようなリスト項目として返る場合があります。

```text
https://<tenant>.sharepoint.com/<site>/Lists/<list>/DispForm.aspx?ID=204
```

Phase 7の`search`はこのURLを返せましたが、`fetch`は`SitePages`だけを取得しようとしていたため、リスト項目では`INVALID_ARGUMENT`になっていました。

## 動作

`fetch`はURLの種類を次の順で判定します。

1. `Lists/<リスト名>/DispForm.aspx?ID=...`ならリスト項目
2. `SitePages/*.aspx`ならモダンページ
3. PDF、DOCX、XLSX、PPTXなら文書抽出
4. それ以外は拒否

リスト項目はSharePoint RESTの`FieldValuesAsText`を読み取り、次を行います。

- リッチテキストをプレーンテキストへ変換
- SharePoint内部列名のエンコードを復号
- バージョン、GUID、Taxonomy補助列などのシステム項目を除外
- タイトル、業務フィールド、変更日時、リストURL、項目IDを返却
- 本文を最大50,000文字に制限

## URL境界

取得対象は次の条件をすべて満たす必要があります。

- HTTPS
- 設定済みSharePointと同一オリジン
- 設定済みサイトパス内
- `Lists/<リスト名>/DispForm.aspx`の直下形式
- `ID`が重複しない正の整数

`Source`やフラグメントは引用用の正規URLから除外します。別サイト、別ホスト、任意のRESTパス、リスト配下の追加階層は拒否します。

## 検索の使い方

規程や規則は文書ファイルとは限らないため、`documents`スコープから始めません。標準の`search`へ正確な名称を渡し、返されたIDをそのまま`fetch`します。

```text
search("05 国内出張旅費規程")
fetch("https://.../Lists/.../DispForm.aspx?ID=204")
```

`search`と`sharepoint_search`の説明にもこの優先手順を記載しています。

## 参考資料

- [Working with lists and list items with REST](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest)
- [ListItem.FieldValuesAsText](https://learn.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.client.listitem.fieldvaluesastext?view=sharepoint-csom)
