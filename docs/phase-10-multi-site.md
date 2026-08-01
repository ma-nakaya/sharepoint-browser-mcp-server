# Phase 10: 複数SharePointサイトの読み取り

## 目的

1つのMCP接続と専用Edgeプロファイルを再利用し、明示的に許可した複数のSharePointサイトを横断検索・取得する。

## 設定

主サイトは従来どおり`SHAREPOINT_SITE_URL`へ指定し、追加サイトは`SHAREPOINT_ADDITIONAL_SITE_URLS`へカンマ、セミコロン、または改行区切りで指定する。

```dotenv
SHAREPOINT_SITE_URL=https://tenant.sharepoint.com/teams/main
SHAREPOINT_ADDITIONAL_SITE_URLS=https://tenant.sharepoint.com/teams/engineering
```

- 最大10サイト
- 全サイトが主サイトと同じSharePointテナントであること
- `/sites/<name>`または`/teams/<name>`のサイトコレクションであること
- 重複URLは正規化後に除外

## 動作

- `search`と`sharepoint_search`は既定で全サイトを並列検索し、URL重複を除いて関連度順に統合する。
- `sharepoint_search`の`siteUrl`または`folderUrl`で1サイトへ限定できる。
- ページ、リスト項目、フォルダー、ファイル、文書解析は入力URLのサイトパスから対応サービスへルーティングする。
- `sharepoint_list_document_libraries`は既定で全サイトを統合し、`siteUrl`指定時は1サイトだけを返す。
- `sharepoint_auth_status`は全サイトを確認し、サイト別結果も返す。
- `npm run login`は設定済みサイトをそれぞれタブで開く。

## セキュリティ境界

- 設定にないサイトURLは、同じテナント内でも拒否する。
- RESTエンドポイントはルーティング後も各サイト固有のURLガードで再検証する。
- 絶対URLとサーバー相対URLのオリジン・サイトパスを照合する。
- Cookie、トークン、Authorizationヘッダーは従来どおり取得・返却・記録しない。
- 全操作は読み取り専用のGETに限定する。
