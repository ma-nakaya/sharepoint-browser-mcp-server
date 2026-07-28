# Phase 2: SharePoint検索とページ本文取得

## 目的

Phase 1で確立したMCP専用Edgeの認証済みセッションを使い、設定済みSharePointサイトを読み取り専用で検索し、SharePointページの作成済み本文を取得する。

## MCPツール

### `sharepoint_search`

SharePoint Search REST APIへGET要求を送り、設定サイト配下のコンテンツを検索する。

- 検索語は1〜200文字
- 返却件数は1〜20件、既定値は10件
- 検索クエリに設定サイトの`Path`制約を追加
- 応答解析時にもURLのHTTPSオリジンとサイトパスを再検証
- タイトル、URL、種類、更新日時、短い要約だけを返却

Phase 5で、フォルダー・ページ／文書・拡張子・更新日による絞り込み、更新日時順、ページング、親URLなどの探索情報を追加した。詳細は[`phase-5-site-search.md`](phase-5-site-search.md)を参照。

### `sharepoint_get_page`

SharePointのResourcePath対応REST APIを使い、設定サイトの`SitePages`ライブラリにある`.aspx`ページのリスト項目を取得する。

- 絶対URLとサーバー相対URLに対応
- HTTPS、SharePointオリジン、サイトパス、`SitePages`、`.aspx`をすべて検証
- `Title`、`FileRef`、`CanvasContent1`、`Modified`だけをAPIへ要求
- 応答の`FileRef`が要求したパスと一致することを再検証
- `CanvasContent1`からスクリプト、スタイル、タグを除去し、プレーンテキストだけを返却
- 本文は最大50,000文字とし、超過時は`truncated: true`を返却

## 処理の流れ

1. MCP入力を長さ・件数・URL境界で検証する。
2. 設定サイト配下の`/_api/` URLを組み立てる。
3. `BrowserContext.request`から認証済みGET要求を送る。
4. Cookie適用や社内プロキシとの差異で取得できない場合、同一SharePointページ内の`fetch`へフォールバックする。
5. HTTP状態、JSON Content-Type、応答サイズを検証する。
6. 検索結果URLまたはページの`FileRef`を設定サイト境界と照合する。
7. 必要最小限の構造化データだけをMCPへ返す。

## セキュリティ境界

- HTTPメソッドはGETだけを使用する。
- REST要求先は設定済みサイト配下の`/_api/`だけに制限する。
- 検索はクエリと応答解析の二段階でサイト境界を適用する。
- ページ取得は設定サイトの`SitePages`配下にある`.aspx`だけを許可する。
- HTTP応答は最大1 MiBまで読み取り、超過時は結果を返さずエラーにする。
- Cookie、トークン、Authorizationヘッダー、REST応答本文をログへ記録しない。
- ページの生HTMLとWebパーツ構成情報をMCP結果へ含めない。

## 制限事項

- 検索インデックスへ反映されていない更新内容は検索結果に現れない。
- 動的Webパーツが実行時に読み込むリスト項目や外部データは、ページ本文へ展開しない。
- ファイル一覧と制限付きダウンロードはPhase 3で対応する。Office文書解析は未対応。
- リスト項目の汎用取得とSharePointへの書き込みは未対応。
