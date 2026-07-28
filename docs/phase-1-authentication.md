# Phase 1: Edge認証セッションと認証状態確認

## 目的

Entra IDのアプリ登録を行わず、利用者がMicrosoft Edgeで認証したSharePointセッションをMCPから読み取り専用で利用できることを確認する。

Phase 1では検索やファイル取得に進まず、認証方式の成立性だけを検証する。

## 処理の流れ

1. `npm run login`でMCP専用のEdgeプロファイルを起動する。
2. 利用者が通常のSSO・MFAを完了する。
3. MCPサーバーは同じプロファイルをpersistent contextとして開く。
4. `BrowserContext.request`で`/_api/web/currentuser`をGETする。
5. Cookieの適用方法や社内プロキシとの差異で失敗した場合は、SharePointページ内の`fetch`へフォールバックする。
6. headedモードでログイン画面へ遷移した場合は、そのページを残して利用者が同じMCPプロセス内で認証できるようにする。
7. MCPツールは認証状態と最小限の利用者情報だけを返す。

## セキュリティ境界

- 対象URLは設定済みのSharePointサイト配下の`/_api/`だけに制限する。
- OneDriveホスト、テナントルート、任意URLを拒否する。
- 専用Edgeプロファイル以外を拒否する。
- プロファイルは現在のWindows利用者だけが参照できる場所へ置き、同期・共有・コミットしない。
- Cookie、アクセストークン、Authorizationヘッダー、ログイン名、メールアドレスを返却・記録しない。
- 書き込み用のRESTメソッドを実装しない。
- MCPの標準出力にはプロトコルメッセージ以外を書き込まない。

## 完了条件

- 専用EdgeプロファイルでSharePointへログインできる。
- `sharepoint_auth_status`が`AUTHENTICATED`を返す。
- セッション失効時に`LOGIN_REQUIRED`を返す。
- 権限不足、サイト不明、通信不能を区別できる。
- 通常のEdgeプロファイルを指定した場合は起動前に拒否する。
- Cookieやトークンがログ・ツール結果へ含まれない。

## Phase 1の対象外

- SharePoint検索
- ページ本文取得
- ファイル一覧・ダウンロード
- Office文書解析
- SharePointへの書き込み
- OneDrive
- Secure MCP Tunnel
