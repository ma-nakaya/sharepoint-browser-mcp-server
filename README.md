# SharePoint Browser MCP Server

Entra IDアプリ登録を使用せず、Microsoft Edgeの認証済みセッションからSharePoint Onlineへ読み取り専用で接続するオンプレMCPサーバーです。

Phase 1の認証状態確認に加え、Phase 2では設定済みサイト内の検索と、`SitePages`ライブラリにあるSharePointページの本文取得に対応します。

## 目的

- Entra IDのアプリ登録、Client ID、Client Secret、証明書を使わない
- 利用者が通常どおりEdgeで完了したSSO・MFAセッションを再利用する
- 設定済みSharePointサイト以外へアクセスしない
- SharePoint REST APIを読み取り専用で利用する
- Cookie、トークン、Authorizationヘッダーをアプリケーションへ取り出さない

## 対応範囲

- TypeScriptのstdio MCPサーバー
- MCP専用Edgeプロファイル
- 初回・再認証用のheadedログインコマンド
- `sharepoint_auth_status` MCPツール
- `/_api/web/currentuser`による認証確認
- `sharepoint_search`による設定済みサイト内の検索
- `sharepoint_get_page`による`SitePages`内の`.aspx`ページ本文取得
- `BrowserContext.request`からページ内`fetch`へのフォールバック
- 設定値、URL制約、レスポンス解析、認証判定、検索・ページ取得の単体テスト

## 対象外

- リスト項目の汎用取得
- ファイル一覧・ダウンロード
- Office文書解析
- 書き込み、アップロード、更新、削除
- OneDrive
- Secure MCP Tunnel

## 動作要件

- Windows 10またはWindows 11
- Microsoft Edge
- Node.js 22以上
- 対象SharePointサイトへアクセスできる社内ネットワーク
- 対象SharePointサイトへアクセスできるユーザーアカウント

## セットアップ

```powershell
npm install
Copy-Item .env.example .env
```

`.env`を編集します。

```dotenv
SHAREPOINT_SITE_URL=https://tenant.sharepoint.com/sites/example
SHAREPOINT_PROFILE_DIR=C:\Apps\SharePointBrowserMcp\edge-profile
SHAREPOINT_HEADLESS=true
LOG_LEVEL=info
```

`SHAREPOINT_SITE_URL`は、SharePoint Onlineの`/sites/<name>`または`/teams/<name>`を指定してください。テナントルート、OneDrive、任意のURLは受け付けません。
`npm run login`、`npm run dev`、`npm start`は、プロジェクト直下に`.env`があれば自動的に読み込みます。

## 初回ログイン・再認証

```powershell
npm run login
```

MCP専用プロファイルでEdgeが起動します。通常どおりSSO・MFAを完了し、対象SharePointサイトが表示されたことを確認してからEdgeを閉じます。

普段利用しているEdgeプロファイルを`SHAREPOINT_PROFILE_DIR`に指定しないでください。既知のEdge既定プロファイル配下は起動前に拒否します。

## 検証

```powershell
npm run check
```

個別に実行する場合:

```powershell
npm run typecheck
npm test
```

## MCPサーバーの起動

開発時:

```powershell
npm run dev
```

ビルド後:

```powershell
npm run build
npm start
```

stdioクライアントの設定例:

```json
{
  "mcpServers": {
    "sharepoint-browser": {
      "command": "node",
      "args": [
        "C:\\Apps\\sharepoint-browser-mcp-server\\dist\\src\\index.js"
      ],
      "env": {
        "SHAREPOINT_SITE_URL": "https://tenant.sharepoint.com/sites/example",
        "SHAREPOINT_PROFILE_DIR": "C:\\Apps\\SharePointBrowserMcp\\edge-profile",
        "SHAREPOINT_HEADLESS": "true"
      }
    }
  }
}
```

## MCPツール

### `sharepoint_auth_status`

設定済みサイトに対する現在のEdgeセッション状態を確認します。

返却する状態:

| 状態 | 意味 |
| --- | --- |
| `AUTHENTICATED` | SharePoint REST APIをユーザーコンテキストで利用できる |
| `LOGIN_REQUIRED` | ログインまたは再認証が必要 |
| `ACCESS_DENIED` | ログイン済みだが対象サイトへの権限がない |
| `SITE_NOT_FOUND` | 対象サイトが存在しない、またはURL設定が誤っている |
| `UNAVAILABLE` | ネットワーク、プロキシ、SharePoint側の障害などで確認できない |

MCP結果には、表示名とサイトURLだけを含めます。メールアドレス、ログイン名、Cookie、トークン、レスポンス本文は返しません。

### `sharepoint_search`

設定済みSharePointサイト配下をキーワード検索します。

- 入力: `query`、任意の`maxResults`（1〜20、既定値10）
- 出力: タイトル、URL、ファイル拡張子、コンテンツ種別、更新日時、短い要約
- 検索APIの結果に設定サイト外のURLが含まれても、MCP結果から除外

### `sharepoint_get_page`

設定済みサイトの`SitePages`ライブラリにある`.aspx`ページから、作成済み本文をプレーンテキストで取得します。

- 入力: 絶対URLまたはサーバー相対URL
- 出力: タイトル、URL、更新日時、本文、切り詰め有無
- 本文は最大50,000文字
- 生の`CanvasContent1` HTML、Webパーツ設定、スクリプト、スタイルは返却しない
- 動的Webパーツが実行時に表示するデータは対象外

## セキュリティ境界

- 対象は設定済みSharePoint Onlineサイトのみ
- REST要求は対象サイト配下の`/_api/`のみ
- ページ取得は設定サイトの`SitePages`配下にある`.aspx`だけを許可
- 検索結果URLを再検証し、設定サイト外の結果を除外
- OneDrive、テナントルート、任意URLを拒否
- MCPツールは読み取り専用、非破壊、冪等として宣言
- Cookieやトークンを独自ファイルへエクスポートしない
- HTTPエラー時のHTML本文を読み取らない
- ログは標準エラーへ出力し、標準出力はMCPプロトコル専用とする

詳細は[`docs/phase-1-authentication.md`](docs/phase-1-authentication.md)と
[`docs/phase-2-sharepoint-read.md`](docs/phase-2-sharepoint-read.md)を参照してください。
