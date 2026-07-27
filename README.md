# SharePoint Browser MCP Server

Entra IDアプリ登録を使用せず、Microsoft Edgeの認証済みセッションからSharePoint Onlineへ読み取り専用で接続するオンプレMCPサーバーです。

Phase 1では、Playwrightのpersistent contextでMCP専用Edgeプロファイルを管理し、`/_api/web/currentuser`を使って認証状態を確認します。SharePoint検索、ページ取得、ファイル取得は後続フェーズで追加します。

## 目的

- Entra IDのアプリ登録、Client ID、Client Secret、証明書を使わない
- 利用者が通常どおりEdgeで完了したSSO・MFAセッションを再利用する
- 設定済みSharePointサイト以外へアクセスしない
- SharePoint REST APIを読み取り専用で利用する
- Cookie、トークン、Authorizationヘッダーをアプリケーションへ取り出さない

## Phase 1の範囲

- TypeScriptのstdio MCPサーバー
- MCP専用Edgeプロファイル
- 初回・再認証用のheadedログインコマンド
- `sharepoint_auth_status` MCPツール
- `/_api/web/currentuser`による認証確認
- `BrowserContext.request`からページ内`fetch`へのフォールバック
- 設定値、URL制約、レスポンス解析、認証判定の単体テスト

## 対象外

- SharePoint検索
- ページ、リスト、ファイル取得
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

## セキュリティ境界

- 対象は設定済みSharePoint Onlineサイトのみ
- REST要求は対象サイト配下の`/_api/`のみ
- OneDrive、テナントルート、任意URLを拒否
- MCPツールは読み取り専用、非破壊、冪等として宣言
- Cookieやトークンを独自ファイルへエクスポートしない
- HTTPエラー時のHTML本文を読み取らない
- ログは標準エラーへ出力し、標準出力はMCPプロトコル専用とする

詳細は[`docs/phase-1-authentication.md`](docs/phase-1-authentication.md)を参照してください。
