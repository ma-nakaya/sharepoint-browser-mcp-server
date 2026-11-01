# SharePoint Browser MCP Server

Entra ID のアプリ登録を使わず、Microsoft Edge の認証済みセッションを再利用して SharePoint Online へ読み取り専用で接続する、TypeScript 製の stdio MCP サーバーです。利用者が通常どおり Edge で完了した SSO・MFA セッションを使い、設定済みの SharePoint サイト群だけを対象に、検索、ページ・ファイル取得、PDF・Office 文書の本文抽出と構造検索を提供します。

## 特徴

- Entra ID のアプリ登録、Client ID、Client Secret、証明書が不要
- Cookie、トークン、Authorization ヘッダーをアプリケーションへ取り出さない
- アクセス先は設定済みサイト群（同一テナント、最大 10 サイト）の `/_api/` 配下だけに制限
- すべての MCP ツールは読み取り専用。書き込み、アップロード、更新、削除は提供しない
- 複数サイトの横断検索・取得に対応

現在の到達点は Phase 10（複数サイト対応）です。Phase 1 の認証状態確認から、検索・ページ取得（Phase 2）、ファイルアクセス（Phase 3）、文書本文抽出（Phase 4）、サイト横断検索（Phase 5）、文書構造検索（Phase 6）、汎用検索・取得（Phase 7）、リスト項目取得（Phase 8）、最大 20 MiB の文書解析（Phase 9）までを含みます。各 Phase の詳細は `docs/` を参照してください。

## 対応範囲と対象外

対応:

- MCP 専用 Edge プロファイルと、初回・再認証用の headed ログインコマンド
- 認証状態確認、サイト横断検索、ページ本文取得、フォルダー・ファイル一覧と取得
- PDF・DOCX・XLSX・PPTX の本文抽出、アウトライン化、文書内検索、ノード単位の本文取得
- `Lists/.../DispForm.aspx?ID=...` 形式の SharePoint リスト項目の本文取得
- `BrowserContext.request` からページ内 `fetch` へのフォールバック

対象外:

- リスト項目の汎用取得、OCR・画像内文字認識
- 旧 Office バイナリ形式（DOC、XLS、PPT）
- OneDrive、公開 HTTP MCP エンドポイント、一般公開プラグイン

## 動作要件

- Windows 10 または Windows 11
- Microsoft Edge
- Node.js 22.16.0 以上
- 対象 SharePoint サイトへアクセスできる社内ネットワークとユーザーアカウント

## セットアップ

```powershell
npm install
Copy-Item .env.example .env
```

`.env` を編集します。

```dotenv
SHAREPOINT_SITE_URL=https://tenant.sharepoint.com/sites/example
SHAREPOINT_ADDITIONAL_SITE_URLS=https://tenant.sharepoint.com/teams/another-site
SHAREPOINT_PROFILE_DIR=C:\Apps\SharePointBrowserMcp\edge-profile
SHAREPOINT_HEADLESS=true
LOG_LEVEL=info
```

`SHAREPOINT_SITE_URL` は主サイト、`SHAREPOINT_ADDITIONAL_SITE_URLS` は追加サイトをカンマ・セミコロン・改行区切りで指定します。いずれも同じ SharePoint Online テナントの `/sites/<name>` または `/teams/<name>` である必要があります。テナントルート、OneDrive、任意の URL は受け付けません。

`npm run login`、`npm run dev`、`npm start` は、プロジェクト直下に `.env` があれば自動的に読み込みます。

### 初回ログイン・再認証

```powershell
npm run login
```

MCP 専用プロファイルで Edge が起動します。通常どおり SSO・MFA を完了し、対象 SharePoint サイトが表示されたことを確認してから Edge を閉じます。

普段利用している Edge プロファイルを `SHAREPOINT_PROFILE_DIR` に指定しないでください。既知の Edge 既定プロファイル配下は起動前に拒否します。

## 使い方

開発時:

```powershell
npm run dev
```

ビルド後:

```powershell
npm run build
npm start
```

stdio クライアントの設定例:

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
        "SHAREPOINT_ADDITIONAL_SITE_URLS": "https://tenant.sharepoint.com/teams/another-site",
        "SHAREPOINT_PROFILE_DIR": "C:\\Apps\\SharePointBrowserMcp\\edge-profile",
        "SHAREPOINT_HEADLESS": "true"
      }
    }
  }
}
```

## 設定

| 環境変数 | 既定値 | 説明 |
| --- | --- | --- |
| `SHAREPOINT_SITE_URL` | （必須） | 主サイトの URL |
| `SHAREPOINT_ADDITIONAL_SITE_URLS` | 空 | 追加サイトの URL（カンマ・セミコロン・改行区切り、同一テナントのみ、合計最大 10 サイト） |
| `SHAREPOINT_PROFILE_DIR` | OS 別のユーザー領域 | MCP 専用 Edge プロファイルの保存先 |
| `SHAREPOINT_BROWSER_CHANNEL` | `msedge` | 使用する Edge チャネル（`msedge`／`msedge-beta`／`msedge-dev`／`msedge-canary`） |
| `SHAREPOINT_HEADLESS` | `true` | Edge をヘッドレスで起動するか |
| `SHAREPOINT_REQUEST_TIMEOUT_MS` | `15000` | リクエストのタイムアウト（ミリ秒） |
| `LOG_LEVEL` | `info` | ログレベル（`.env.example` 記載） |

## ツール一覧

### `search` / `fetch`

汎用の読み取り専用検索・取得ツールです。`search` は設定済みサイト群から引用可能なページ、リスト項目、対応文書を横断検索し、`fetch` は返された ID から対象サイトを選び、本文と引用 URL を取得します。

- `search` 入力: `query`。出力: `id`、`title`、絶対 `url` を持つ最大 10 件の `results`
- `fetch` 入力: `search` が返した `id`。出力: `id`、`title`、`text`、`url`、`metadata`
- 対応対象: `SitePages` の `.aspx`、`Lists` 直下の `DispForm.aspx?ID=...`、PDF、DOCX、XLSX、PPTX
- 規程・規則は文書ファイルとは限らないため、最初に正確な名称を `search` へ渡してください

リスト項目取得は [docs/phase-8-list-item-fetch.md](docs/phase-8-list-item-fetch.md) を参照してください。

### `sharepoint_auth_status`

設定済み全サイトに対する現在の Edge セッション状態を確認します。全サイトが読める場合だけ全体状態を `AUTHENTICATED` とし、サイト別状態も返します。

| 状態 | 意味 |
| --- | --- |
| `AUTHENTICATED` | SharePoint REST API をユーザーコンテキストで利用できる |
| `LOGIN_REQUIRED` | ログインまたは再認証が必要 |
| `ACCESS_DENIED` | ログイン済みだが対象サイトへの権限がない |
| `SITE_NOT_FOUND` | 対象サイトが存在しない、または URL 設定が誤っている |
| `UNAVAILABLE` | ネットワーク、プロキシ、SharePoint 側の障害などで確認できない |

MCP 結果には表示名とサイト URL だけを含めます。メールアドレス、ログイン名、Cookie、トークン、レスポンス本文は返しません。認証確認は `/_api/web/currentuser` を使います。

### `sharepoint_search`

SharePoint 自身の検索インデックスを使い、設定済み全サイト配下のページとファイルを横断検索します。

- 必須入力: `query`
- サイト限定: 任意の `siteUrl`。省略時は全設定サイトを検索
- 件数・ページング: `maxResults`（1〜20、既定値 10）、`startRow`（0〜50,000）
- 対象限定: `scope`（`all`、`pages`、`documents`）、任意の `folderUrl`
- 絞り込み: `fileExtensions`（最大 10 種類）、`modifiedAfter`、`modifiedBefore`
- 並び順: `relevance` または `modified-desc`
- 出力: タイトル、URL、結果種別、親 URL、ファイル拡張子、コンテンツ種別、更新日時、短い要約、ページング情報
- 複数サイト検索のページングでは `siteUrl` または `folderUrl` で 1 サイトに限定します
- 検索 API の結果に設定サイト群外の URL が含まれても、MCP 結果から除外します
- `kind=page` の URL は `sharepoint_get_page`、対応する `kind=document` の URL は `sharepoint_extract_document_text` へ渡せます

### `sharepoint_get_page`

設定済みサイト群の `SitePages` ライブラリにある `.aspx` ページから、作成済み本文をプレーンテキストで取得します。

- 入力: 絶対 URL またはサーバー相対 URL
- 出力: タイトル、URL、更新日時、本文、切り詰め有無
- 本文は最大 50,000 文字
- 生の `CanvasContent1` HTML、Web パーツ設定、スクリプト、スタイルは返しません
- 動的 Web パーツが実行時に表示するデータは対象外です

### `sharepoint_list_document_libraries` / `sharepoint_list_folder`

`sharepoint_list_document_libraries` は、設定済み全サイト内の非表示ではないドキュメントライブラリと、そのルートフォルダー URL を返します。任意の `siteUrl` で 1 サイトに限定できます。

`sharepoint_list_folder` は、指定フォルダー直下のフォルダーとファイルをそれぞれ最大 100 件まで返します。

- 入力: 絶対 URL またはサーバー相対 URL、任意の `maxResults`（1〜100、既定値 50）
- 出力: 名前、URL、サイズ、更新日時、バージョン、ダウンロード可否
- 設定サイト群外や、直接の子ではない応答項目は除外します

### `sharepoint_download_file`

設定済みサイト群内の許可形式ファイルを、MCP の埋め込みバイナリ resource として返します。

- 最大 5 MiB
- 対応形式: PDF、DOCX、XLSX、PPTX、TXT、Markdown、CSV、JSON、XML、BMP、GIF、JPEG、PNG、WebP
- 実行形式、スクリプト、HTML、SVG、マクロ有効 Office 形式、その他の形式は拒否
- SharePoint メタデータと実データのサイズを照合し、SHA-256 をメタデータとして返却

### `sharepoint_extract_document_text`

設定済みサイト群の PDF・DOCX・XLSX・PPTX を取得し、プレーンテキストとして返します。

- 解析元ファイルは最大 20 MiB、抽出本文は最大 100,000 文字、PDF は最大 200 ページ
- DOCX は本文、ヘッダー、フッター、脚注、文末脚注を抽出
- XLSX はシート名とセルの保存値を抽出
- PPTX はスライド本文と対応可能なスピーカーノートを抽出
- Office ZIP は最大 1,000 部品、対象 XML 1 部品 2 MiB、対象 XML 合計 8 MiB
- PDF 内の文書アクションや Office マクロは実行しません
- スキャン PDF の OCR、数式・図形・画像の意味解析、Excel 数式の再計算は行いません

### `sharepoint_get_document_outline`

PDF・DOCX・XLSX・PPTX を共通ノードへ変換し、短いプレビューを持つアウトラインを返します。PDF はページ単位、Word は見出し階層と補助部品単位、Excel はシート単位、PowerPoint はスライド単位です。各ノードに安定した形式別 ID と根拠位置を付与します。外部 LLM 呼び出しやローカルへの自動保存は行いません。

### `sharepoint_search_document`

1 つの PDF・Office 文書内を検索し、関連ノードの ID、根拠位置、スコア、短いスニペットを返します。検索語は 1〜200 文字、最大 20 件（既定値 10 件）です。スコアは同一文書・同一呼び出し内の順位付け専用です。

### `sharepoint_get_document_nodes`

アウトラインまたは文書内検索で得たノード ID を最大 20 件指定し、その部分だけの本文を返します。合計本文は最大 100,000 文字です。前回結果の `sha256` を `expectedSha256` へ渡すと、探索中に文書が更新された場合は取得を拒否します。

## セキュリティと制約

- 対象は設定済み SharePoint Online サイト群のみ（最大 10 サイト、同一テナント）。REST 要求は対象サイト配下の `/_api/` のみです
- ページ取得は各設定サイトの `SitePages` 配下にある `.aspx` だけ、リスト項目取得は各設定サイトの `Lists/<リスト名>/DispForm.aspx` と単一の正の整数 `ID` だけを許可します
- 検索結果 URL を再検証し、設定サイト群外の結果を除外します。検索対象フォルダー、拡張子、更新日の入力を検証し、検索結果の親 URL もサイト境界と照合します
- フォルダー・ファイル URL と SharePoint 応答パスを再検証します
- MCP resource として返すバイナリは、許可形式と 5 MiB 上限を満たす場合だけ返します
- PDF・Office 文書の解析元ファイルは 20 MiB まで許可し、元バイナリ自体は MCP 応答へ含めません
- Office ZIP は展開前後のサイズと部品数を制限し、DOCTYPE・ENTITY を拒否します。抽出本文は 100,000 文字に制限します
- 構造化ノードは最大 500 件、選択ノードは最大 20 件です
- 文書アウトライン、文書内検索、選択本文を外部 LLM や永続インデックスへ送信・保存しません
- OneDrive、テナントルート、任意 URL は拒否します
- MCP ツールは読み取り専用、非破壊、冪等として宣言しています
- Cookie やトークンを独自ファイルへエクスポートしません。HTTP エラー時の HTML 本文は読み取りません
- ログは標準エラーへ出力し、標準出力は MCP プロトコル専用とします

詳細は [Phase 1: 認証](docs/phase-1-authentication.md)、[Phase 2: SharePoint 読み取り](docs/phase-2-sharepoint-read.md)、[Phase 3: ファイルアクセス](docs/phase-3-file-access.md)、[Phase 4: 文書本文](docs/phase-4-document-text.md)、[Phase 5: サイト検索](docs/phase-5-site-search.md)、[Phase 6: 文書構造](docs/phase-6-document-structure.md)、[Phase 8: リスト項目](docs/phase-8-list-item-fetch.md)、[Phase 9: 大規模文書](docs/phase-9-large-document-processing.md)、[Phase 10: 複数サイト](docs/phase-10-multi-site.md) を参照してください。

## 開発

```powershell
npm run check
```

個別に実行する場合:

```powershell
npm run typecheck
npm test
```

設定値、URL 制約、レスポンス解析、認証判定、検索、ページ、ファイル、文書抽出の単体テストを含みます。

## ライセンス

MIT License（`LICENSE` を参照）
