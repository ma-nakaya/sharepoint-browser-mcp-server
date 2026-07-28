# SharePoint Browser MCP Server

Entra IDアプリ登録を使用せず、Microsoft Edgeの認証済みセッションからSharePoint Onlineへ読み取り専用で接続するオンプレMCPサーバーです。

Phase 1の認証状態確認、Phase 2の検索・ページ本文取得、Phase 3のファイルアクセス、Phase 4のPDF・Office文書本文抽出、Phase 5のSharePointサイト横断検索に加え、Phase 6では文書構造のアウトライン・検索・部分取得に対応します。

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
- `sharepoint_search`による設定済みサイト内の横断検索、種類・フォルダー・拡張子・更新日による絞り込み
- `sharepoint_get_page`による`SitePages`内の`.aspx`ページ本文取得
- `sharepoint_list_document_libraries`による可視ドキュメントライブラリ一覧
- `sharepoint_list_folder`による直下フォルダー・ファイル一覧
- `sharepoint_download_file`による許可形式・5 MiB以下のファイル取得
- `sharepoint_extract_document_text`によるPDF・DOCX・XLSX・PPTX本文抽出
- `sharepoint_get_document_outline`によるページ・見出し・シート・スライドの構造化
- `sharepoint_search_document`による1文書内のノード検索
- `sharepoint_get_document_nodes`による選択ノードだけの本文取得
- `BrowserContext.request`からページ内`fetch`へのフォールバック
- 設定値、URL制約、レスポンス解析、認証判定、検索、ページ、ファイル、文書抽出の単体テスト

## 対象外

- リスト項目の汎用取得
- OCR、画像内文字認識
- 旧Officeバイナリ形式（DOC、XLS、PPT）
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

SharePoint自身の検索インデックスを使い、設定済みサイト配下のページとファイルを横断検索します。

- 必須入力: `query`
- 件数・ページング: `maxResults`（1〜20、既定値10）、`startRow`（0〜50,000）
- 対象限定: `scope`（`all`、`pages`、`documents`）、任意の`folderUrl`
- 絞り込み: `fileExtensions`（最大10種類）、`modifiedAfter`、`modifiedBefore`
- 並び順: `relevance`または`modified-desc`
- 出力: タイトル、URL、結果種別、親URL、ファイル拡張子、コンテンツ種別、更新日時、短い要約、ページング情報
- 検索APIの結果に設定サイト外のURLが含まれても、MCP結果から除外
- `kind=page`のURLは`sharepoint_get_page`、対応する`kind=document`のURLは`sharepoint_extract_document_text`へ渡せる

### `sharepoint_get_page`

設定済みサイトの`SitePages`ライブラリにある`.aspx`ページから、作成済み本文をプレーンテキストで取得します。

- 入力: 絶対URLまたはサーバー相対URL
- 出力: タイトル、URL、更新日時、本文、切り詰め有無
- 本文は最大50,000文字
- 生の`CanvasContent1` HTML、Webパーツ設定、スクリプト、スタイルは返却しない
- 動的Webパーツが実行時に表示するデータは対象外

### `sharepoint_list_document_libraries`

設定済みサイト内の非表示ではないドキュメントライブラリと、そのルートフォルダーURLを返します。

### `sharepoint_list_folder`

指定フォルダー直下のフォルダーとファイルを、それぞれ最大100件まで返します。

- 入力: 絶対URLまたはサーバー相対URL、任意の`maxResults`（1〜100、既定値50）
- 出力: 名前、URL、サイズ、更新日時、バージョン、ダウンロード可否
- 設定サイト外や直接の子ではない応答項目を除外

### `sharepoint_download_file`

設定済みサイト内の許可形式ファイルをMCPの埋め込みバイナリresourceとして返します。

- 最大5 MiB
- 対応形式: PDF、DOCX、XLSX、PPTX、TXT、Markdown、CSV、JSON、XML、BMP、GIF、JPEG、PNG、WebP
- 実行形式、スクリプト、HTML、SVG、マクロ有効Office形式、その他の形式は拒否
- SharePointメタデータと実データのサイズを照合
- SHA-256をメタデータとして返却

### `sharepoint_extract_document_text`

設定済みサイトのPDF・DOCX・XLSX・PPTXを取得し、プレーンテキストとして返します。

- 抽出本文は最大100,000文字
- PDFは最大200ページ
- DOCXは本文、ヘッダー、フッター、脚注、文末脚注を抽出
- XLSXはシート名とセルの保存値を抽出
- PPTXはスライド本文と対応可能なスピーカーノートを抽出
- Office ZIPは最大1,000部品、対象XML 1部品2 MiB、対象XML合計8 MiB
- PDF内の文書アクションやOfficeマクロは実行しない
- スキャンPDFのOCR、数式・図形・画像の意味解析、Excel数式の再計算は行わない

### `sharepoint_get_document_outline`

PDF・DOCX・XLSX・PPTXを共通ノードへ変換し、短いプレビューを持つアウトラインを返します。

- PDFはページ単位
- Wordは見出し階層と補助部品単位
- Excelはシート単位
- PowerPointはスライド単位
- 各ノードに安定した形式別IDと根拠位置を付与
- 外部LLM呼び出しやローカルへの自動保存は行わない

### `sharepoint_search_document`

1つのPDF・Office文書内を検索し、関連ノードのID、根拠位置、スコア、短いスニペットを返します。

- 検索語は1〜200文字
- 最大20件、既定値10件
- スコアは同一文書・同一呼び出し内の順位付け専用

### `sharepoint_get_document_nodes`

アウトラインまたは文書内検索で得たノードIDを最大20件指定し、その部分だけの本文を返します。合計本文は最大100,000文字です。前回結果の`sha256`を`expectedSha256`へ渡すと、探索中に文書が更新された場合は取得を拒否します。

## セキュリティ境界

- 対象は設定済みSharePoint Onlineサイトのみ
- REST要求は対象サイト配下の`/_api/`のみ
- ページ取得は設定サイトの`SitePages`配下にある`.aspx`だけを許可
- 検索結果URLを再検証し、設定サイト外の結果を除外
- 検索対象フォルダー、拡張子、更新日の入力を検証し、検索結果の親URLもサイト境界と照合
- フォルダー・ファイルURLとSharePoint応答パスを再検証
- バイナリは許可形式と5 MiB上限を満たす場合だけ返却
- Office ZIPは展開前後のサイズと部品数を制限し、DOCTYPE・ENTITYを拒否
- 抽出本文を100,000文字に制限
- 構造化ノードは最大500件、選択ノードは最大20件
- 文書アウトライン、文書内検索、選択本文を外部LLMや永続インデックスへ送信・保存しない
- OneDrive、テナントルート、任意URLを拒否
- MCPツールは読み取り専用、非破壊、冪等として宣言
- Cookieやトークンを独自ファイルへエクスポートしない
- HTTPエラー時のHTML本文を読み取らない
- ログは標準エラーへ出力し、標準出力はMCPプロトコル専用とする

詳細は[`docs/phase-1-authentication.md`](docs/phase-1-authentication.md)、
[`docs/phase-2-sharepoint-read.md`](docs/phase-2-sharepoint-read.md)、
[`docs/phase-3-file-access.md`](docs/phase-3-file-access.md)、
[`docs/phase-4-document-text.md`](docs/phase-4-document-text.md)、
[`docs/phase-5-site-search.md`](docs/phase-5-site-search.md)、
[`docs/phase-6-document-structure.md`](docs/phase-6-document-structure.md)を参照してください。
