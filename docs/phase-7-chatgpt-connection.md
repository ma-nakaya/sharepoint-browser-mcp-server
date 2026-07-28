# Phase 7: ChatGPT接続

## 目的

認証済みEdgeセッションを使うstdio MCPサーバーを、公開HTTPエンドポイントへ変更せずChatGPTから利用できるようにします。

接続にはOpenAI Secure MCP Tunnelを使用します。`tunnel-client`が社内PCからOpenAIへ外向きHTTPS接続を作り、受け取ったMCP要求をローカルのstdioサーバーへ転送します。SharePointサーバー、Edgeプロファイル、ローカルMCPポートをインターネットへ公開しません。

## ChatGPT互換ツール

既存の詳細ツールに加えて、会社知識・Deep Research互換の読み取り専用ツールを公開します。

### `search`

- 入力は`query`文字列だけ
- SharePoint自身の検索インデックスを使用
- `SitePages`ページ、`Lists`のリスト項目、PDF、DOCX、XLSX、PPTXだけを返却
- 各結果は`id`、`title`、引用可能な絶対`url`を持つ
- 最大10件

### `fetch`

- 入力は`search`が返した`id`だけ
- `SitePages`の`.aspx`ページは作成済み本文を取得
- `Lists/<リスト名>/DispForm.aspx?ID=...`はリスト項目の表示用フィールドを取得
- PDF、DOCX、XLSX、PPTXは既存の安全な文書抽出処理を使用
- `id`、`title`、`text`、`url`、`metadata`を返却
- 設定済みサイト外のURLと未対応形式は既存のURL境界で拒否

両ツールは`structuredContent`と同値のJSONテキストを返し、読み取り専用・非破壊・冪等として宣言します。

## 前提条件

- `npm run login`が完了し、`sharepoint_auth_status`が`AUTHENTICATED`
- OpenAI Platformで作成した`Tunnel ID`
- Tunnels Read + Use権限
- ChatGPTワークスペースでDeveloper modeを利用できること
- OpenAI公式[`tunnel-client`](https://github.com/openai/tunnel-client/releases/latest)
- `CONTROL_PLANE_API_KEY`または既存の`OPENAI_API_KEY`

Platformのトンネル設定:

<https://platform.openai.com/settings/organization/tunnels>

## セットアップ

### 1. ビルド

```powershell
npm install
npm run build
```

### 2. tunnel-clientを配置

既存の`C:\Apps\TunnelClient\tunnel-client.exe`を優先して使用します。存在しない環境では、公式の最新リリースを取得してリポジトリ内の`.tools/tunnel-client.exe`へ配置するか、`PATH`へ追加します。いずれにも配置しない場合は、後続コマンドへ`-TunnelClientPath`で実行ファイルを指定できます。

### 3. トンネルプロファイルを初期化

```powershell
npm run tunnel:init -- -TunnelId tunnel_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

既定プロファイル名は`sharepoint-browser-chatgpt`です。スクリプトは現在のプロセス、`.env.local`、`.env`の順にキーを探し、`OPENAI_API_KEY`をプロセス内だけで`CONTROL_PLANE_API_KEY`として再利用します。キー値を引数、ログ、プロファイルへ出力しません。

### 4. 診断

```powershell
npm run tunnel:doctor
```

stdio MCP起動、OpenAI側のTunnel IDと権限、外向きHTTPS接続を確認します。SharePoint認証状態は事前に`sharepoint_auth_status`で確認します。

### 5. トンネルを起動

```powershell
npm run tunnel:start
```

ChatGPTから利用している間、このプロセスを起動したままにします。

### 6. Windowsへ常駐化

`CONTROL_PLANE_API_KEY`がユーザー環境変数に設定されていることを確認してから、次を実行します。

```powershell
npm run tunnel:install-task
```

次の構成を作成し、登録直後に起動します。

- `C:\Apps\TunnelClient\run-sharepoint-mcp-tunnel.vbs`
- タスクスケジューラの`SharePoint MCP Tunnel`
- ユーザーログオン時に自動起動
- `wscript.exe`経由でコンソールウィンドウを非表示
- 異常終了時は1分間隔で最大3回再起動
- 72時間の実行時間制限なし
- ログは`C:\Apps\TunnelClient\sharepoint-mcp-tunnel.log`

タスクは既存の`gh-cli`、`outlook-com`プロファイルやタスクを変更せず、`sharepoint-browser-chatgpt`プロファイルだけを使用します。

ビルド後や設定変更後に常駐Tunnelを再起動する場合は、子プロセスを含めて単一インスタンスへ整理する専用コマンドを使用します。

```powershell
npm run tunnel:restart-task
```

タスクスケジューラ画面から短時間に停止・起動を繰り返すと、VBSの子`TunnelClient`が残る場合があります。専用コマンドは対象プロファイルのプロセスツリーだけを停止し、起動後のTunnel数が1つであることを検証します。

## ChatGPT側の設定

1. ChatGPTで`Settings → Security and login → Developer mode`を有効化
2. ChatGPT Pluginsで追加ボタンを選択
3. Connectionとして`Tunnel`を選択
4. 対象のTunnel IDを選択または入力
5. 検出された`search`、`fetch`と既存のSharePointツールを確認
6. 新しいチャットで作成したプラグインを有効化

ツール定義を変更した場合は、ChatGPT側で接続をRefreshします。

## 確認用プロンプト

- 「SharePointで安全衛生方針を検索して」
- 「05 国内出張旅費規程を検索して本文を取得して」
- 「07 海外出張旅費規程を検索して本文を取得して」
- 「その検索結果の本文を取得し、出典付きで要約して」
- 「SharePointにある今年度計画についてDeep Researchして」
- 「SharePointへファイルをアップロードして」— 読み取り専用のため実行されないこと

## セキュリティ境界

- MCPサーバーはstdioのままで、待受ポートを作らない
- `tunnel-client`から`api.openai.com:443`への外向きHTTPSだけを使用
- SharePoint Cookie、トークン、Authorizationヘッダーを抽出しない
- OpenAI APIキーをGit、コマンド引数、ログへ出さない
- Tunnel IDを対象Platform組織・ChatGPTワークスペースへ関連付け、利用者を限定する
- SharePoint本文にはプロンプトインジェクションが含まれ得るため、他の書き込み可能な外部ツールとの同時利用を最小化する

## 公式資料

- [Secure MCP Tunnel](https://developers.openai.com/api/docs/guides/secure-mcp-tunnels)
- [MCPサーバーの構築](https://developers.openai.com/plugins/build/mcp-server)
- [ChatGPTへの接続とテスト](https://developers.openai.com/plugins/deploy/connect-chatgpt)
- [`search`／`fetch`互換仕様](https://developers.openai.com/api/docs/mcp)
