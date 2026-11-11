# Phase 11: Edge の複数セッション共有

## 目的

Copilot CLI、Claude Code など複数のエージェントセッションから、同じ SharePoint MCP を同時に使えるようにする。認証は MCP 専用プロファイルへの 1 回のみとし、セッションごとのサインインは要求しない。

## 背景（従来方式の制約）

Edge（Chromium）は 1 つのプロファイルフォルダーを同時に 1 つのブラウザーしか開けない。従来は MCP サーバープロセスごとに `launchPersistentContext` で同じプロファイルを掴み続けたため、最初に SharePoint を読んだセッションが占有し、後から起動したセッションは全サイト `UNAVAILABLE` で失敗していた。

## 方式

- `SHAREPOINT_EDGE_MODE=shared`（既定）では、Edge の起動を `SharedEdgeManager`（`src/browser/shared-edge.ts`）に一元化する。
- 最初にブラウザを必要としたサーバーが、ファイルロックで起動役を決め、Edge を `--remote-debugging-address=127.0.0.1 --remote-debugging-port=<SHAREPOINT_CDP_PORT>`（既定 9229）付きで detached 起動する。起動した PID・ポート・プロファイルパスは、プロファイルの親フォルダーにある `shared-edge.json` に記録する。
- 全サーバー（起動役を含む）は `chromium.connectOverCDP()` でその Edge に接続し、既定コンテキスト（＝専用プロファイル）を共有する。各サーバーは自分のページだけを開閉する。
- サーバープロセスの終了時は CDP から切断するだけで、Edge は終了させない。次に起動したセッションは稼働中の Edge へ再接続するため、再起動も再認証も発生しない。
- Edge が落ちていた場合は、次のツール呼び出し時にロック取得→再起動が自動で起きる。セッション Cookie はプロファイルに保存されているため、再認証は不要（セッション自体が失効している場合を除く）。

## ログイン・停止との関係

- `npm run login` はプロファイルを排他で開く必要があるため、開始時に状態ファイルの PID を検証して（コマンドラインに当該プロファイルを含む場合のみ）共有 Edge を停止する。ログイン完了後、共有 Edge は次のツール呼び出しで自動的に再起動する。
- `npm run edge:stop` は共有 Edge を明示的に停止する。PID の使い回し対策として、停止時もコマンドラインが当該プロファイルを指すことを確認してから終了する。

## 設定

| 環境変数 | 既定値 | 説明 |
| --- | --- | --- |
| `SHAREPOINT_EDGE_MODE` | `shared` | `exclusive` で従来方式（セッションごとに起動）へ戻せる |
| `SHAREPOINT_CDP_PORT` | `9229` | 共有 Edge のデバッグポート（1024〜65535） |
| `SHAREPOINT_EDGE_EXECUTABLE_PATH` | 自動検出 | Program Files / LocalAppData から `msedge.exe` を探す。見つからない環境ではフルパスを指定する |

## セキュリティ境界

- デバッグポートは 127.0.0.1 のみにバインドする。CDP はブラウザ全体を操作できるため、信頼境界は「同一マシンの同一ユーザーで動くプロセス」に置く。専用プロファイルフォルダー自体がその境界内にあることと同等とみなす。
- Cookie、トークン、Authorization ヘッダーを取り出さない原則は従来どおり。CDP 接続はページ操作と `context.request` のためだけに使う。
- 全操作は読み取り専用のまま変わらない。

## 検証

- 単体テスト（`test/shared-edge.test.ts`、`test/config.test.ts`）：起動済み再利用、同時起動時の単一起動、引数生成、プロファイル照合付き停止、設定のパース。
- 実ブラウザ統合確認（Chromium で代替検証）：2 プロセスの同時接続、片方の切断後もブラウザが継続すること、全セッション終了後の再接続で再起動が起きないこと、`edge:stop` 相当でブラウザが終了すること。
- Windows 実機（Edge）では未検証。切り分け時は `SHAREPOINT_EDGE_MODE=exclusive` で従来動作へ即時復帰できる。
