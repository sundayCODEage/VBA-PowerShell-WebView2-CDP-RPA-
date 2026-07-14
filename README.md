#### 現在、READMEの修正中です。(前半部分は終了、後半のマニュアルは重複整理の必要あり)

#### 概要と開発の背景
本プロジェクトは、財務会計システム等に見られる**複雑なDOM構造、多段iframe、セキュリティ制約を突破し、高速かつ安全に自動化制御を行う汎用RPA操作エンジン**です。<br>
Selenium等の外部OSSのインストールを必要とせず、任意のフォルダに配置するだけで職場のPCでも利用できる「完全ポータブル環境」で構築されています。VBAを司令塔とし、PowerShell 5.1経由でMicrosoft Edgeのレンダリングコア（WebView2）を完全制御します。<br>
私はPowerShell等の知識がほぼ無い状態からスタートし、主にAI（Gemini）との対話を通じてこのハイブリッドアーキテクチャを完成させました。

> **私のメッセージ**
> 無料枠のAI利用ではコード規模が大きくなり限界を感じたため、私の本エンジンの機能追加・汎用開発はこれで「総括（終了）」とします。<br>
> 近年、Claude Code等のようにAIが自動でコードを組み上げる時代（_使った事も無いし、そんな職場でも無いですが_）になりつつありますが **「VBA×PowerShell×WebView2のハイブリッド制御」「CDP通信による強行突破」「Shadow DOMや多段iframeの透過」** といった現場の泥臭い課題を解決する基礎アーキテクチャとして、**新しい視点を持つ誰か**（あるいはAI自身）がこのコードをベースに発展させてくれることを願っています。<br>
> ※テストは実施していますが、未検証のコードも含まれるため、利用・改修は自己責任でお願いします。 (2026/07/12 sundayCODEage)

### ⚙️ システムアーキテクチャとハイブリッド連携
このエンジンは、WebView2（Chromiumベース）をスタンドアロンのデスクトップUI（WinForms）に埋め込み、「Native」と「CDP」の2つの通信経路を状況に応じて切り替える**ハイブリッドアーキテクチャ**を採用しています。
呼び出し元であるVBAからは、JSON形式（VBA-JSONライブラリ利用）でパラメータを送信して操作を指示します。

**動作機構の概要図**
```text
    +-------------------------------------------------------------+
    |                         VBA (Excel)                         |
    |  - コマンド送信 (StdIn.WriteLine JSONPayload)                |
    |  - 応答待ちループ (DoEvents / タイムアウト・死活監視)          |
    +------------------------------+------------------------------+
                                   | (標準入出力パイプ)
                                   v
    +-------------------------------------------------------------+
    |               PowerShell 5.1 (Ps_Engine_Core)               |
    |  - 動的ルーティング (& $method @parameters)                  |
    |  - [RESULT], [SUCCESS], [ERROR] プレフィックス制御           |
    +------------------------------+------------------------------+
                                   | (ドットソース結合モジュール群)
                                   v
    +-------------------------------------------------------------+
    |               WebView2 フォーム・ブラウザコア                 |
    +-------------------------------------------------------------+
                |                                   |
                | (経路A: Native制御)               | (経路B: CDP制御)
                | - ExecuteScriptAsync              | - WebSocket通信
                | - 標準DOM API (querySelector等)   | - 9222等のデバッグポート経由
                v                                   v
    +-------------------------------------------------------------+
    |             ターゲットWebページ (DOM / iframe)               |
    +-------------------------------------------------------------+
```  

### 役割分担と透過的アクセス機構
* **WebView2 (Native):** DOMレンダリング、セッション管理、UI表示、通常のJSインジェクションを担う（通常の画面遷移や安定したDOM要素へのアクセス）。
* **CDP (Chrome DevTools Protocol):** 標準のDOM操作（JSからの click() 等）では反応しない業務システムやSPAに対し、OSレベルに近いネイティブなマウス/キーボードイベントを直接発火させます。
* **多段 iframe 透過アクセス:** エンジン内部のJSユーティリティ（utilFindInFrames）により、対象要素がどの階層のiframeに存在していても自動探索・操作が可能です（クロスオリジン制約も適切にハンドリング）。

### 📁 構成モジュール一覧（ドットソース読み込み）
エンジンは以下のスクリプト群で構成され、用途に合わせて動的にロードされます。<br>
WebView2の稼働に必要なDLL（`Microsoft.Web.WebView2.Core.dll` 等）は、同梱のツール 「WebView2 DLL 自動セットアップ_R..」 を実行することで自動で3つのDLLがダウンロードされ、`Libs` フォルダへ格納されます。（ときどき最新バージョンの確認は必要です。）<br>
* Microsoft.Web.WebView2.Core.dll (基本コア)
* Microsoft.Web.WebView2.WinForms.dll (UI表示用)
* WebView2Loader.dll (PC内のEdgeランタイム本体と接続する重要ファイル。PowerShellから直接ロードはされません)
　<br><br>

| モジュール名 | 役割・機能 |
| :--- | :--- |
| **`Ps_Engine_Core_v202.ps1`** | **司令塔・ルーター・共通操作（Shadow DOM対応版）** |
| `Lib-WebView2_Init_v101.ps1` | ブラウザ画面（WinForms）の初期化と起動 |
| `Lib-WebView2_Native_v101.ps1` | ネイティブ通信（ExecuteScriptAsync等） |
| `Lib-WebCDP_v101.ps1` | WebSocketによるCDP高速通信・OSレベル操作エミュレート |
| `Lib-WebAction_v201.ps1` | Web標準操作（クリック、入力等） + Shadow DOM対応 |
| `Lib-WebSafeAction_v101.ps1` | 曖昧な指定から一意のCSSを逆生成し、非表示罠を回避する安全クリック（Robust DOM Utilities） |
| `Lib-DesktopUIA_v102.ps1` | デスクトップ操作（UIA） |
| `Lib-WebXPath_v101.ps1` | XPathによる特殊要素操作・自動補正 |
| `Lib-WebDebug_v101.ps1` | HTML/CSV保存、多段iframeツリー解析、スクショ等の開発支援機能 |

<details>
  <summary>（構成）●●RPA実行フォルダの例</summary>
  ├─ Ps_Engine_Core_v●●.ps1<br>
  ├─ Lib-必要モジュール_v●●.ps1<br>
  ├─ ●●VBA実行ファイル.xlsm<br>
  ├─ [ Libs ]　/ WebView2 DLL 自動セットアップ_R.. による自動セットアップ<br>
  │　　├─ Microsoft.Web.WebView2.Core.dll<br>
  │　　├─ Microsoft.Web.WebView2.WinForms.dll<br>
  │　　└─ WebView2Loader.dll<br>
  ├─ [ Logs] <br>
  │　　├─ [ SESSION_yyyymmdd_ hhmm ] <br>
  │　　│　　├─ Engine_SystemLog.txt　/ 実行ログファイル<br>
  │　　│　　└─ （デバックファイル.xxx）<br>
  │　　└─　作成上限は5回に制限中<br>
  └─ [ UserData ]　/ （自動作成）<br>
  　　　└─ [ EBWebView ]<br>
  　　　　　　├─ セッションデータやキャッシュ、Cookieを管理<br>
  　　　　　　　（独立したユーザーデータフォルダ（UDF）を生成）<br>
</details>

<details>
  <summary>（実行）●●実行ログファイルの例</summary>
<2026-07-xx 17:41:49> <Info> [System] 情報: Browser/ Width-Height (1536) - (816)<br>
<2026-07-xx 17:41:49> <Info> [System] 情報: 開発モードスイッチ (True)<br>
　VBA側から指定: If Not rpaEngine.StartEngine(sessionId, ENGINE_PATH, useCdpPort, True) Then<br>
　…<Info> [Engine] 実行: 関数名 | Params: { パラメータ }　/ False: パラメータを出力しない。<br>

<2026-07-xx 17:41:49> <Info> [System] 情報: 通信モードスイッチ (9222)　/ 通信モード (0: 標準, 9222等: CDP)<br>
<2026-07-xx 17:41:49> <Success> [System] 成功: モジュールをロード (Lib-WebAction_v201.ps1)<br>
<2026-07-xx 17:41:49> <Success> [System] 成功: モジュールをロード (Lib-WebXPath_v101.ps1)<br>
<2026-07-xx 17:41:50> <Success> [System] 成功: モジュールをロード (Lib-DesktopUIA_v102.ps1)<br>
<2026-07-xx 17:41:51> <Success> [System] 成功: モジュールをロード (Lib-WebDebug_v101.ps1)<br>
<2026-07-xx 17:41:51> <Success> [System] 成功: モジュールをロード (Lib-WebSafeAction_v101.ps1)<br>
<2026-07-xx 17:41:51> <Info> [System] 開始: ブラウザシステムの初期化 ...<br>
<2026-07-xx 17:41:51> <Success> [System] 成功: DLLロード (Microsoft.Web.WebView2.Core.dll Version: 1.0.4022.49)<br>
<2026-07-xx 17:41:51> <Success> [System] 成功: DLLロード (Microsoft.Web.WebView2.WinForms.dll Version: 1.0.4022.49)<br>
<2026-07-xx 17:41:54> <Info> [System] 起動モード: CDP有効 (Port: 9222)<br>
<2026-07-xx 17:41:55> <Success> [System] 成功: WebView2エンジン初期化完了<br>
<2026-07-xx 17:41:56> <Info> [System] 情報: 接続先ランタイム (Version: 150.0.4078.65)<br>
<2026-07-xx 17:41:56> <Success> [System] 成功: モジュールをロード (Lib-WebView2_Init_v101.ps1)<br>
<2026-07-xx 17:41:56> <Success> [System] 成功: モジュールをロード (Lib-WebView2_Native_v101.ps1)<br>
<2026-07-xx 17:41:56> <Success> [System] 成功: モジュールをロード (Lib-WebCDP_v101.ps1)<br>
<2026-07-xx 17:42:07> <Info> [System] 情報: CDPポート (9222) の状態 - Listen (127.0.0.1)<br>
<2026-07-xx 17:42:08> <Info> [System] 情報: 親プロセス監視開始 (PID: 3036, Name: EXCEL)<br>
<2026-07-xx 17:42:08> <Success> [System] 成功: エンジン待機状態<br>
<2026-07-xx 17:42:08> <Info> [Engine] 実行: Invoke-WebNavigation | Params: {"Url":"http s://●●●challenge.com/"}<br>
<2026-07-xx 17:42:09> <Info> [Engine] 実行: Wait-WebPageLoad | Params: {}<br>
<2026-07-xx 17:42:15> <Info> [P01] URL更新: https:// ●●●challenge.com/<br>
<2026-07-xx 17:42:38> <Info> [Engine] 実行: Set-EngineConfig | Params: {"EnableHighlight":false}　/ 選択ﾊｲﾗｲﾄの切り替え<br>
<2026-07-xx 17:42:38> <Info> [Set-EngineConfig] 設定変更: EnableHighlight = False　/ #● コメント漏れ<br>
<2026-07-xx 17:42:38> <Info> [Engine] 実行: Invoke-WebClick | Params: {"Selector":"button.xxx"}<br>
※ 基本的に”エラー情報”以外は返さない。<br>
</details>

 ### 📁 VBA側の構成モジュール
この汎用RPA操作エンジンは、VBAを司令塔として機能します。PowerShellへの操作指示は、パラメータをJSON形式に変換して送信します（`VBA-JSON-2.3.1` ライブラリを利用）。
具体的な使い方は、同梱のテストコード **`概要VBA（クラスモジュール、標準モジュール）`** は必須<br>
`RPAのテスト（Challenge サイト）` / `RPAのテスト（フォームのテストコード）` / `RPAエンジンの開発テストコード）` を参考にしてください。
* **クラスモジュール（名前：`Ps_Engine`）**<br>
  PowerShellプロセスの起動、標準入出力（パイプ）を通じたJSONメッセージの送受信、およびプロセスの死活監視（タイムアウト管理等）を担う、エンジン通信の心臓部となるクラスです。
* **標準モジュール（`JsonConverter` / VBA-JSON v2.3.1）**<br>
  PowerShellとのデータ連携に必須となるJSONパーサーです。<br>
  [VBA-JSON公式](https://github.com/VBA-tools/VBA-JSON) よりダウンロードしてインポートをお願いします。
* **標準モジュール（共通）**<br>
  RPAエンジンを安定稼働させるための裏方処理が詰まっています。UIフリーズを防ぐためのWindows API（`Sleep` や `GetTickCount` 等）の宣言、パラメータをJSON化するための補助関数（`CreateParams` 等）、およびPowerShellから返ってきたエラー文字列の解析（パース処理）など、VBA全体で使い回す汎用的なユーティリティ関数群です。

### 📘 運用設計とコア機能の詳細
#### 1 実行環境とデータ管理 (ユーザーデータフォルダ)
WebView2は、OS標準のEdgeブラウザとは完全に分離された独立のユーザーデータフォルダ（UDF）を生成します（ `\UserData\EBWebView` ）。これにより、RPAの処理がユーザーの手作業と干渉せず、Cookieやセッションを安全に永続化・管理できます。長期間の運用やアップデート時、古いキャッシュが悪影響を及ぼすのを防ぐため、ジョブの節目で Clear-WebCache を実行しクリーンな状態に保つことを推奨します。
#### 2 プロセス間通信と異常系のパース（VBAとの連携）
VBAとPowerShell間のパイプ通信において、プレフィックスルール（`[RESULT]`, `[SUCCESS]`, `[ERROR]`）を定義。PowerShell内で発生した例外は、VBA側の正規表現パーサーによって構造体（Timeout, 未発見, JSエラー等）に変換され、VBAの実行時エラーとして上位へ伝達されます。VBA側では、UIフリーズを防止しつつ安全にプロセスを監視する3段階の防波堤（SendAndReceive）を実装しています。<br>
VBA（呼び出し元）とPowerShell（実行エンジン）は、標準入出力（パイプ）を介した厳密なメッセージプロトコルで同期をとります。<br>
PowerShellは標準出力（STDOUT）へメッセージを返す際、行頭に必ず以下の識別子を付与します。<br>

[RESULT]... : 関数の戻り値（データ、取得文字列など）。<br>
[SUCCESS] : コマンドが完全に、正常終了したことを示すシグナル。<br>
[ERROR]... : PowerShell内で発生した例外メッセージ。<br>

PowerShell側で生成されたエラー文字列は、VBA側の正規表現パーサー（ParseRpaError）によって RpaExceptionInfo 構造体に変換され、「Timeout」「未発見」「JSエラー」といった種別として上位へ伝達されます。<br>
またVBA側では、UIフリーズを防止しつつ安全にプロセスを監視する3段階の防波堤（プロセスの生存確認、日跨ぎ対応タイムアウト、UIフリーズ防止）を実装しています。

#### 3 「型統一ラッパー」による完全な型維持
PowerShellとJavaScriptという異言語間の通信において発生する「型の消失（すべて文字列になる等）」を防ぐため、実行エンジン最下層（Invoke-WebView2NativeScript および Invoke-CdpScript）に以下のJSONラッパー構造を標準実装しました。

``` JavaScript
// エンジンの内部ラッパー構造
const result = (function() { $jsCode })();
return JSON.stringify({ 
    status: "success", 
    data: result !== undefined ? result : null 
});
``` 

スクリプト共通窓口（Invoke-WebScript）を経由してこのラッパーを通過することで、JS側で返された純粋なデータ型（Boolean, Null, Number, String）は、一切の揺らぎなくPowerShell側のネイティブな型として復元されます。<br>
これにより、 $res -eq "true" のような文字列判定のコードを「完全せん滅」し、if ($res) のような直接評価による極めてクリーンでハイスピードなコードベースを実現しています。<br>
(※JSコードをPowerShellへ渡す際、スクリプトの最終行には絶対に // によるコメントを書かないでください。ラッパーコードがコメントアウトされ構文エラーを引き起こします。)

#### 4 仮想タブ管理と「ゴーストウィンドウ」の完全破棄
業務システムで多発するポップアップ画面（window.open）に対し、エンジンは別ウィンドウを開かせず、同一フォーム内の「仮想タブ」として捕獲します。<br>
さらに、ポップアップが自ら閉じた際は、画面上に透明なUIの死骸（ゴースト）が残ってマウス操作をブロックしないよう、.GetNewClosure() を用いた厳密なガベージコレクションを自動実行し、メモリとUIをクリーンに保ちます。

#### 5 セレクタ指定ガイド (CSS / XPath)
**CSSセレクタ (Selector)**: input[name='username'] のように指定。エンジンがiframeを自動探索するため、フレーム階層を意識する必要はありません。<br>
<br>
**XPath (XPath)**: 複雑な表構造やテキストを含む要素を狙う場合に使用。//button[text()='送信'] と記述すれば、エンジン内部の Normalize-XPath 関数により自動的に [contains(normalize-space(.), '送信')] に変換され、HTMLの改行や空白によるエラーを未然に防ぎます。

## 📚 PowerShell 汎用RPA操作エンジン 開発・運用マニュアル

本エンジンに実装されている、モジュール別の全関数リファレンスです。

### 🛠️ [Core] 司令塔・ルーティングモジュール
* **`Write-DebugLog`**: コンソール出力とファイル出力（世代管理対応）を行うロギング機能。
* **`New-EngineException`**: `[ERROR]`プレフィックスでVBAへ返す例外文字列をフォーマット生成。
* **`Get-ActiveWebView`**: 現在アクティブなタブのWebView2インスタンスを取得。
* **`Set-ActiveTab`**: タブIDを直接指定してアクティブタブを切り替え（前面化）。
* **`List-Tabs`**: 起動中の全タブ情報（ID、URL、タイトル）をJSONで取得。
* **`Switch-Tab`**: タブIDによる切り替え。CDPの再接続処理も包含。
* **`Switch-TabByTitle`**: タイトルの部分一致検索によるタブ切り替え。
* **`Wait-Condition`**: UIフリーズを防止しつつ、指定条件がTrueになるまで待機（汎用）。
* **`Invoke-WebScript`**: JS実行のルーティング。CDPが有効ならCDP、失敗時はNativeへフォールバック。
* **`Set-EngineConfig`**: 実行時のエンジン設定（要素ハイライトのON/OFF等）を動的に変更。

### 🚀 [Init] & [Native] ブラウザ初期化・ネイティブ通信
* **`Clear-WebCache`**: UDFのキャッシュ、Cookie、LocalStorage等を非同期で完全削除。
* **`Invoke-WebView2NativeScript`**: `ExecuteScriptAsync` を使用したJS実行。JSONアンエスケープとリトライ機構を内包。

### ⚡ [CDP] 高速通信モジュール (WebSocket)
* **`Connect-CdpSession`**: `/json` エンドポイントからTargetIdを探査し、WebSocketセッションを確立。
* **`Invoke-CdpCommand`**: JSON-RPCメッセージの送受信。タイムアウトと自動再接続を管理。
* **`Invoke-CdpScript`**: CDP経由でのJS評価(`Runtime.evaluate`)。戻り値のJSONデコードを含む。
* **`Invoke-CdpNativeClick`**: CDPを使用し、OSレベルのマウスダウン/アップイベントを座標指定でエミュレート。
* **`Set-CdpNativeTextInput`**: CDPを使用し、キーボード入力をOSレベルでエミュレート（SPA対策）。

### 🖱️ [Action] Web標準操作モジュール
* **`Invoke-WebNavigation`**: 指定URLへのページ遷移を実行。
* **`Wait-WebPageLoad`**: DOMの `readyState=complete` を全iframe含めて再帰的に待機。
* **`Wait-WebDocumentReady`**: 画面全体の読み込みステータス完了を待機。
* **`Wait-WebUrlContains` / `Wait-WebTitleContains`**: URLやタイトルに指定文字列が含まれるまで待機。
* **`Wait-WebElement`**: 指定要素がDOM上に出現し、かつ画面上に可視化されるまで待機。
* **`Wait-WebElementInFrame`**: 指定したiframe内の要素が出現・可視化されるまで待機。
* **`Invoke-WebClickInFrame`**: 指定したiframe内の要素をスクロールしてクリック。
* **`Wait-WebElementInvisible`**: 指定要素が非表示になる、またはDOMから消滅するまで待機。
* **`Wait-WebScreenUnlock`**: 業務システム特有のローディングマスク（透過レイヤー）の解除を待機。
* **`Invoke-WebClick`**: 多段iframeを透過的に探索し、対象要素をクリック。
* **`Set-WebTextInput`**: テキストボックスに値を入力し、`input`/`change`イベントを発火。
* **`Select-WebDropdown`**: ドロップダウン（select）の指定値を選択し、`change`イベントを発火。
* **`Set-WebCheckbox`**: チェックボックスの状態（True/False）を判定し、差異があれば切り替え。
* **`Get-WebText`**: 要素の `innerText` または `value` を取得。
* **`Get-WebUrl` / `Get-WebTitle`**: 現在のURL、およびページタイトルを取得。
* **`Enable-SilentDownload`**: DLダイアログを抑制し、指定フォルダ・ファイル名での裏側ダウンロードを有効化。
* **`Wait-FileDownload`**: `.crdownload` の消失および排他ロック解除を確認し、DL完了を待機。

### 🛡️ [SafeAction] フェイルセーフ・安全クリック（Robust DOM）
* **`Get-WebCssSelectorHint`**: 曖昧なXPathから、可視状態の要素を厳密に判定し、レイアウト変更に強い一意のCSSセレクタを逆生成する。
* **`Invoke-WebSafeClick`**: 生成されたCSSセレクタを使用し、スクロールと可視性の最終確認を行った上で、隠し要素の誤爆を防ぎ安全にクリックを実行する。

### 🎯 [XPath] XPath特殊操作モジュール
* **`Normalize-XPath`**: XPathの表記揺れ（改行・空白）を自動補正する内部関数。
* **`Wait-WebXPathElement`**: XPath指定で要素の可視化を待機。デバッグ時は赤枠ハイライトを実行。
* **`Wait-WebXPathElementDisappear`**: XPath要素の非表示・消滅を待機。
* **`Invoke-WebXPathClick`**: XPath要素に対し、hover/mousedown/up等の一連のマウスイベントを完全エミュレート。
* **`Set-WebXPathTextInput`**: XPath要素へフォーカスし、テキスト入力と各種イベント発火を実行。
* **`Get-WebXPathText`**: XPath要素のタグを判別し、適切なテキスト（valueまたはinnerText）を取得。

### 🖥️ [UIA] デスクトップ操作モジュール
* **`Switch-AppWindow`**: Win32APIを用いて指定した外部ウィンドウを最前面へ引き上げ。
* **`Invoke-UiaAction`**: UIAutomationを用い、バックグラウンドパターンまたは物理キー送信でOS要素を操作。
* **`Invoke-UiaSafeSaveAs`**: 「名前を付けて保存」ダイアログを捕捉し、クリップボード経由でパスを入力・保存。
* **`Invoke-DesktopSendKeys`**: 対象のウィンドウへ物理キー（SendKeys）を送信。

### 🐛 [Debug] デバッグ・証跡モジュール
* **`Export-WebHtml`**: クロスオリジンを考慮し、全iframeを含むHTMLスナップショットを保存。
* **`Export-WebScreenshot`**: CDP、またはネイティブAPIへフォールバックして画面のPNGスクショを保存。
* **`Export-WebTableToCsv`**: テーブル要素を解析し、VBA取込用の配列文字列を含むCSVを生成。
* **`Export-WebElementsToCsv`**: 画面内の操作可能要素（input, a, button等）の属性を総ざらいしてCSV化。
* **`Export-WebFrameTreeToCsv`**: 多段iframeのネスト構造をツリー形式で解析しCSV化。
* **`Export-WindowScreenshot`**: Win32API等を使用し、ブラウザの枠を含むウィンドウ全体のスクショを保存。
* **`Export-WindowHierarchyToCsv`**: OS上で起動している全プロセスのハンドルとタイトル一覧をCSV出力。
* **`Write-DebugTextFile`**: 任意の文字列をデバッグ用テキストファイルへ追記保存。


## 現在、READMEの修正中です。

**PowerShell 汎用RPA操作エンジン 開発・運用マニュアル**<br>
(ほぼ正しい　AI作成)

#### (2) 実行環境とデータ管理<br>
(2.1) ユーザーデータフォルダ (\UserData\EBWebView)<br>
WebView2は、セッションデータやキャッシュ、Cookieを管理するために、実行フォルダ直下に独立したユーザーデータフォルダ（UDF）を生成します。<br>
•	完全な環境分離: OS標準のEdgeブラウザやユーザーのWindowsログインセッションとは完全に分離されるため、手作業とRPAが干渉しません。<br>
•	状態の永続化: アプリケーションを終了してもCookieやLocalStorageは残存し、次回起動時に引き継がれます。<br>
(2.2) キャッシュのクリアと安定化<br>
長期間の運用や業務システムのアップデート時、古いキャッシュがRPAの挙動に悪影響を及ぼすのを防ぐため、テスト開始時やジョブの節目で Clear-WebCache を実行し、クリーンな状態で処理を開始することを推奨します。<br>

#### (3) プロセス間通信とエラーハンドリング設計<br>
VBA（呼び出し元）とPowerShell（実行エンジン）は独立したプロセスとして稼働するため、標準入出力（パイプ）を介した厳密なメッセージプロトコルで同期をとります。<br>
(3.1) 通信プレフィックスルール (PowerShell -> VBA)<br>
PowerShellは標準出力（STDOUT）へメッセージを返す際、行頭に必ず以下の識別子を付与します。<br>
(1)	[RESULT]... : 関数の戻り値（データ、取得文字列など）。
(2)	[SUCCESS] : コマンドが完全に、正常終了したことを示すシグナル。
(3) [ERROR]... : PowerShell内で発生した例外メッセージ。

(3.2) 異常系のパースと伝播<br>
PowerShell側で生成されたエラー文字列は、VBA側の正規表現パーサー（ParseRpaError）によって RpaExceptionInfo 構造体に変換され、「どの関数で」「どのような種別のエラーが（Timeout, 未発見, JSエラー等）」「どのような詳細情報と共に」発生したかが正確にVBAの実行時エラーとして上位プロシージャへ伝達されます。<br>

(3.3) VBA側の堅牢な3段階監視（SendAndReceive）<br>
プロセスハングアップを防ぐため、VBA側では以下の3つの防波堤を敷いています。<br>
(1)	プロセスの生存確認: psProcess.Status を監視し、エンジンの予期せぬクラッシュを即座に検知。
(2)	日跨ぎ対応タイムアウト: Windows API GetTickCount を使用し、ミリ秒単位で安全なタイムアウト判定を実施（無限待機の防止）。
(3)	UIフリーズ防止: Sleep と DoEvents を組み合わせ、レスポンスを待ちながらもExcel自体の操作性を維持。

#### (4) セレクタ指定ガイド (CSS / XPath)<br>
(4.1) CSSセレクタの指定方法<br>
CSSセレクタを用いた操作（Set-WebTextInput等）は全てiframeを自動探索するため、通常の操作においてフレーム階層を意識する必要はありません。

•	ID指定: #login-button<br>
•	クラス指定: .submit-btn<br>
•	属性指定: input[name='username']<br>

(4.2) XPathの指定方法と自動補正機能<br>
複雑な表構造や、特定のテキストを持つ要素を狙う場合はXPathを使用します。
本エンジンでは、関数 Normalize-XPath により、表記揺れに対する自動補正が行われます。

•	記述例: //button[text()='送信']<br>
•	自動補正: 内部で自動的に [contains(normalize-space(.), '送信')] に変換され、HTMLソース上の余分な改行や空白による「要素が見つからない」エラーを未然に防ぎます。

#### (X) 「型統一ラッパー」の検証と今後の適用について
本エンジンでは、PowerShell（制御側）とJavaScript（ブラウザ側）という異なる言語間の通信において発生する「型の消失」や「文字列化による揺らぎ」を完全に排除するため、「JSONラッパー構造の全面的な基準化」を採用しています。
1. スクリプト実行の共通窓口（Invoke-WebScript）
すべてのDOM操作や値の取得、要素の待機処理において、JavaScriptを実行する際の最上位の共通窓口（ルーター）となるのが Invoke-WebScript です。
VBAから送られてきたコマンドや、内部のアクション関数から呼び出されたスクリプトは、まずこの関数に集約されます。<br>
•	ハイブリッド・ルーティング: CDPポートが有効（$global:CdpPort -gt 0）な場合は、低レイテンシでネイティブイベントを発火できる Invoke-CdpScript を選択します。通常モード、またはCDPが無効な環境では Invoke-WebView2NativeScript を選択します。<br>
•	フェイルセーフ（自動フォールバック）: CDP通信側でWebSocketの切断や致命的な通信エラーを検知した場合、システムを停止させることなく、WebView2標準のJSインジェクション（Native側）へ自動的に処理を切り替えてリトライを実行する堅牢な二重化構造が組まれています。<br>
2. ラッパー構造がもたらす「型の完全維持」のメカニズム
Invoke-WebScript によってルーティングされた先である、WebView2ネイティブ通信（Invoke-WebView2NativeScript）およびCDP通信（Invoke-CdpScript）の最下層コア部分には、渡されたJavaScriptコード（$jsCode）を評価する際、以下のラッパー構造が標準実装されています。<br>
【エンジンの内部ラッパー構造】<br>
JavaScript
const result = (function() { $jsCode })();
return JSON.stringify({ 
    status: "success", 
    data: result !== undefined ? result : null 
});
Invoke-WebScript を経由してこの最下層ラッパーを通過することにより、JavaScript側で返却された純粋なデータ型が安全にJSON化され、PowerShell側の ConvertFrom-Json によってPowerShellのネイティブなデータ型として100%完全に復元（維持）されます。<br>
•	Boolean型: JS return true; ➔ JSON true ➔ PS $true (System.Boolean)<br>
•	Null型: JS return null; ➔ JSON null ➔ PS $null<br>
•	Number型: JS return 123; ➔ JSON 123 ➔ PS 123 (System.Int32)<br>
•	String型: JS return "OK"; ➔ JSON "OK" ➔ PS "OK" (System.String)<br>
3. 異言語間連携におけるコーディングルール（完全せん滅作戦）<br>
共通窓口（Invoke-WebScript）と型統一ラッパーによって、データ型の完全な維持がアーキテクチャレベルで保証されたため、RPAエンジンの関数を開発・保守する際は以下のルールを徹底します。これにより、冗長な型判定コードを「完全せん滅」し、極めてクリーンでハイスピードなコードベースを維持します。<br>
【JavaScript側（ヒアドキュメント内）のルール】<br>
1.	純粋な型で return する: 成功/失敗の判定は return true; または return false; を使用し、無意味な文字列（return "success"; など）は使用しない。<br>
2.	末尾のコメント禁止（Syntax Error トラップの回避）: JavaScriptをヒアドキュメント（"@）でPowerShellへ渡す際、スクリプトの最終行（"@ の直前）には絶対に // によるコメントを書かない。 ※ ラッパーコードの閉じカッコ )(); までコメントアウトされてしまい、致命的な構文エラー（戻り値 $null エラー）を引き起こす原因となります。<br>
【PowerShell側のルール】<br>
1.	泥臭い文字列判定の禁止: 過去の資産に見られる $res -eq "true" や $res -match "false" のような、文字列としての保険的な比較は絶対に書かない。<br>
2.	スマートな型評価: Invoke-WebScript から受け取った戻り値は完全に型が保証されているため、以下のように直接評価する。<br>
o	成功判定: if ($res) { ... }<br>
o	失敗判定: if (-not $res) { throw ... }<br>

#### (X) 仮想タブ管理と「ゴーストウィンドウ」の完全破棄<br>
業務システムで多発するポップアップ画面（window.open）に対し、エンジンは別ウィンドウを開かせず、同一フォーム内の「仮想タブ」として捕獲します（NewWindowRequested イベント）。 さらに、ポップアップが window.close() で自ら閉じた際は、画面上に透明なUIの死骸（ゴースト）が残ってマウス操作をブロックしないよう、.GetNewClosure() を用いた厳密なガベージコレクション（Dispose および管理リストからの除外）を自動実行し、メモリとUIをクリーンに保ちます。
