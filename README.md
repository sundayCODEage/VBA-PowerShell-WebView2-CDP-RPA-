この(汎用RPA操作エンジン)開発コード（の目標）

財務会計システム等に見られる複雑なDOM構造、多段iframe、セキュリティ制約を突破し、高速かつ安全に自動化制御を行う汎用RPA操作エンジン。(目指しましたが、)

（私は、powershellなどは殆ど判らない中でのスタート、ほぼAIさんの力です（主にGEMINI） ただ、無料利用での開発はコードが大きくなり限界を感じます。
汎用RPAとした事で、テストはしたつもりですが利用していないコードも多く検証・改修はお願いします。）

● バージョンは　PowerShell 5.1<br>
Ps_Engine_Core_v107.ps1 （司令塔・ルーター・共通操作）<br>
以下は、ドットソースで読み込む。<br>
Lib-WebView2_Init_v101.ps1 （ブラウザ画面起動）<br>
Lib-WebView2_Native_v101.ps1 （ネイティブ通信）<br>
Lib-WebCDP_v101.ps1 （WebSocket・CDP高速通信）<br>
Lib-WebAction_v101.ps1 （Web標準操作）<br>
Lib-DesktopUIA_v101.ps1 （デスクトップ操作・UIA) <br>
Lib-WebXPath_v101.ps1 （XPathによる特殊要素操作）<br>
Lib-WebDebug_v101.ps1 （HTML/CSV保存・スクショ・デバッグメモ）<br>
で構成します。

● WebView2の必要DDLは、WebView2 DLL 自動セットアップで、Join-Path $PSScriptRoot **"Libs"**　へ格納する。<br>
（Microsoft.Web.WebView2.Core.dll / Microsoft.Web.WebView2.WinForms.dll / WebView2Loader.dll)

● システムアーキテクチャとハイブリッド連携<br>

•この汎用RPA操作エンジンは、主コードはVBAで パラメータをJSON形式（VBA-JSON-2.3.1ライブラリを利用）で 操作指示を受けます。<br>
•Microsoft EdgeのレンダリングコアであるWebView2（Chromiumベース）をスタンドアロンのデスクトップUI（WinForms）に埋め込み、
「Native（JSインジェクション）」と「CDP（WebSocket経由のデバッガ制御）」の2つの通信経路を状況に応じて切り替えるハイブリッドアーキテクチャを採用した。<br>
•WebView2 (Native) の役割: DOMレンダリング、セッション管理、UI表示、および標準のJSインジェクション（ExecuteScriptAsync）を担う。
（通常のWeb画面遷移や安定したDOM要素へのアクセス）<br>
•CDP の役割: 標準のDOM操作（JSからの click() 等）では反応しない厳格な業務システムや、イベントフックが複雑なSPAに対し、
OSレベルに近いネイティブなマウス/キーボードイベントを直接発火<br>
•多段 iframe 透過アクセス: エンジン内部のJSユーティリティ（utilFindInFrames）により、対象要素がどの階層のiframeに存在していても、
フレーム切り替えを意識することなく自動探索・操作<br>

**PowerShell 汎用RPA操作エンジン 開発・運用マニュアル**<br>
(ほぼ正しい　AI作成)

概要<br>
本ドキュメントは、PowerShell 5.1およびVBA（標準入出力JSON通信）を基盤とした汎用RPA操作エンジンのアーキテクチャ、全関数仕様、および運用設計について網羅的に解説するものです。
財務会計システム等に見られる複雑なDOM構造、多段iframe、セキュリティ制約を突破し、高速かつ安全に自動化制御を行うための実践的な仕様を定義しています。

#### (1) システムアーキテクチャとハイブリッド連携原理<br>
本エンジンは、Microsoft EdgeのレンダリングコアであるWebView2（Chromiumベース）をスタンドアロンのデスクトップUI（WinForms）に埋め込み、「Native（JSインジェクション）」と「CDP（WebSocket経由のデバッガ制御）」の2つの通信経路を状況に応じて切り替えるハイブリッドアーキテクチャを採用しています。

1.1 動作機構の概要図
<pre style="font-family: 'Consolas', 'Courier New', monospace; font-size: 14px;">
　　　　　+-------------------------------------------------------------+
　　　　　|                     VBA (Excel)                             |
　　　　　|  - コマンド送信 (StdIn.WriteLine JSONPayload)                |
　　　　　|  - 応答待ちループ (DoEvents / タイムアウト・死活監視)          |
　　　　　+------------------------------+------------------------------+
　　　　　                               | (標準入出力パイプ)
　　　　　                               v
　　　　　+-------------------------------------------------------------+
　　　　　|              PowerShell 5.1 (Ps_Engine_Core)                |
　　　　　|  - 動的ルーティング (& $method @parameters)                  |
　　　　　|  - [RESULT], [SUCCESS], [ERROR] プレフィックス制御           |
　　　　　+------------------------------+------------------------------+
　　　　　                               | (ドットソース結合モジュール群)
　　　　　                               v
　　　　　+-------------------------------------------------------------+
　　　　　|                 WebView2 フォーム・ブラウザコア               |
　　　　　+-------------------------------------------------------------+
　　　　　       |                                             |
　　　　　       | (経路A: Native制御)                          | (経路B: CDP制御)
　　　　　       | - ExecuteScriptAsync                        | - WebSocket通信
　　　　　       | - 標準DOM API (querySelector等)             | - 9222等のデバッグポート経由
　　　　　       v                                             v
　　　　　+-------------------------------------------------------------+
　　　　　|                 ターゲットWebページ (DOM / iframe)           |
　　　　　+-------------------------------------------------------------+
</pre>

(1.2) 役割分担と透過的アクセス機構<br>
•	WebView2 (Native) の役割: DOMレンダリング、セッション管理、UI表示、および標準のJSインジェクション（ExecuteScriptAsync）を担います。これにより、通常のWeb画面遷移や安定したDOM要素へのアクセスを実現します。<br>
•	CDP の役割: 標準のDOM操作（JSからの click() 等）では反応しない厳格な業務システムや、イベントフックが複雑なSPAに対し、OSレベルに近いネイティブなマウス/キーボードイベントを直接発火させます。<br>
•	多段 iframe 透過アクセス: エンジン内部のJSユーティリティ（utilFindInFrames）により、対象要素がどの階層のiframeに存在していても、フレーム切り替えを意識することなく自動探索・操作が可能です（クロスオリジン制約も適切にハンドリング）。<br>

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

#### (5) モジュール別 全関数リファレンス（全53関数）
**[Core] 司令塔・ルーティングモジュール<br>**
Write-DebugLog	コンソール出力とファイル出力（世代管理対応）を行うロギング機能。<br>
New-EngineException	[ERROR]プレフィックスでVBAへ返す例外文字列をフォーマット生成。<br>
Get-ActiveWebView	現在アクティブなタブのWebView2インスタンスを取得。<br>
Set-ActiveTab	タブIDを直接指定してアクティブタブを切り替え（前面化）。<br>
List-Tabs	起動中の全タブ情報（ID、URL、タイトル）をJSONで取得。<br>
Switch-Tab	タブIDによる切り替え。CDPの再接続処理も包含。<br>
Switch-TabByTitle	タイトルの部分一致検索によるタブ切り替え。<br>
Wait-Condition	UIフリーズを防止しつつ、指定条件がTrueになるまで待機（汎用）。<br>
Invoke-WebScript	JS実行のルーティング。CDPが有効ならCDP、失敗時はNativeへフォールバック。<br>
Set-EngineConfig	実行時のエンジン設定（要素ハイライトのON/OFF等）を動的に変更。<br>

**[Init] ブラウザ初期化モジュール<br>**
Clear-WebCache	UDFのキャッシュ、Cookie、LocalStorage等を非同期で完全削除。

**[Native] ネイティブ通信モジュール<br>**
Invoke-WebView2NativeScript	ExecuteScriptAsync を使用したJS実行。JSONアンエスケープとリトライ機構を内包。

**[CDP] 高速通信モジュール (WebSocket)<br>**
Connect-CdpSession	/json エンドポイントからTargetIdを探査し、WebSocketセッションを確立。<br>
Invoke-CdpCommand	JSON-RPCメッセージの送受信。タイムアウトと自動再接続を管理。<br>
Invoke-CdpScript	CDP経由でのJS評価(Runtime.evaluate)。戻り値のJSONデコードを含む。<br>
Invoke-CdpNativeClick	CDPを使用し、OSレベルのマウスダウン/アップイベントを座標指定でエミュレート。<br>
Set-CdpNativeTextInput	CDPを使用し、キーボード入力をOSレベルでエミュレート（SPA対策）。<br>

**[Action] Web標準操作モジュール<br>**
Invoke-WebNavigation	指定URLへのページ遷移を実行。<br>
Wait-WebPageLoad	DOMの readyState=complete を全iframe含めて再帰的に待機。<br>
Wait-WebDocumentReady	画面全体の読み込みステータス完了を待機。<br>
Wait-WebUrlContains	現在のURLに指定文字列が含まれるまで待機。<br>
Wait-WebTitleContains	ページタイトルに指定文字列が含まれるまで待機。<br>
Wait-WebElement	指定要素がDOM上に出現し、かつ画面上に可視化されるまで待機。<br>
Wait-WebElementInFrame	指定したiframe内の要素が出現・可視化されるまで待機。<br>
Invoke-WebClickInFrame	指定したiframe内の要素をスクロールしてクリック。<br>
Wait-WebElementInvisible	指定要素が非表示になる、またはDOMから消滅するまで待機。<br>
Wait-WebScreenUnlock	業務システム特有のローディングマスク（透過レイヤー）の解除を待機。<br>
Invoke-WebClick	多段iframeを透過的に探索し、対象要素をクリック。<br>
Set-WebTextInput	テキストボックスに値を入力し、input/changeイベントを発火。<br>
Select-WebDropdown	ドロップダウン（select）の指定値を選択し、changeイベントを発火。<br>
Set-WebCheckbox	チェックボックスの状態（True/False）を判定し、差異があれば切り替え。<br>
Get-WebText	要素のinnerTextまたはvalueを取得。<br>
Get-WebUrl / Title	現在のURL、およびページタイトルを取得。<br>
Enable-SilentDownload	DLダイアログを抑制し、指定フォルダ・ファイル名での裏側ダウンロードを有効化。<br>
Wait-FileDownload	.crdownloadの消失および排他ロック解除を確認し、DL完了を待機。<br>

**[XPath] XPath特殊操作モジュール<br>**
Normalize-XPath	XPathの表記揺れ（改行・空白）を自動補正する内部関数。<br>
Wait-WebXPathElement	XPath指定で要素の可視化を待機。デバッグ時は赤枠ハイライトを実行。<br>
Wait-WebXPathElementDisappear	XPath要素の非表示・消滅を待機。<br>
Invoke-WebXPathClick	XPath要素に対し、hover/mousedown/up等の一連のマウスイベントを完全エミュレート。<br>
Set-WebXPathTextInput	XPath要素へフォーカスし、テキスト入力と各種イベント発火を実行。<br>
Get-WebXPathText	XPath要素のタグを判別し、適切なテキスト（valueまたはinnerText）を取得。<br>

**[UIA] デスクトップ操作モジュール<br>**
Switch-AppWindow	Win32APIを用いて指定した外部ウィンドウを最前面へ引き上げ。<br>
Invoke-UiaAction	UIAutomationを用い、バックグラウンドパターンまたは物理キー送信でOS要素を操作。<br>
Invoke-UiaSafeSaveAs	「名前を付けて保存」ダイアログを捕捉し、クリップボード経由でパスを入力・保存。<br>

**[Debug] デバッグ・証跡モジュール<br>**
Export-WebHtml	クロスオリジンを考慮し、全iframeを含むHTMLスナップショットを保存。<br>
Export-WebScreenshot	CDP、またはネイティブAPIへフォールバックして画面のPNGスクショを保存。<br>
Export-WebTableToCsv	テーブル要素を解析。人間用CSVと、画像名抽出等を含むVBA取込用配列文字列を生成。<br>
Export-WebElementsToCsv	画面内の操作可能要素（input, a, button等）の属性を総ざらいしてCSV化。<br>
Export-WebFrameTreeToCsv	多段iframeのネスト構造をツリー形式で解析しCSV化。<br>
Export-WindowScreenshot	Win32API/System.Drawingを使用し、ブラウザの枠を含むウィンドウ全体のスクショを保存。<br>
Export-WindowHierarchyToCsv	OS上で起動している全プロセスのハンドルとタイトル一覧をCSV出力。<br>
Write-DebugTextFile	任意の文字列をデバッグ用テキストファイルへ追記保存。<br>

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

