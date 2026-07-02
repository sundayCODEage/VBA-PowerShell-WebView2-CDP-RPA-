この開発コード（の目標）

財務会計システム等に見られる複雑なDOM構造、多段iframe、セキュリティ制約を突破し、高速かつ安全に自動化制御を行う。(目指しましたが、)

（私は、powershellなどは殆ど判らない中でのスタート、ほぼAIさんの力です（主にGEMINI） ただ、無料利用での開発はコードが大きくなり限界を感じます。
汎用RPAとした事で、テストはしたつもりですが利用していないコードも多く検証・改修はお願いします。）

● システムアーキテクチャとハイブリッド連携
Microsoft EdgeのレンダリングコアであるWebView2（Chromiumベース）をスタンドアロンのデスクトップUI（WinForms）に埋め込み、
「Native（JSインジェクション）」と「CDP（WebSocket経由のデバッガ制御）」の2つの通信経路を状況に応じて切り替えるハイブリッドアーキテクチャを採用した。

•WebView2 (Native) の役割: DOMレンダリング、セッション管理、UI表示、および標準のJSインジェクション（ExecuteScriptAsync）を担う。
（通常のWeb画面遷移や安定したDOM要素へのアクセス）

•CDP の役割: 標準のDOM操作（JSからの click() 等）では反応しない厳格な業務システムや、イベントフックが複雑なSPAに対し、
OSレベルに近いネイティブなマウス/キーボードイベントを直接発火

•多段 iframe 透過アクセス: エンジン内部のJSユーティリティ（utilFindInFrames）により、対象要素がどの階層のiframeに存在していても、
フレーム切り替えを意識することなく自動探索・操作

•この汎用RPA操作エンジンは、主コードはVBAで パラメータをJSON形式（VBA-JSON-2.3.1ライブラリを利用）で 操作指示を受けます。

バージョンは　PowerShell5.1

Ps_Engine_Core_v107.ps1 （司令塔・ルーター・共通操作）
以下は、ドットソースで読み込む。

Lib-WebView2_Init_v101.ps1 （ブラウザ画面起動）
Lib-WebView2_Native_v101.ps1 （ネイティブ通信）
Lib-WebCDP_v101.ps1 （WebSocket・CDP高速通信）
Lib-WebAction_v101.ps1 （Web標準操作）
Lib-DesktopUIA_v101.ps1 （デスクトップ操作・UIA) 
Lib-WebXPath_v101.ps1 （XPathによる特殊要素操作）
Lib-WebDebug_v101.ps1 （HTML/CSV保存・スクショ・デバッグメモ）

で構成します。

WebView2の必要DDLは、WebView2 DLL 自動セットアップで、Join-Path $PSScriptRoot "Libs"　へ格納する。

（Microsoft.Web.WebView2.Core.dll / Microsoft.Web.WebView2.WinForms.dll / WebView2Loader.dll)


