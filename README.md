この開発コード（の目標）

財務会計システム等に見られる複雑なDOM構造、多段iframe、セキュリティ制約を突破し、高速かつ安全に自動化制御を行う。(目指しましたが、)

（私は、powershellなどは殆ど判らない中でのスタート、ほぼAIさんの力です（主にGEMINI） ただ、無料利用での開発はコードが大きくなり限界を感じます。
汎用RPAとした事で、テストはしたつもりですが利用していないコードも多く検証・改修はお願いします。）

● バージョンは　PowerShell 5.1<br>
Ps_Engine_Core_v107.ps1 （司令塔・ルーター・共通操作）
以下は、ドットソースで読み込む。
/ Lib-WebView2_Init_v101.ps1 （ブラウザ画面起動）
/ Lib-WebView2_Native_v101.ps1 （ネイティブ通信）
/ Lib-WebCDP_v101.ps1 （WebSocket・CDP高速通信）
/ Lib-WebAction_v101.ps1 （Web標準操作）
/ Lib-DesktopUIA_v101.ps1 （デスクトップ操作・UIA) 
/ Lib-WebXPath_v101.ps1 （XPathによる特殊要素操作）
/ Lib-WebDebug_v101.ps1 （HTML/CSV保存・スクショ・デバッグメモ）
で構成します。

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

● WebView2の必要DDLは、WebView2 DLL 自動セットアップで、Join-Path $PSScriptRoot "Libs"　へ格納する。<br>
（Microsoft.Web.WebView2.Core.dll / Microsoft.Web.WebView2.WinForms.dll / WebView2Loader.dll)


**PowerShell 汎用RPA操作エンジン 開発・運用マニュアル**<br>
(ほぼ正しい　AI作成)

概要<br>
本ドキュメントは、PowerShell 5.1およびVBA（標準入出力JSON通信）を基盤とした汎用RPA操作エンジンのアーキテクチャ、全関数仕様、および運用設計について網羅的に解説するものです。
財務会計システム等に見られる複雑なDOM構造、多段iframe、セキュリティ制約を突破し、高速かつ安全に自動化制御を行うための実践的な仕様を定義しています。

#### (1) システムアーキテクチャとハイブリッド連携原理<br>
本エンジンは、Microsoft EdgeのレンダリングコアであるWebView2（Chromiumベース）をスタンドアロンのデスクトップUI（WinForms）に埋め込み、「Native（JSインジェクション）」と「CDP（WebSocket経由のデバッガ制御）」の2つの通信経路を状況に応じて切り替えるハイブリッドアーキテクチャを採用しています。

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

