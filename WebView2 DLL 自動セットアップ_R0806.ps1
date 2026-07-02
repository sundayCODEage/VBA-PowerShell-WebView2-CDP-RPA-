# 定期的に、AIに照会（確認）すること
# プロント： WebView2 ポータブル環境用 DLL を下記コードで取得していたが、現在の安定最新に更新したい。

# R08/06/xx 更新

# ==============================================================================
# WebView2 ポータブル環境用 DLL 自動セットアップスクリプト
# ==============================================================================

# 1. 設定：バージョンとディレクトリ
$libsDir = Join-Path $PSScriptRoot "Libs"
$tempDir = Join-Path $PSScriptRoot "Temp_WebView2_Setup"

# WebView2 SDK のバージョン (最新の安定版を指定)
# $wv2Version = "1.0.2849.39" 
  $wv2Version = "1.0.4022.49"

# $packageUrl = "https://www.nuget.org/api/v2/package/Microsoft.Web.WebView2/$wv2Version"
  $packageUrl = "https://www.nuget.org/api/v2/package/Microsoft.Web.WebView2/1.0.4022.49"

# 抽出対象ファイルのマッピング (ZIP内のパス -> 展開後のファイル名)
# PowerShell 5.1 (VBA連携) を想定し net462 用を抽出します
$fileMap = @(
    @{ Src = "lib/net462/Microsoft.Web.WebView2.Core.dll";     Dest = "Microsoft.Web.WebView2.Core.dll" }
    @{ Src = "lib/net462/Microsoft.Web.WebView2.WinForms.dll"; Dest = "Microsoft.Web.WebView2.WinForms.dll" }
    @{ Src = "build/native/x64/WebView2Loader.dll";           Dest = "WebView2Loader.dll" } 
)

# 2. 準備
if (-not (Test-Path $libsDir)) { New-Item -ItemType Directory -Path $libsDir | Out-Null }
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -ItemType Directory -Path $tempDir | Out-Null

Write-Host "--- WebView2 ポータブルコンポーネント取得開始 ---" -ForegroundColor Cyan
Write-Host "Version: $wv2Version"

# 3. パッケージのダウンロード
$zipPath = Join-Path $tempDir "wv2.zip"
$extractPath = Join-Path $tempDir "extract"

Write-Host "Downloading WebView2 SDK..." -NoNewline
try {
    Invoke-WebRequest -Uri $packageUrl -OutFile $zipPath -ErrorAction Stop
    Write-Host " [Success]" -ForegroundColor Green

    Write-Host "Extracting package..." -NoNewline
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
    Write-Host " [Success]" -ForegroundColor Green

    # 4. 必要ファイルの選別とコピー
    foreach ($item in $fileMap) {
        $sourceFull = Join-Path $extractPath $item.Src
        $destFull = Join-Path $libsDir $item.Dest
        
        Write-Host "Installing: $($item.Dest)..." -NoNewline
        if (Test-Path $sourceFull) {
            Copy-Item $sourceFull -Destination $destFull -Force
            Write-Host " [Done]" -ForegroundColor Green
        } else {
            Write-Host " [Error: Not Found in Package]" -ForegroundColor Red
        }
    }
} catch {
    Write-Host " [Failed: $($_.Exception.Message)]" -ForegroundColor Red
}

# 5. 後片付け
Write-Host "Cleaning up temporary files..." -NoNewline
Remove-Item $tempDir -Recurse -Force
Write-Host " [Done]" -ForegroundColor Cyan

Write-Host "`nセットアップ完了！" -ForegroundColor White
Write-Host "以下のファイルが $libsDir に配置されました:"
Get-ChildItem $libsDir | Select-Object Name, Length | Out-String