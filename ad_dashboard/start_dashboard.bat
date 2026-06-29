@echo off
chcp 65001 > nul

:: このファイルがどこに置かれていても、ad_dashboardフォルダを探す
:: デスクトップから実行された場合はリポジトリのパスを使う
set DASHBOARD_DIR=%~dp0

:: デスクトップに置かれた場合（.batファイルの場所がad_dashboardでない）
:: ad_dashboard\app.py が存在しないときはフォルダ選択を促す
if not exist "%DASHBOARD_DIR%app.py" (
    echo [情報] app.pyが見つかりません。
    echo ad_dashboardフォルダのパスを入力してください。
    echo 例: C:\Users\yourname\gh-psychiatry-weekly-agent\ad_dashboard
    echo.
    set /p DASHBOARD_DIR="パス: "
    if not exist "%DASHBOARD_DIR%\app.py" (
        echo [エラー] app.pyが見つかりません: %DASHBOARD_DIR%
        pause
        exit /b 1
    )
)

echo ブラウザでダッシュボードを起動中...
echo 停止するにはこのウィンドウを閉じてください。
echo.
cd /d "%DASHBOARD_DIR%"
python -m streamlit run app.py --server.headless false
pause
