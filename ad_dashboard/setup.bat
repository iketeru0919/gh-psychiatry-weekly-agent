@echo off
chcp 65001 > nul
echo ========================================
echo  ADダッシュボード セットアップ
echo ========================================
echo.

:: Pythonチェック
python --version > nul 2>&1
if errorlevel 1 (
    echo [エラー] Pythonがインストールされていません。
    echo https://www.python.org/downloads/ からインストールしてください。
    pause
    exit /b 1
)

echo [1/2] 必要なライブラリをインストール中...
python -m pip install streamlit pandas plotly openpyxl --quiet
if errorlevel 1 (
    echo [エラー] インストールに失敗しました。
    pause
    exit /b 1
)

echo [2/2] デスクトップに起動ショートカットを作成中...
set SCRIPT_DIR=%~dp0
set DESKTOP=%USERPROFILE%\Desktop

:: 起動用.batをデスクトップにコピー
copy "%SCRIPT_DIR%start_dashboard.bat" "%DESKTOP%\ADダッシュボード起動.bat" > nul

echo.
echo ========================================
echo  セットアップ完了！
echo  デスクトップの「ADダッシュボード起動.bat」を
echo  ダブルクリックしてください。
echo ========================================
pause
