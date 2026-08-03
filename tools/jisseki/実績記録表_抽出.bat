@echo off
chcp 932 >nul
setlocal

rem このバッチと同じフォルダのスクリプトを実行する
cd /d "%~dp0"

where py >nul 2>&1
if errorlevel 1 (
    echo [エラー] Python ランチャ ^(py^) が見つかりません。
    echo Python をインストールしてから再実行してください。
    pause
    exit /b 1
)

echo 実績記録表の抽出を開始します...
echo.
py "%~dp0jisseki_cell_extract.py"
set RESULT=%errorlevel%

echo.
if not "%RESULT%"=="0" (
    echo [エラー] 処理が異常終了しました ^(コード %RESULT%^)。
    echo 上に表示されているメッセージを担当者へ連絡してください。
) else (
    echo 処理が完了しました。
)

echo 何かキーを押すと閉じます。
pause >nul
endlocal
exit /b %RESULT%
