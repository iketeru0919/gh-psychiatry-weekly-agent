@echo off
chcp 932 >nul
setlocal

cd /d "%~dp0"

where py >nul 2>&1
if errorlevel 1 (
    echo [エラー] Python ランチャ ^(py^) が見つかりません。
    echo Python をインストールしてから再実行してください。
    pause
    exit /b 1
)

echo 対象ファイルの確認を開始します...
echo.
py "%~dp0jisseki_file_check.py"
set RESULT=%errorlevel%

echo.
if not "%RESULT%"=="0" (
    echo [エラー] 処理が異常終了しました ^(コード %RESULT%^)。
) else (
    echo 処理が完了しました。
)

echo 何かキーを押すと閉じます。
pause >nul

rem endlocal を単独行にすると RESULT が消えてから exit /b が展開されるため、
rem 終了コードが渡らない。同じ行に書いて先に展開させる。
endlocal & exit /b %RESULT%
