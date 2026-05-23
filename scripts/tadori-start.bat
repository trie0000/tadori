@echo off
REM Tadori 起動ランチャー (デスクトップから double-click 用)
REM   - 中継サーバを起動 (未起動なら)
REM   - SharePoint サイトをブラウザで開く
REM
REM デスクトップにショートカットを作りたい場合:
REM   このファイルを右クリック → 送る → デスクトップ (ショートカットを作成)
REM
REM 設定:
REM   同じフォルダの tadori-ai-relay.env に TADORI_SITE_URL を設定
REM   例: TADORI_SITE_URL=https://contoso.sharepoint.com/sites/xxx
setlocal
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%tadori-start.ps1" %*
set EC=%errorlevel%
if not "%EC%"=="0" (
    echo.
    echo [tadori-start] ----------------------------------------------------------
    echo [tadori-start] エラーで終了しました (exit code %EC%)
    echo [tadori-start] 上のメッセージを確認してください。
    echo [tadori-start] ----------------------------------------------------------
    pause
)
endlocal
