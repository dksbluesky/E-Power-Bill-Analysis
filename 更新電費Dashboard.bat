@echo off
chcp 65001 >nul
echo.
echo ⚡ 電費 Dashboard 更新中...
echo.

python update_dashboard.py --excel "C:\Users\dk098\Documents\Bills\E-Power\電費明細 by claude.xlsm"

echo.
if %errorlevel% equ 0 (
    echo ✅ 完成！GitHub Pages 約 1-2 分鐘後更新。
) else (
    echo ❌ 發生錯誤，請檢查上方訊息。
)
echo.
pause
