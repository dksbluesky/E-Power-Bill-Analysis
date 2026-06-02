@echo off
chcp 65001 >nul
echo.
echo ⚡ 電費 Dashboard 更新中...
echo.

cd /d "D:\AI application code\E-bill & Air con\E-Power"
python update_dashboard.py --excel "電費明細_by_claude_v2.xlsm"

echo.
if %errorlevel% equ 0 (
    echo ✅ 完成！GitHub Pages 約 1-2 分鐘後更新。
) else (
    echo ❌ 發生錯誤，請檢查上方訊息。
)
echo.
pause
