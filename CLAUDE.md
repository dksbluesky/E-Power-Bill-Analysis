# 電費分析 Dashboard — Claude Code 操作指南

## 專案說明
台南住宅電費追蹤系統，雙月帳單分析（流動費用 vs 公共費用 vs 節電獎勵）。

## 檔案結構
```
./
├── index.html              ← 電費 GitHub Pages dashboard
├── ac_dashboard.html       ← 冷氣耗電 GitHub Pages dashboard
├── update_dashboard.py     ← 同步腳本（從 Excel 更新兩個 HTML 並 push）
├── 更新電費Dashboard.bat   ← Windows 雙擊執行同步
├── ElecBill_VBA.bas        ← VBA 模組（貼入 .xlsm 使用）
├── CLAUDE.md               ← 本說明檔
└── .gitignore              ← 排除 .xlsm / .xlsx
```

> ⚠️ Excel 檔（電費明細_by_claude_v2.xlsm）放在本機，不進 git repo

## 常用指令

### 更新 dashboard（最常用）
```bash
python update_dashboard.py --excel "C:\Users\dk098\Documents\Bills\E-Power\電費明細_by_claude_v2.xlsm"
```
這個指令會：
1. 讀取 .xlsm 的 Raw Data sheet（第 4 行開始）
2. 更新 index.html 的 JS 資料陣列（DATA BLOCK 區塊）
3. 重新計算加權均價並更新 ac_dashboard.html（RATE_NS / RATE_S）
4. git commit + push 到 GitHub（兩個 HTML 一起）

### 只更新 HTML，不 push（測試用）
```bash
python update_dashboard.py  ← 直接執行即可（路徑已內建） --dry-run
```

## Raw Data 欄位對照（第 4 行起）
| 欄 | 欄位名稱 | 說明 |
|----|---------|------|
| A  | 帳單月份 | 格式 113/09，必填 |
| B  | 用電度數 | 整數，必填 |
| C  | 總電費   | 整數（元），必填 |
| D  | 季節     | 夏季 / 非夏季，必填 |
| E  | 費率版本 | A 或 B（公式自動帶出）|
| F  | 流動費用 | 公式自動計算（雙月240度級距） |
| G  | 公共費用 | 公式 = C - F + L |
| H  | 流動%   | 公式 = F / C |
| I  | 備註     | E=實際帳單確認 / C=公式估算 |
| J  | 半年     | 公式自動帶出 |
| K  | 年度     | 公式自動帶出 |
| L  | 節電獎勵 | **手動填入**（115年起有，無則填0）|

## 費率版本規則
- 版本 A：113/04/01 – 114/09/30
- 版本 B：114/10/01 起

## 雙月級距（與台電一致）
| 段別 | 度數範圍 | 說明 |
|------|---------|------|
| 第1段 | 0–240度 | 月120度 × 2個月 |
| 第2段 | 241–660度 | |
| 第3段 | 661–1000度 | |
| 第4段 | 1001–1400度 | |
| 第5段 | 1401–2000度 | |
| 第6段 | ≥2000度 | |

## 季節定義
- 夏季：6月–9月（帳單月份 07/09 通常含夏季）
- 非夏季：其餘月份
- 混合：跨季或跨費率版本的期別（如 113/11、114/07、114/11）

## 節電獎勵
- 115年起台電開始提供節電獎勵
- 金額見台電 App 帳單明細的「節電獎勵」項目
- 新增資料後**手動填入 L 欄**
- 公共費用 = 總費用 - 流動費用 + 節電獎勵

## 異常資料
- 113/09：312度，實收 $80（正常流動費估算 $580），原因：113/07 溢繳抵扣

## 加權均價（冷氣 Dashboard 用）
由 update_dashboard.py 自動從帳單計算：
- **非夏季**：從版本B非夏季帳單（115年）算加權均價 ≈ $1.90/度
- **夏季**：從夏季帳單（114/09）算加權均價 ≈ $2.01/度

## GitHub Pages 網址
- 電費：`https://dksbluesky.github.io/E-Power-Bill-Analysis/`
- 冷氣：`https://dksbluesky.github.io/E-Power-Bill-Analysis/ac_dashboard.html`
