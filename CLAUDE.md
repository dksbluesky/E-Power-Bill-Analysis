# 電費分析 Dashboard — Claude Code 操作指南

## 專案說明
台南住宅電費追蹤系統，雙月帳單分析（流動費用 vs 公共費用）。

## 檔案結構
```
./
├── index.html            ← GitHub Pages 網頁版 dashboard（瀏覽器開）
├── update_dashboard.py   ← 同步腳本（從 Excel 更新 HTML 並 push）
└── CLAUDE.md             ← 本說明檔
```

> ⚠️ Excel 檔（電費分析.xlsm）放在本機，不進 git repo

## 常用指令

### 更新 dashboard（最常用）
```bash
python update_dashboard.py --excel "/path/to/電費分析.xlsm"
```
這個指令會：
1. 讀取 .xlsm 的 Raw Data sheet（第 4 行開始）
2. 更新 index.html 裡的 JS 資料陣列（DATA BLOCK 區塊）
3. git commit + push 到 GitHub

### 只更新 HTML，不 push（測試用）
```bash
python update_dashboard.py --excel "/path/to/電費分析.xlsm" --dry-run
```

### 自訂 commit message
```bash
python update_dashboard.py --excel "/path/to/電費分析.xlsm" --message "新增 115/07 資料"
```

## Raw Data 欄位對照（第 4 行起）
| 欄 | 欄位名稱 | 說明 |
|----|---------|------|
| A  | 帳單月份 | 格式 113/09，必填 |
| B  | 用電度數 | 整數，必填 |
| C  | 總電費   | 整數（元），必填 |
| D  | 季節     | 夏季 / 非夏季，必填 |
| E  | 費率版本 | A 或 B（公式自動帶出）|
| F  | 流動費用 | 公式自動計算 |
| G  | 公共費用 | 公式自動計算 |
| H  | 流動%   | 公式自動計算 |
| J  | 半年     | 公式自動帶出 |
| K  | 年度     | 公式自動帶出 |

## 費率版本規則
- 版本 A：113/04/01 – 114/09/30
- 版本 B：114/10/01 起

## 夏季定義
- 夏季：6月～9月（帳單月份 07、09 通常為夏季或混合）
- 非夏季：其餘月份

## 異常資料
- 113/09：312度，實收 $80（正常應為 $2,372），原因：113/07 溢繳抵扣
- 流動費用欄顯示估算值 $672，公共費用標示 N/A

## index.html 資料區塊格式
腳本自動找到並替換以下區塊：
```javascript
// ════════════════════════════════════════════
// DATA BLOCK — Claude Code 同步更新此區塊
// ════════════════════════════════════════════
const RAW=[
  {period:"113/09",deg:312,total:80,summer:true,ver:"A",anom:true,half:"113下",year:"113"},
  ...
];
// ════════════════════════════════════════════
```

## GitHub Pages
- 推上去後約 1–2 分鐘自動部署
- 網址：https://[你的帳號].github.io/[repo名稱]/
