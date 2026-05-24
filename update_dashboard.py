#!/usr/bin/env python3
"""
update_dashboard.py
-------------------
讀取電費 Excel (.xlsm)，更新 index.html 的 JS 資料陣列，
然後 git commit + push 到 GitHub。

使用方式：
  python update_dashboard.py
  python update_dashboard.py --excel "電費分析.xlsm"
  python update_dashboard.py --dry-run   # 只更新 html，不 push
"""

import argparse, json, re, subprocess, sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("❌ 缺少 openpyxl，請執行：pip install openpyxl")
    sys.exit(1)

# ── 設定 ──────────────────────────────────────────────────────────────────────
DEFAULT_EXCEL = "電費分析.xlsm"   # ← 如果檔名不同請修改
HTML_FILE     = "index.html"
RAW_DATA_SHEET = "Raw Data"
DATA_ROW_START = 4   # Raw Data 資料從第幾行開始

# ── 費率計算 ──────────────────────────────────────────────────────────────────
RATES = {
    "A": {
        "s": [(120,1.68),(210,2.45),(170,3.70),(200,5.04),(299,6.24),(9999,8.46)],
        "n": [(120,1.68),(210,2.16),(170,3.03),(200,4.14),(299,5.07),(9999,6.63)],
    },
    "B": {
        "s": [(120,1.78),(210,2.55),(170,3.80),(200,5.14),(299,6.44),(9999,8.86)],
        "n": [(120,1.78),(210,2.26),(170,3.13),(200,4.24),(299,5.27),(9999,7.03)],
    },
}

def calc_flow(deg, season, ver):
    key = "s" if season == "夏季" else "n"
    tiers = RATES.get(ver, RATES["A"])[key]
    total, rem = 0, int(deg)
    for lim, rate in tiers:
        chunk = min(rem, lim); total += chunk * rate; rem -= chunk
        if rem <= 0: break
    return round(total)

# ── 讀取 Excel ────────────────────────────────────────────────────────────────
def read_excel(path: Path) -> list[dict]:
    print(f"📂 讀取 {path.name} ...")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    if RAW_DATA_SHEET not in wb.sheetnames:
        print(f"❌ 找不到工作表「{RAW_DATA_SHEET}」")
        sys.exit(1)

    ws = wb[RAW_DATA_SHEET]
    records = []

    for row in ws.iter_rows(min_row=DATA_ROW_START, values_only=True):
        period = row[0]   # A
        deg    = row[1]   # B
        total  = row[2]   # C
        season = row[3]   # D
        ver    = row[4]   # E  (may be formula result)

        if not period or str(period).strip() == "":
            continue

        period = str(period).strip()
        if not deg or not total:
            continue

        try:
            deg   = int(deg)
            total = int(total)
        except (ValueError, TypeError):
            continue

        # season default
        if not season or str(season).strip() == "":
            season = "非夏季"
        season = str(season).strip()

        # version: use formula result if available, else calculate
        if ver and str(ver).strip() in ("A", "B"):
            ver = str(ver).strip()
        else:
            yr = int(period[:3])
            mo = int(period[4:6])
            ver = "B" if (yr > 114 or (yr == 114 and mo >= 10)) else "A"

        summer = (season == "夏季")
        anom   = (total == 80 and period == "113/09")  # known anomaly
        flow   = calc_flow(deg, season, ver)
        pub    = None if anom else total - flow

        # half / year
        mo2  = int(period[4:6])
        half = period[:3] + ("上" if mo2 <= 6 else "下")
        yr2  = period[:3]

        records.append({
            "period": period,
            "deg":    deg,
            "total":  total,
            "summer": summer,
            "ver":    ver,
            "anom":   anom,
            "half":   half,
            "year":   yr2,
        })

    wb.close()
    print(f"✅ 讀取完成，共 {len(records)} 筆資料")
    return records

# ── 產生 JS 資料陣列 ──────────────────────────────────────────────────────────
def build_js_array(records: list[dict]) -> str:
    lines = ["const RAW=["]
    for r in records:
        lines.append(
            f'  {{period:"{r["period"]}",deg:{r["deg"]},total:{r["total"]},'
            f'summer:{"true" if r["summer"] else "false"},'
            f'ver:"{r["ver"]}",anom:{"true" if r["anom"] else "false"},'
            f'half:"{r["half"]}",year:"{r["year"]}"}},')
    # remove trailing comma from last line
    if lines[-1].endswith(","):
        lines[-1] = lines[-1][:-1]
    lines.append("];")
    return "\n".join(lines)

# ── 更新 HTML ─────────────────────────────────────────────────────────────────
DATA_START_MARKER = "// ══════════════════════════════════════════════"
DATA_END_MARKER   = "// ══════════════════════════════════════════════"

def update_html(html_path: Path, records: list[dict]) -> bool:
    content = html_path.read_text(encoding="utf-8")

    new_block = (
        "// ══════════════════════════════════════════════\n"
        "// DATA BLOCK — Claude Code 同步更新此區塊\n"
        "// ══════════════════════════════════════════════\n"
        + build_js_array(records) + "\n"
        "// ══════════════════════════════════════════════"
    )

    # replace between the two marker lines
    pattern = (
        r"// ═+\n"
        r"// DATA BLOCK.*?\n"
        r"// ═+\n"
        r".*?"
        r"// ═+"
    )
    new_content, count = re.subn(pattern, new_block, content, flags=re.DOTALL)

    if count == 0:
        print("⚠️  找不到 DATA BLOCK 標記，請確認 index.html 格式正確")
        return False

    html_path.write_text(new_content, encoding="utf-8")
    print(f"✅ {html_path.name} 已更新（{len(records)} 筆資料）")
    return True

# ── Git push ──────────────────────────────────────────────────────────────────
def git_push(message: str):
    cmds = [
        ["git", "add", HTML_FILE],
        ["git", "commit", "-m", message],
        ["git", "push"],
    ]
    for cmd in cmds:
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            # commit might fail if nothing changed
            if "nothing to commit" in result.stdout + result.stderr:
                print("ℹ️  資料未變動，不需要 push")
                return
            print(f"❌ 指令失敗：{' '.join(cmd)}")
            print(result.stderr)
            sys.exit(1)
    print("🚀 已成功 push 到 GitHub！")

# ── main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="更新電費 Dashboard")
    parser.add_argument("--excel",   default=DEFAULT_EXCEL, help="Excel 檔路徑")
    parser.add_argument("--dry-run", action="store_true",   help="只更新 HTML，不 push")
    parser.add_argument("--message", default="",            help="自訂 commit message")
    args = parser.parse_args()

    repo_root = Path(__file__).parent
    excel_path = Path(args.excel) if Path(args.excel).is_absolute() else repo_root / args.excel
    html_path  = repo_root / HTML_FILE

    if not excel_path.exists():
        print(f"❌ 找不到 Excel 檔案：{excel_path}")
        print(f"   請確認檔案路徑，或用 --excel 指定路徑")
        sys.exit(1)

    if not html_path.exists():
        print(f"❌ 找不到 {HTML_FILE}")
        sys.exit(1)

    records = read_excel(excel_path)
    if not records:
        print("❌ 沒有讀到任何資料，請確認 Raw Data sheet")
        sys.exit(1)

    ok = update_html(html_path, records)
    if not ok:
        sys.exit(1)

    if args.dry_run:
        print("ℹ️  dry-run 模式，跳過 git push")
        return

    # build commit message
    latest = records[-1]["period"]
    msg = args.message or f"update dashboard: 最新資料至 {latest}"
    git_push(msg)
    print(f"\n✅ 完成！網頁 dashboard 已更新至 {latest}")

if __name__ == "__main__":
    main()
