#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
elec_dashboard_app.py
----------------------
簡易桌面小工具：按鈕同步電費 + 冷氣 Dashboard（呼叫既有的 update_dashboard.py）。
雙擊 啟動電費Dashboard工具.bat 開啟。
"""

import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import scrolledtext

REPO_ROOT = Path(__file__).parent
SCRIPT = REPO_ROOT / "update_dashboard.py"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("電費用量 Dashboard 工具")
        self.geometry("560x440")
        self.configure(bg="#F0F4FA")

        tk.Label(self, text="⚡ 電費 Dashboard 工具", font=("Microsoft JhengHei", 14, "bold"),
                 bg="#F0F4FA", fg="#1E3A5F").pack(pady=(16, 4))
        tk.Label(self, text="從 Excel 同步電費 / 冷氣 Dashboard 資料",
                 font=("Microsoft JhengHei", 9), bg="#F0F4FA", fg="#6B7280").pack(pady=(0, 12))

        btn_frame = tk.Frame(self, bg="#F0F4FA")
        btn_frame.pack(pady=4)

        tk.Button(btn_frame, text="🔄 同步並上傳到 GitHub", font=("Microsoft JhengHei", 11),
                  bg="#2563EB", fg="white", padx=14, pady=8, relief="flat",
                  command=self.sync_and_push).grid(row=0, column=0, padx=6)

        tk.Button(btn_frame, text="🧪 只測試（不上傳）", font=("Microsoft JhengHei", 11),
                  bg="#94A3B8", fg="white", padx=14, pady=8, relief="flat",
                  command=self.dry_run).grid(row=0, column=1, padx=6)

        self.log = scrolledtext.ScrolledText(self, height=16, font=("Consolas", 9), bg="white")
        self.log.pack(fill="both", expand=True, padx=16, pady=12)

    def write_log(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.update_idletasks()

    def _run(self, extra_args):
        self.log.delete("1.0", "end")
        cmd = [sys.executable, str(SCRIPT)] + extra_args
        self.write_log(f"$ {' '.join(cmd)}\n")
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        try:
            proc = subprocess.Popen(cmd, cwd=REPO_ROOT, stdout=subprocess.PIPE,
                                     stderr=subprocess.STDOUT, text=True, encoding="utf-8",
                                     errors="replace", env=env)
            for line in proc.stdout:
                self.write_log(line.rstrip())
            proc.wait()
            if proc.returncode == 0:
                self.write_log("\n✅ 執行完成")
            else:
                self.write_log(f"\n❌ 執行失敗（exit code {proc.returncode}）")
        except Exception as e:
            self.write_log(f"發生錯誤：{e}")

    def sync_and_push(self):
        self._run([])

    def dry_run(self):
        self._run(["--dry-run"])


if __name__ == "__main__":
    App().mainloop()
