"""发票自动重命名和校验工具 - 图形界面"""
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


EXCEL_FILETYPES = [("Excel files", "*.xlsx *.xls *.xlsm")]


class App:
    def __init__(self, root):
        root.title("发票自动重命名和校验工具")
        root.geometry("700x520")
        root.resizable(True, True)

        # --- 重命名 ---
        frame_rename = tk.LabelFrame(root, text="第一步：重命名发票文件", padx=10, pady=10)
        frame_rename.pack(fill="x", padx=10, pady=(10, 5))

        tk.Label(frame_rename, text="PDF 文件夹:").grid(row=0, column=0, sticky="w")
        self.rename_dir = tk.StringVar()
        tk.Entry(frame_rename, textvariable=self.rename_dir, width=60).grid(row=0, column=1, padx=5)
        tk.Button(frame_rename, text="浏览", command=lambda: self.browse_dir(self.rename_dir)).grid(row=0, column=2)

        tk.Label(frame_rename, text="合同号索引 Excel:").grid(row=1, column=0, sticky="w", pady=(5, 0))
        self.rename_excel = tk.StringVar()
        tk.Entry(frame_rename, textvariable=self.rename_excel, width=60).grid(row=1, column=1, padx=5, pady=(5, 0))
        tk.Button(frame_rename, text="浏览", command=lambda: self.browse_file(self.rename_excel)).grid(row=1, column=2, pady=(5, 0))

        self.btn_rename = tk.Button(frame_rename, text="开始重命名", command=self.run_rename, bg="#4CAF50", fg="white", width=15)
        self.btn_rename.grid(row=2, column=1, pady=(10, 0))

        # --- 校验 ---
        frame_verify = tk.LabelFrame(root, text="第二步：校验发票信息", padx=10, pady=10)
        frame_verify.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_verify, text="PDF 文件夹:").grid(row=0, column=0, sticky="w")
        self.verify_dir = tk.StringVar()
        tk.Entry(frame_verify, textvariable=self.verify_dir, width=60).grid(row=0, column=1, padx=5)
        tk.Button(frame_verify, text="浏览", command=lambda: self.browse_dir(self.verify_dir)).grid(row=0, column=2)

        tk.Label(frame_verify, text="发票验证 Excel:").grid(row=1, column=0, sticky="w", pady=(5, 0))
        self.verify_excel = tk.StringVar()
        tk.Entry(frame_verify, textvariable=self.verify_excel, width=60).grid(row=1, column=1, padx=5, pady=(5, 0))
        tk.Button(frame_verify, text="浏览", command=lambda: self.browse_file(self.verify_excel)).grid(row=1, column=2, pady=(5, 0))

        self.btn_verify = tk.Button(frame_verify, text="开始校验", command=self.run_verify, bg="#2196F3", fg="white", width=15)
        self.btn_verify.grid(row=2, column=1, pady=(10, 0))

        # --- 日志 ---
        self.log = scrolledtext.ScrolledText(root, height=10, state="disabled", font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=10, pady=(5, 10))

    def browse_dir(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def browse_file(self, var):
        path = filedialog.askopenfilename(filetypes=EXCEL_FILETYPES)
        if path:
            var.set(path)

    def log_msg(self, msg):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def run_cmd(self, cmd, btn, done_msg):
        btn.config(state="disabled")
        self.log_msg(f"\n{'='*50}")
        self.log_msg(f"执行: {' '.join(cmd)}")
        self.log_msg("=" * 50)

        def worker():
            try:
                env = os.environ.copy()
                env["PYTHONIOENCODING"] = "utf-8"
                result = subprocess.run(
                    cmd, capture_output=True, encoding="utf-8", errors="replace", env=env
                )
                # Windows 终端可能是 GBK，尝试解码
                stdout = result.stdout or ""
                stderr = result.stderr or ""
                if result.returncode != 0 and stderr:
                    self.log_msg(f"[错误] {stderr}")
                if stdout:
                    self.log_msg(stdout)
                self.log_msg(f"\n{done_msg}")
            except Exception as e:
                self.log_msg(f"[错误] {e}")
            finally:
                btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    def run_rename(self):
        d = self.rename_dir.get().strip()
        e = self.rename_excel.get().strip()
        if not d or not e:
            messagebox.showwarning("提示", "请先选择 PDF 文件夹和合同号索引 Excel 文件。")
            return
        cmd = [sys.executable, "main.py", "rename", "--dir", d, "--excel", e]
        self.run_cmd(cmd, self.btn_rename, "重命名完成！")

    def run_verify(self):
        d = self.verify_dir.get().strip()
        e = self.verify_excel.get().strip()
        if not d or not e:
            messagebox.showwarning("提示", "请先选择 PDF 文件夹和发票验证 Excel 文件。")
            return
        cmd = [sys.executable, "main.py", "verify", "--dir", d, "--excel", e]
        self.run_cmd(cmd, self.btn_verify, "校验完成！")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
