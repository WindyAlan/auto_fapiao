"""发票自动重命名和校验工具 - 图形界面"""
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


class App:
    def __init__(self, root):
        root.title("Invoice Auto-Rename & Verify")
        root.geometry("700x520")
        root.resizable(True, True)

        # --- Rename section ---
        frame_rename = tk.LabelFrame(root, text="Step 1: Rename Invoices", padx=10, pady=10)
        frame_rename.pack(fill="x", padx=10, pady=(10, 5))

        tk.Label(frame_rename, text="PDF Folder:").grid(row=0, column=0, sticky="w")
        self.rename_dir = tk.StringVar()
        tk.Entry(frame_rename, textvariable=self.rename_dir, width=60).grid(row=0, column=1, padx=5)
        tk.Button(frame_rename, text="Browse", command=lambda: self.browse_dir(self.rename_dir)).grid(row=0, column=2)

        tk.Label(frame_rename, text="Contract Excel:").grid(row=1, column=0, sticky="w", pady=(5, 0))
        self.rename_excel = tk.StringVar()
        tk.Entry(frame_rename, textvariable=self.rename_excel, width=60).grid(row=1, column=1, padx=5, pady=(5, 0))
        tk.Button(frame_rename, text="Browse", command=lambda: self.browse_file(self.rename_excel)).grid(row=1, column=2, pady=(5, 0))

        self.btn_rename = tk.Button(frame_rename, text="Start Rename", command=self.run_rename, bg="#4CAF50", fg="white", width=15)
        self.btn_rename.grid(row=2, column=1, pady=(10, 0))

        # --- Verify section ---
        frame_verify = tk.LabelFrame(root, text="Step 2: Verify Invoices", padx=10, pady=10)
        frame_verify.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_verify, text="PDF Folder:").grid(row=0, column=0, sticky="w")
        self.verify_dir = tk.StringVar()
        tk.Entry(frame_verify, textvariable=self.verify_dir, width=60).grid(row=0, column=1, padx=5)
        tk.Button(frame_verify, text="Browse", command=lambda: self.browse_dir(self.verify_dir)).grid(row=0, column=2)

        tk.Label(frame_verify, text="Invoice Excel:").grid(row=1, column=0, sticky="w", pady=(5, 0))
        self.verify_excel = tk.StringVar()
        tk.Entry(frame_verify, textvariable=self.verify_excel, width=60).grid(row=1, column=1, padx=5, pady=(5, 0))
        tk.Button(frame_verify, text="Browse", command=lambda: self.browse_file(self.verify_excel)).grid(row=1, column=2, pady=(5, 0))

        self.btn_verify = tk.Button(frame_verify, text="Start Verify", command=self.run_verify, bg="#2196F3", fg="white", width=15)
        self.btn_verify.grid(row=2, column=1, pady=(10, 0))

        # --- Log output ---
        self.log = scrolledtext.ScrolledText(root, height=10, state="disabled", font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=10, pady=(5, 10))

    def browse_dir(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def browse_file(self, var):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
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
        self.log_msg(f"Running: {' '.join(cmd)}")
        self.log_msg("=" * 50)

        def worker():
            try:
                result = subprocess.run(
                    cmd, capture_output=True, text=True, encoding="utf-8", errors="replace"
                )
                if result.stdout:
                    self.log_msg(result.stdout)
                if result.returncode != 0 and result.stderr:
                    self.log_msg(f"[ERROR] {result.stderr}")
                self.log_msg(f"\n{done_msg}")
            except Exception as e:
                self.log_msg(f"[ERROR] {e}")
            finally:
                btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    def run_rename(self):
        d = self.rename_dir.get().strip()
        e = self.rename_excel.get().strip()
        if not d or not e:
            messagebox.showwarning("Missing", "Please select both PDF folder and Contract Excel.")
            return
        cmd = [sys.executable, "main.py", "rename", "--dir", d, "--excel", e]
        self.run_cmd(cmd, self.btn_rename, "Rename complete!")

    def run_verify(self):
        d = self.verify_dir.get().strip()
        e = self.verify_excel.get().strip()
        if not d or not e:
            messagebox.showwarning("Missing", "Please select both PDF folder and Invoice Excel.")
            return
        cmd = [sys.executable, "main.py", "verify", "--dir", d, "--excel", e]
        self.run_cmd(cmd, self.btn_verify, "Verify complete!")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
