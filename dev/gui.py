# -*- coding: utf-8 -*-
"""
运输协议邮件外发 — 图形界面入口。

提供：SMTP 账号维护、签名维护、项目文件夹创建、日志查看。
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading

# 确保 dev/ 目录在 sys.path 中以导入 config 和 send_emails_smtp
if not getattr(sys, "frozen", False):
    _dev_dir = os.path.dirname(os.path.abspath(__file__))
    if _dev_dir not in sys.path:
        sys.path.insert(0, _dev_dir)

from config import (
    default_root_dir,
    load_smtp_credentials,
    save_smtp_credentials,
    load_signature,
    save_signature,
    load_log_content,
    create_project_folders,
    list_project_folders,
)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("运输协议邮件外发")
        self.root.minsize(520, 420)
        self._root_dir_value = default_root_dir()

        self._build_ui()

    def _root(self):
        return self._root_dir_value

    def _build_ui(self):
        # 选项卡
        self._nb = ttk.Notebook(self.root)
        self._nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        nb = self._nb

        # 1) SMTP
        smtp_frame = ttk.Frame(nb, padding=8)
        nb.add(smtp_frame, text="SMTP 账号")
        ttk.Label(smtp_frame, text="用户名 (邮箱):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self._smtp_user = ttk.Entry(smtp_frame, width=48)
        self._smtp_user.grid(row=1, column=0, sticky=tk.EW, pady=2)
        ttk.Label(smtp_frame, text="密码:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self._smtp_pass = ttk.Entry(smtp_frame, width=48, show="*")
        self._smtp_pass.grid(row=3, column=0, sticky=tk.EW, pady=2)
        smtp_btn = ttk.Frame(smtp_frame)
        smtp_btn.grid(row=4, column=0, sticky=tk.W, pady=8)
        ttk.Button(smtp_btn, text="加载", command=self._on_load_smtp).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(smtp_btn, text="保存", command=self._on_save_smtp).pack(side=tk.LEFT)
        smtp_frame.columnconfigure(0, weight=1)

        # 2) 签名
        sig_frame = ttk.Frame(nb, padding=8)
        nb.add(sig_frame, text="邮件签名")
        ttk.Label(sig_frame, text="Signature.txt 内容:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self._sig_text = scrolledtext.ScrolledText(sig_frame, width=60, height=12, wrap=tk.WORD)
        self._sig_text.grid(row=1, column=0, sticky=tk.NSEW, pady=2)
        sig_btn = ttk.Frame(sig_frame)
        sig_btn.grid(row=2, column=0, sticky=tk.W, pady=4)
        ttk.Button(sig_btn, text="加载", command=self._on_load_signature).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(sig_btn, text="保存", command=self._on_save_signature).pack(side=tk.LEFT)
        sig_frame.columnconfigure(0, weight=1)
        sig_frame.rowconfigure(1, weight=1)

        # 3) 项目
        proj_frame = ttk.Frame(nb, padding=8)
        nb.add(proj_frame, text="项目管理")
        self._proj_tab_index = 2
        ttk.Label(proj_frame, text="勾选需要外发的项目（名称以「项目」结尾的文件夹）：").grid(
            row=0, column=0, sticky=tk.W, pady=2
        )
        proj_toolbar = ttk.Frame(proj_frame)
        proj_toolbar.grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Button(proj_toolbar, text="刷新列表", command=self._on_refresh_projects).pack(side=tk.LEFT, padx=(0, 8))
        self._proj_check_frame = ttk.Frame(proj_frame)
        self._proj_check_frame.grid(row=2, column=0, sticky=tk.NSEW, pady=4)
        proj_frame.rowconfigure(2, weight=1)
        self._project_vars = {}  # folder_name -> tk.BooleanVar()
        ttk.Separator(proj_frame, orient=tk.HORIZONTAL).grid(row=3, column=0, sticky=tk.EW, pady=8)
        ttk.Label(proj_frame, text="新建项目（在工作根目录下创建「项目名称/待外发」与「项目名称/已外发」）：").grid(
            row=4, column=0, sticky=tk.W, pady=2
        )
        self._project_name = ttk.Entry(proj_frame, width=48)
        self._project_name.grid(row=5, column=0, sticky=tk.EW, pady=2)
        ttk.Button(proj_frame, text="创建项目文件夹", command=self._on_create_project).grid(
            row=6, column=0, sticky=tk.W, pady=8
        )
        proj_frame.columnconfigure(0, weight=1)

        # 4) 日志
        log_frame = ttk.Frame(nb, padding=8)
        nb.add(log_frame, text="日志")
        ttk.Label(log_frame, text="email_smtp_log.log:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self._log_text = scrolledtext.ScrolledText(log_frame, width=60, height=14, wrap=tk.WORD)
        self._log_text.grid(row=1, column=0, sticky=tk.NSEW, pady=2)
        ttk.Button(log_frame, text="刷新", command=self._on_refresh_log).grid(row=2, column=0, sticky=tk.W, pady=4)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)

        # 底部：进度条 + 剩余时间 + 发送/取消按钮
        bottom = ttk.Frame(self.root, padding="8 8")
        bottom.pack(fill=tk.X)

        progress_frame = ttk.Frame(bottom)
        progress_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        style = ttk.Style()
        try:
            style.theme_use(style.theme_use())
        except Exception:
            pass
        style.configure("green.Horizontal.TProgressbar", troughcolor="#f0f0f0", background="#4caf50")

        self._progress_var = tk.DoubleVar(value=0)
        self._progress_bar = ttk.Progressbar(progress_frame, variable=self._progress_var, maximum=100, style="green.Horizontal.TProgressbar")
        self._progress_bar.pack(fill=tk.X, expand=True)

        self._percent_label = ttk.Label(progress_frame, text="0%", background="", anchor=tk.CENTER)
        self._percent_label.place(in_=self._progress_bar, relx=0.5, rely=0.5, anchor=tk.CENTER)

        self._eta_label = ttk.Label(bottom, text="剩余时间: --", anchor=tk.CENTER)
        self._eta_label.pack(side=tk.LEFT, padx=12)

        right_frame = ttk.Frame(bottom)
        right_frame.pack(side=tk.RIGHT)

        self._btn_cancel = ttk.Button(right_frame, text="取消", command=self._on_cancel_send, state=tk.DISABLED)
        self._btn_cancel.pack(side=tk.RIGHT, padx=(6, 0))

        self._btn_send = ttk.Button(right_frame, text="开始批量发送", command=self._on_start_batch_send)
        self._btn_send.pack(side=tk.RIGHT)

    # ----- 项目管理 -----

    def _on_refresh_projects(self):
        for w in self._proj_check_frame.winfo_children():
            w.destroy()
        self._project_vars.clear()
        root = self._root()
        for name in list_project_folders(root):
            var = tk.BooleanVar(value=False)
            self._project_vars[name] = var
            ttk.Checkbutton(self._proj_check_frame, text=name, variable=var).pack(anchor=tk.W)
        if not self._project_vars:
            ttk.Label(
                self._proj_check_frame,
                text="（当前根目录下没有以「项目」结尾的文件夹，请先创建）",
                foreground="gray",
            ).pack(anchor=tk.W)

    def _get_selected_projects(self):
        return [name for name, var in self._project_vars.items() if var.get()]

    # ----- 批量发送 -----

    def _on_start_batch_send(self):
        """开始批量发送：未选项目时提示；已选则确认数量后执行。"""
        selected = self._get_selected_projects()
        if not selected:
            messagebox.showwarning(
                "未选择项目",
                "请先在「项目管理」选项卡中勾选要外发的项目，再点击「开始批量发送」。",
            )
            self._nb.select(self._proj_tab_index)
            return

        # 预览任务数量
        try:
            from send_emails_smtp import count_pending_tasks
            task_count = count_pending_tasks(self._root(), project_names=selected)
        except Exception:
            task_count = None

        if task_count == 0:
            messagebox.showinfo("无待发送任务", "选中的项目下没有待发送的文件，或无匹配的供应商邮箱。")
            return

        if task_count is not None:
            confirm_msg = f"即将向 {task_count} 个供应商发送邮件\n项目：{', '.join(selected)}\n\n确定开始发送？"
        else:
            confirm_msg = f"即将对以下项目执行批量发送：\n{', '.join(selected)}\n\n确定开始发送？"

        if not messagebox.askyesno("确认发送", confirm_msg):
            return

        # 重置进度条
        self._progress_var.set(0)
        self._percent_label.config(text="0%")
        self._eta_label.config(text="剩余时间: --")

        def worker():
            try:
                from send_emails_smtp import main as smtp_main
                result = smtp_main(
                    self._root(),
                    project_names=selected,
                    progress_callback=self._progress_callback,
                    stop_event=self._stop_event,
                )
                if result:
                    s = result.get("success", 0)
                    f = result.get("failed", 0)
                    cancelled = result.get("cancelled", False)
                    if cancelled:
                        summary = f"发送已取消。\n已发送成功: {s} 封，失败: {f} 封。"
                    else:
                        summary = f"发送完成。\n成功: {s} 封，失败: {f} 封。"
                    self.root.after(0, lambda: messagebox.showinfo("批量发送结果", summary))
                else:
                    self.root.after(0, lambda: messagebox.showinfo("批量发送", "发送任务已执行，请到「日志」选项卡查看结果。"))
                self.root.after(0, self._on_refresh_log)
            except ImportError:
                self.root.after(0, lambda: messagebox.showinfo(
                    "提示",
                    "SMTP 发送模块未找到（send_emails_smtp.py），请先部署后再使用「开始批量发送」。",
                ))
            finally:
                self.root.after(0, lambda: self._btn_send.config(state=tk.NORMAL))
                self.root.after(0, lambda: self._btn_cancel.config(state=tk.DISABLED))

        self._btn_send.config(state=tk.DISABLED)
        self._btn_cancel.config(state=tk.NORMAL)
        # 每次发送前创建新的 stop_event，避免上次取消后遗留
        self._stop_event = threading.Event()
        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _progress_callback(self, percent, rate, eta_seconds, completed, total):
        def cb():
            self._progress_var.set(percent)
            self._percent_label.config(text=f"{int(percent)}%")
            if eta_seconds is None:
                self._eta_label.config(text="剩余时间: --")
            else:
                total_s = int(eta_seconds)
                h, rem = divmod(total_s, 3600)
                m, s = divmod(rem, 60)
                if h:
                    text = f"剩余时间: {h:d}:{m:02d}:{s:02d}"
                else:
                    text = f"剩余时间: {m:02d}:{s:02d}"
                self._eta_label.config(text=text)
        self.root.after(0, cb)

    def _on_cancel_send(self):
        if hasattr(self, "_stop_event") and self._stop_event is not None:
            self._stop_event.set()
            self._btn_cancel.config(state=tk.DISABLED)
            messagebox.showinfo("取消", "已请求取消发送任务，正在停止中…")

    # ----- SMTP -----

    def _on_load_smtp(self, silent=False):
        u, p = load_smtp_credentials(self._root())
        self._smtp_user.delete(0, tk.END)
        self._smtp_user.insert(0, u)
        self._smtp_pass.delete(0, tk.END)
        self._smtp_pass.insert(0, p)
        if not silent:
            messagebox.showinfo("SMTP", "已从 smtp_config.ini 加载（若文件不存在则为空）。")

    def _on_save_smtp(self):
        u = self._smtp_user.get().strip()
        p = self._smtp_pass.get()
        if save_smtp_credentials(self._root(), u, p):
            messagebox.showinfo("SMTP", "已保存到 smtp_config.ini。")
        else:
            messagebox.showerror("SMTP", "保存失败，请检查工作根目录是否有效。")

    # ----- 签名 -----

    def _on_load_signature(self, silent=False):
        content = load_signature(self._root())
        self._sig_text.delete("1.0", tk.END)
        self._sig_text.insert("1.0", content)
        if not silent:
            messagebox.showinfo("签名", "已从 Signature.txt 加载。")

    def _on_save_signature(self):
        content = self._sig_text.get("1.0", tk.END)
        if save_signature(self._root(), content):
            messagebox.showinfo("签名", "已保存到 Signature.txt。")
        else:
            messagebox.showerror("签名", "保存失败，请检查工作根目录是否有效。")

    # ----- 项目创建 -----

    def _on_create_project(self):
        name = self._project_name.get()
        ok, msg = create_project_folders(self._root(), name)
        if ok:
            self._on_refresh_projects()
            messagebox.showinfo("项目管理", msg)
        else:
            messagebox.showerror("项目管理", msg)

    # ----- 日志 -----

    def _on_refresh_log(self):
        content = load_log_content(self._root())
        self._log_text.delete("1.0", tk.END)
        self._log_text.insert("1.0", content)

    # ----- 启动 -----

    def run(self):
        self._on_load_smtp(silent=True)
        self._on_load_signature(silent=True)
        self._on_refresh_projects()
        self._on_refresh_log()
        self.root.mainloop()


def main():
    app = App()
    app.run()


if __name__ == "__main__":
    main()
