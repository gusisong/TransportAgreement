# -*- coding: utf-8 -*-
"""
运输协议邮件外发 — 图形界面入口。

提供：SMTP 账号维护、签名维护、项目文件夹创建、日志查看。
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import configparser
import threading

# 打包后 exe 所在目录为工作根目录；开发时脚本在 dev/ 下则工作根目录为项目根（上一级）
def _default_root_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.dirname(script_dir)
    except Exception:
        return os.getcwd()


# ---------------------------------------------------------------------------
# SMTP 配置读写
# ---------------------------------------------------------------------------

SMTP_SECTION = "smtp"
CONFIG_FILENAME = "smtp_config.ini"
SIGNATURE_FILENAME = "Signature.txt"
LOG_FILENAME = "email_smtp_log.log"


def _config_path(root_dir):
    # 优先查找 dev/ 目录下的配置文件，若不存在再回退到根目录
    dev_path = os.path.join(root_dir, "dev", CONFIG_FILENAME)
    if os.path.isfile(dev_path):
        return dev_path
    return os.path.join(root_dir, CONFIG_FILENAME)


def _signature_path(root_dir):
    # 优先查找 dev/ 目录下的签名文件，若不存在再回退到根目录
    dev_path = os.path.join(root_dir, "dev", SIGNATURE_FILENAME)
    if os.path.isfile(dev_path):
        return dev_path
    return os.path.join(root_dir, SIGNATURE_FILENAME)


def _log_path(root_dir):
    return os.path.join(root_dir, LOG_FILENAME)


def load_smtp_credentials(root_dir):
    """
    从根目录下的 smtp_config.ini 读取 username 和 password。
    @param root_dir {str} 工作根目录
    @returns {tuple} (username, password)，缺失则为 ("", "")
    """
    path = _config_path(root_dir)
    if not os.path.isfile(path):
        return "", ""
    try:
        cfg = configparser.ConfigParser()
        cfg.read(path, encoding="utf-8")
        if not cfg.has_section(SMTP_SECTION):
            return "", ""
        u = cfg.get(SMTP_SECTION, "username", fallback="")
        p = cfg.get(SMTP_SECTION, "password", fallback="")
        return (u.strip(), p.strip())
    except Exception:
        return "", ""


def save_smtp_credentials(root_dir, username, password):
    """
    将 username、password 写入根目录下的 smtp_config.ini，保留其余配置项；
    若文件不存在则写入默认 host/port/use_ssl/use_tls 以便 SMTP 脚本可用。
    @param root_dir {str} 工作根目录
    @param username {str} 登录账号
    @param password {str} 密码
    @returns {bool} 是否成功
    """
    path = _config_path(root_dir)
    try:
        cfg = configparser.ConfigParser()
        if os.path.isfile(path):
            cfg.read(path, encoding="utf-8")
        if not cfg.has_section(SMTP_SECTION):
            cfg.add_section(SMTP_SECTION)
        if not cfg.has_option(SMTP_SECTION, "host"):
            cfg.set(SMTP_SECTION, "host", "smtp.csvw.com")
        if not cfg.has_option(SMTP_SECTION, "port"):
            cfg.set(SMTP_SECTION, "port", "587")
        if not cfg.has_option(SMTP_SECTION, "use_ssl"):
            cfg.set(SMTP_SECTION, "use_ssl", "false")
        if not cfg.has_option(SMTP_SECTION, "use_tls"):
            cfg.set(SMTP_SECTION, "use_tls", "true")
        cfg.set(SMTP_SECTION, "username", username)
        cfg.set(SMTP_SECTION, "password", password)
        with open(path, "w", encoding="utf-8") as f:
            cfg.write(f)
        return True
    except Exception:
        return False


def load_signature(root_dir):
    """
    从根目录下的 Signature.txt 读取内容。
    @param root_dir {str} 工作根目录
    @returns {str} 签名内容，失败返回空字符串
    """
    path = _signature_path(root_dir)
    if not os.path.isfile(path):
        return ""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return ""


def save_signature(root_dir, content):
    """
    将内容写入根目录下的 Signature.txt。
    @param root_dir {str} 工作根目录
    @param content {str} 签名内容
    @returns {bool} 是否成功
    """
    path = _signature_path(root_dir)
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        return True
    except Exception:
        return False


def _normalize_project_folder_name(project_name):
    """
    若输入结尾不是「项目」，则自动补全「项目」；已有则不变。
    @param project_name {str} 用户输入
    @returns {str} 实际用于创建文件夹的名称
    """
    name = (project_name or "").strip()
    if not name:
        return ""
    if name.endswith("项目"):
        return name
    return name + "项目"


def create_project_folders(root_dir, project_name):
    """
    在根目录下创建 项目文件夹名/待外发 和 项目文件夹名/已外发。
    若用户输入结尾不是「项目」，则自动补全「项目」再创建（如 "AUDI E7X" -> "AUDI E7X项目"）。
    @param root_dir {str} 工作根目录
    @param project_name {str} 用户输入的项目名称
    @returns {tuple} (success: bool, message: str)
    """
    raw = (project_name or "").strip()
    if not raw:
        return False, "请输入项目名称。"
    name = _normalize_project_folder_name(raw)
    bad = ["已外发", "待外发"]
    if name in bad:
        return False, f"项目名称不能为「{name}」。"
    project_path = os.path.join(root_dir, name)
    pending = os.path.join(project_path, "待外发")
    sent = os.path.join(project_path, "已外发")
    try:
        os.makedirs(pending, exist_ok=True)
        os.makedirs(sent, exist_ok=True)
        return True, f"已创建：{pending}\n{sent}"
    except Exception as e:
        return False, str(e)


def load_log_content(root_dir):
    """
    读取根目录下 SMTP 日志文件内容。
    @param root_dir {str} 工作根目录
    @returns {str} 日志内容，不存在或失败返回提示文本
    """
    path = _log_path(root_dir)
    if not os.path.isfile(path):
        return f"[ 日志文件不存在: {path} ]"
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return f.read()
    except Exception as e:
        return f"[ 读取失败: {e} ]"


def list_project_folders(root_dir):
    """
    列出根目录下所有以「项目」结尾的子文件夹名称。
    @param root_dir {str} 工作根目录
    @returns {list} 文件夹名称列表，已排序
    """
    if not os.path.isdir(root_dir):
        return []
    out = []
    for name in os.listdir(root_dir):
        if name in ("已外发", "待外发"):
            continue
        path = os.path.join(root_dir, name)
        if os.path.isdir(path) and name.endswith("项目"):
            out.append(name)
    return sorted(out)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("运输协议邮件外发")
        self.root.minsize(520, 420)
        self._root_dir_value = _default_root_dir()

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
        self._proj_tab_index = 2  # 用于「开始批量发送」未选项目时切换到此页
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

        # 底部：包含进度条（左）、剩余时间（中）、开始批量发送按钮（右）
        bottom = ttk.Frame(self.root, padding="8 8")
        bottom.pack(fill=tk.X)

        # 左侧：进度条（带百分比覆盖），使用绿色样式
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

        # 在进度条上放置百分比标签（覆盖）
        self._percent_label = ttk.Label(progress_frame, text="0%", background="", anchor=tk.CENTER)
        # use place to overlay label on progressbar
        self._percent_label.place(in_=self._progress_bar, relx=0.5, rely=0.5, anchor=tk.CENTER)

        # 中间：预估剩余时间
        self._eta_label = ttk.Label(bottom, text="剩余时间: --", anchor=tk.CENTER)
        self._eta_label.pack(side=tk.LEFT, padx=12)

        # 右侧：开始批量发送与取消按钮
        right_frame = ttk.Frame(bottom)
        right_frame.pack(side=tk.RIGHT)

        self._btn_cancel = ttk.Button(right_frame, text="取消", command=self._on_cancel_send, state=tk.DISABLED)
        self._btn_cancel.pack(side=tk.RIGHT, padx=(6, 0))

        self._btn_send = ttk.Button(right_frame, text="开始批量发送", command=self._on_start_batch_send)
        self._btn_send.pack(side=tk.RIGHT)

    def _on_refresh_projects(self):
        """刷新「项目管理」中的项目列表（以「项目」结尾的文件夹）。"""
        for w in self._proj_check_frame.winfo_children():
            w.destroy()
        self._project_vars.clear()
        root = self._root()
        for name in list_project_folders(root):
            var = tk.BooleanVar(value=False)
            self._project_vars[name] = var
            cb = ttk.Checkbutton(
                self._proj_check_frame,
                text=name,
                variable=var,
            )
            cb.pack(anchor=tk.W)
        if not self._project_vars:
            ttk.Label(
                self._proj_check_frame,
                text="（当前根目录下没有以「项目」结尾的文件夹，请先创建）",
                foreground="gray",
            ).pack(anchor=tk.W)

    def _get_selected_projects(self):
        """返回用户勾选的项目文件夹名称列表。"""
        return [name for name, var in self._project_vars.items() if var.get()]

    def _on_start_batch_send(self):
        """开始批量发送：未选项目时提示并引导至项目管理；已选则调用 SMTP 发送。"""
        selected = self._get_selected_projects()
        if not selected:
            messagebox.showwarning(
                "未选择项目",
                "请先在「项目管理」选项卡中勾选要外发的项目，再点击「开始批量发送」。",
            )
            self._nb.select(self._proj_tab_index)
            return
        # run sending in a background thread
        def worker():
            try:
                _root = self._root()
                if not getattr(sys, "frozen", False):
                    _dev = os.path.dirname(os.path.abspath(__file__))
                    if _dev not in sys.path:
                        sys.path.insert(0, _dev)
                from send_emails_smtp import main as smtp_main
                # create or reuse stop event
                if not hasattr(self, "_stop_event") or self._stop_event is None:
                    self._stop_event = threading.Event()
                smtp_main(
                    _root,
                    project_names=selected,
                    progress_callback=self._progress_callback,
                    stop_event=self._stop_event,
                )
                # notify and refresh logs on main thread
                self.root.after(0, lambda: messagebox.showinfo("批量发送", "发送任务已执行，请到「日志」选项卡查看结果。"))
                self.root.after(0, self._on_refresh_log)
            except ImportError:
                self.root.after(0, lambda: messagebox.showinfo(
                    "提示",
                    "SMTP 发送模块未找到（send_emails_smtp.py），请先部署后再使用「开始批量发送」。",
                ))
            finally:
                # re-enable button and disable cancel
                self.root.after(0, lambda: self._btn_send.config(state=tk.NORMAL))
                self.root.after(0, lambda: self._btn_cancel.config(state=tk.DISABLED))

        # disable send button and enable cancel
        self._btn_send.config(state=tk.DISABLED)
        self._btn_cancel.config(state=tk.NORMAL)
        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _progress_callback(self, percent, rate, eta_seconds, completed, total):
        # called from worker thread; schedule UI update on main thread
        def cb():
            self._progress_var.set(percent)
            self._percent_label.config(text=f"{int(percent)}%")
            if eta_seconds is None:
                self._eta_label.config(text="剩余时间: --")
            else:
                # format seconds to H:MM:SS
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

    def _on_load_smtp(self):
        u, p = load_smtp_credentials(self._root())
        self._smtp_user.delete(0, tk.END)
        self._smtp_user.insert(0, u)
        self._smtp_pass.delete(0, tk.END)
        self._smtp_pass.insert(0, p)
        messagebox.showinfo("SMTP", "已从 smtp_config.ini 加载（若文件不存在则为空）。")

    def _on_save_smtp(self):
        u = self._smtp_user.get().strip()
        p = self._smtp_pass.get()
        if save_smtp_credentials(self._root(), u, p):
            messagebox.showinfo("SMTP", "已保存到 smtp_config.ini。")
        else:
            messagebox.showerror("SMTP", "保存失败，请检查工作根目录是否有效。")

    def _on_load_signature(self):
        content = load_signature(self._root())
        self._sig_text.delete("1.0", tk.END)
        self._sig_text.insert("1.0", content)
        messagebox.showinfo("签名", "已从 Signature.txt 加载。")

    def _on_save_signature(self):
        content = self._sig_text.get("1.0", tk.END)
        if save_signature(self._root(), content):
            messagebox.showinfo("签名", "已保存到 Signature.txt。")
        else:
            messagebox.showerror("签名", "保存失败，请检查工作根目录是否有效。")

    def _on_create_project(self):
        name = self._project_name.get()
        ok, msg = create_project_folders(self._root(), name)
        if ok:
            self._on_refresh_projects()
            messagebox.showinfo("项目管理", msg)
        else:
            messagebox.showerror("项目管理", msg)

    def _on_refresh_log(self):
        content = load_log_content(self._root())
        self._log_text.delete("1.0", tk.END)
        self._log_text.insert("1.0", content)

    def run(self):
        self._on_load_smtp()
        self._on_load_signature()
        self._on_refresh_projects()
        self._on_refresh_log()
        self.root.mainloop()


def main():
    app = App()
    app.run()


if __name__ == "__main__":
    main()
