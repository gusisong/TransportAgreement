# -*- coding: utf-8 -*-
"""
公共配置模块 — 路径查找、常量、配置文件读写。

gui.py 和 send_emails_smtp.py 共享此模块，消除重复代码。
"""

import os
import sys
import configparser


# ---------------------------------------------------------------------------
# 文件名常量
# ---------------------------------------------------------------------------
SMTP_SECTION = "smtp"
CONFIG_FILENAME = "smtp_config.ini"
SIGNATURE_FILENAME = "Signature.txt"
LOG_FILENAME = "email_smtp_log.log"
CSV_FILENAME = "EmailAddress.csv"
SKIP_NAMES = ("已外发", "待外发")


# ---------------------------------------------------------------------------
# 路径查找（优先 dev/，回退根目录）
# ---------------------------------------------------------------------------

def default_root_dir():
    """
    工作根目录自动检测：
    - 打包后（frozen）→ exe 所在目录
    - 开发时 → 脚本 dev/ 的上一级
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.dirname(script_dir)
    except Exception:
        return os.getcwd()


def config_path(root_dir):
    """优先查找 dev/ 目录下的配置文件，若不存在再回退到根目录。"""
    dev_path = os.path.join(root_dir, "dev", CONFIG_FILENAME)
    if os.path.isfile(dev_path):
        return dev_path
    return os.path.join(root_dir, CONFIG_FILENAME)


def signature_path(root_dir):
    """优先查找 dev/ 目录下的签名文件，若不存在再回退到根目录。"""
    dev_path = os.path.join(root_dir, "dev", SIGNATURE_FILENAME)
    if os.path.isfile(dev_path):
        return dev_path
    return os.path.join(root_dir, SIGNATURE_FILENAME)


def log_path(root_dir):
    return os.path.join(root_dir, LOG_FILENAME)


# ---------------------------------------------------------------------------
# SMTP 配置读写
# ---------------------------------------------------------------------------

def load_smtp_config(root_dir):
    """
    从根目录下 smtp_config.ini 读取 SMTP 配置。
    @returns {dict|None} host, port, use_ssl, use_tls, username, password
    """
    path = config_path(root_dir)
    if not os.path.isfile(path):
        return None
    try:
        cfg = configparser.ConfigParser()
        cfg.read(path, encoding="utf-8")
        if not cfg.has_section(SMTP_SECTION):
            return None
        s = cfg[SMTP_SECTION]
        return {
            "host": s.get("host", "").strip(),
            "port": cfg.getint(SMTP_SECTION, "port", fallback=587),
            "use_ssl": cfg.getboolean(SMTP_SECTION, "use_ssl", fallback=False),
            "use_tls": cfg.getboolean(SMTP_SECTION, "use_tls", fallback=True),
            "username": s.get("username", "").strip(),
            "password": s.get("password", "").strip(),
        }
    except Exception:
        return None


def load_smtp_credentials(root_dir):
    """
    读取 smtp_config.ini 的 username, password。
    @returns {tuple} (username, password)
    """
    cfg = load_smtp_config(root_dir)
    if cfg is None:
        return "", ""
    return cfg.get("username", ""), cfg.get("password", "")


def save_smtp_credentials(root_dir, username, password):
    """
    将 username、password 写入 smtp_config.ini，保留其余配置项。
    @returns {bool} 是否成功
    """
    path = config_path(root_dir)
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


# ---------------------------------------------------------------------------
# 签名读写
# ---------------------------------------------------------------------------

def load_signature(root_dir):
    """读取 Signature.txt 内容。"""
    path = signature_path(root_dir)
    if not os.path.isfile(path):
        return ""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return ""


def save_signature(root_dir, content):
    """写入 Signature.txt。"""
    path = signature_path(root_dir)
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 日志读取
# ---------------------------------------------------------------------------

def load_log_content(root_dir, max_lines=500):
    """
    读取日志文件最后 N 行。
    @returns {str} 日志内容
    """
    path = log_path(root_dir)
    if not os.path.isfile(path):
        return f"[ 日志文件不存在: {path} ]"
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        if len(lines) > max_lines:
            lines = lines[-max_lines:]
            return f"[ 仅显示最后 {max_lines} 行 ]…\n" + "".join(lines)
        return "".join(lines)
    except Exception as e:
        return f"[ 读取失败: {e} ]"


# ---------------------------------------------------------------------------
# 项目文件夹
# ---------------------------------------------------------------------------

def normalize_project_folder_name(project_name):
    """若输入结尾不是「项目」，则自动补全。"""
    name = (project_name or "").strip()
    if not name:
        return ""
    if name.endswith("项目"):
        return name
    return name + "项目"


def create_project_folders(root_dir, project_name):
    """
    在根目录下创建 项目文件夹名/待外发 和 项目文件夹名/已外发。
    @returns {tuple} (success: bool, message: str)
    """
    raw = (project_name or "").strip()
    if not raw:
        return False, "请输入项目名称。"
    name = normalize_project_folder_name(raw)
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


def list_project_folders(root_dir):
    """列出根目录下所有以「项目」结尾的子文件夹名称。"""
    if not os.path.isdir(root_dir):
        return []
    out = []
    for name in os.listdir(root_dir):
        if name in SKIP_NAMES:
            continue
        path = os.path.join(root_dir, name)
        if os.path.isdir(path) and name.endswith("项目"):
            out.append(name)
    return sorted(out)
