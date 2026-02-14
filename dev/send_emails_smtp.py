# -*- coding: utf-8 -*-
"""
运输协议邮件批量外发 — SMTP 版。

基于需求文档_SMTP版.md：从根目录按项目扫描待外发 xlsx，按供应商聚合后经 SMTP 发送，
每封抄送登录账号，成功后移动文件到已外发，并写日志。
仅使用标准库：smtplib、email、csv、configparser、logging、shutil、os。
"""

import os
import sys
import csv
import logging
import shutil
import smtplib
import configparser
from collections import deque
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time

LOG_FILENAME = "email_smtp_log.log"
CONFIG_FILENAME = "smtp_config.ini"
SMTP_SECTION = "smtp"
CSV_FILENAME = "EmailAddress.csv"
SIGNATURE_FILENAME = "Signature.txt"
SKIP_NAMES = ("已外发", "待外发")


def _setup_logging(root_dir):
    """配置日志：写入根目录下 email_smtp_log.log，UTF-8。"""
    log_path = os.path.join(root_dir, LOG_FILENAME)
    handler = logging.FileHandler(log_path, mode="a", encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    if root.handlers:
        root.handlers.clear()
    root.addHandler(handler)
    return log_path


def load_smtp_config(root_dir):
    """
    从根目录下 smtp_config.ini 读取 SMTP 配置。
    @param root_dir {str} 根目录
    @returns {dict} host, port, use_ssl, use_tls, username, password；失败返回 None
    """
    # 优先在 dev/ 子目录查找配置文件（便于把用户可见文件留在根目录）
    path = os.path.join(root_dir, "dev", CONFIG_FILENAME)
    if not os.path.isfile(path):
        path = os.path.join(root_dir, CONFIG_FILENAME)
    if not os.path.isfile(path):
        logging.error(f"SMTP 配置文件不存在: {path}")
        return None
    try:
        cfg = configparser.ConfigParser()
        cfg.read(path, encoding="utf-8")
        if not cfg.has_section(SMTP_SECTION):
            logging.error("smtp_config.ini 中缺少 [smtp] 节")
            return None
        s = cfg[SMTP_SECTION]
        port = cfg.getint(SMTP_SECTION, "port", fallback=587)
        use_ssl = cfg.getboolean(SMTP_SECTION, "use_ssl", fallback=False)
        use_tls = cfg.getboolean(SMTP_SECTION, "use_tls", fallback=True)
        return {
            "host": s.get("host", "").strip(),
            "port": port,
            "use_ssl": use_ssl,
            "use_tls": use_tls,
            "username": s.get("username", "").strip(),
            "password": s.get("password", "").strip(),
        }
    except Exception as e:
        logging.error(f"读取或解析 SMTP 配置失败: {e}")
        return None


def read_email_addresses(root_dir):
    """
    从根目录下 EmailAddress.csv 读取供应商邮箱，按第一列分组，第三列用分号合并。
    @param root_dir {str} 根目录
    @returns {dict} 供应商代码( str ) -> 邮箱字符串( 多个用 ; 分隔 )
    """
    path = os.path.join(root_dir, CSV_FILENAME)
    if not os.path.isfile(path):
        logging.error(f"EmailAddress.csv 不存在: {path}")
        return {}
    try:
        by_code = {}
        with open(path, "r", encoding="utf-8", newline="") as f:
            reader = csv.reader(f)
            header = next(reader, None)
            for row in reader:
                if len(row) < 3:
                    continue
                code, _, email = row[0].strip(), row[1].strip(), row[2].strip()
                if not code or not email:
                    continue
                if code not in by_code:
                    by_code[code] = []
                by_code[code].append(email)
        return {k: ";".join(v) for k, v in by_code.items()}
    except FileNotFoundError:
        logging.error(f"EmailAddress.csv 未找到: {path}")
        return {}
    except Exception as e:
        logging.error(f"读取 EmailAddress.csv 失败: {e}")
        return {}


def read_signature(root_dir):
    """从根目录下 Signature.txt 读取签名内容。"""
    # 优先从 dev/ 子目录读取 Signature.txt
    path = os.path.join(root_dir, "dev", SIGNATURE_FILENAME)
    if not os.path.isfile(path):
        path = os.path.join(root_dir, SIGNATURE_FILENAME)
    if not os.path.isfile(path):
        logging.error(f"Signature.txt 未找到: {path}")
        return ""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        logging.error(f"读取 Signature.txt 失败: {e}")
        return ""


def collect_supplier_files(pending_path):
    """
    扫描待外发目录下 .xlsx，按供应商代码（文件名 _ 分割倒数第二段且为 5 位数字）聚合。
    @param pending_path {str} 待外发目录
    @returns {dict} 供应商代码(str) -> [文件路径, ...]
    """
    result = {}
    if not os.path.isdir(pending_path):
        return result
    for filename in os.listdir(pending_path):
        if not filename.endswith(".xlsx") or "_" not in filename:
            continue
        parts = filename.split("_")
        if len(parts) < 3:
            logging.warning(f"文件命名不符合约定，跳过: {filename}")
            continue
        code = parts[-2]
        if not (code.isdigit() and len(code) == 5):
            logging.warning(f"文件供应商代码格式不正确(非5位数字)，跳过: {filename}")
            continue
        path = os.path.join(pending_path, filename)
        result.setdefault(code, []).append(path)
    return result


def send_one_email(smtp_config, to_emails, cc_email, subject, body_html, attachment_paths):
    """
    发送一封邮件：To、CC、HTML 正文、附件。
    @returns {bool} 是否成功
    """
    host = smtp_config["host"]
    port = smtp_config["port"]
    use_ssl = smtp_config["use_ssl"]
    use_tls = smtp_config["use_tls"]
    username = smtp_config["username"]
    password = smtp_config["password"]
    from_addr = username

    to_list = [a.strip() for a in (to_emails or "").split(";") if a.strip()]
    if not to_list:
        return False

    msg = MIMEMultipart()
    msg["From"] = from_addr
    msg["To"] = "; ".join(to_list)
    msg["Cc"] = cc_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    for path in attachment_paths or []:
        if not os.path.isfile(path):
            continue
        with open(path, "rb") as fp:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(fp.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment", filename=os.path.basename(path))
        msg.attach(part)

    recipients = to_list + [cc_email] if cc_email else to_list
    try:
        if use_ssl:
            server = smtplib.SMTP_SSL(host, port)
        else:
            server = smtplib.SMTP(host, port)
        if use_tls and not use_ssl:
            server.starttls()
        if username or password:
            server.login(username, password)
        server.sendmail(from_addr, recipients, msg.as_string())
        server.quit()
        return True, None
    except smtplib.SMTPResponseException as e:
        # SMTP server response with code (e.g., 421)
        code = getattr(e, 'smtp_code', None)
        logging.error(f"发送邮件失败: {e}")
        return False, code
    except Exception as e:
        logging.error(f"发送邮件失败: {e}")
        return False, None


class RateLimiter:
    def __init__(self, initial_delay=1.0, max_delay=10.0, min_delay=0.1):
        self.current_delay = initial_delay
        self.max_delay = max_delay
        self.min_delay = min_delay
        self.success_count = 0

    def on_success(self):
        self.success_count += 1
        if self.success_count >= 3:  # 连续成功3次
            self.current_delay = max(self.min_delay, self.current_delay * 0.9)  # 小幅提速
            self.success_count = 0

    def on_failure(self):
        self.current_delay = min(self.max_delay, self.current_delay * 2)  # 大幅减速
        self.success_count = 0

    def wait(self):
        # Sleep in small increments to allow responsive cancellation checks
        remaining = self.current_delay
        interval = 0.2
        while remaining > 0:
            sleep_for = min(interval, remaining)
            time.sleep(sleep_for)
            remaining -= sleep_for


def send_one_email_with_rate_limiter(rate_limiter, smtp_config, to_emails, cc_email, subject, body_html, attachment_paths):
    rate_limiter.wait()
    success, code = send_one_email(smtp_config, to_emails, cc_email, subject, body_html, attachment_paths)
    if success:
        rate_limiter.on_success()
    else:
        rate_limiter.on_failure()
    return success, code


def send_with_retries(rate_limiter, smtp_config, to_emails, cc_email, subject, body_html, attachment_paths, max_retries=3):
    """Send with exponential backoff retries.

    Returns tuple (success: bool, last_error_code: Optional[int]).
    The first attempt respects rate_limiter; subsequent retries do not re-run the full rate_limiter.wait
    to avoid double waiting.
    """
    attempt = 0
    delay = 1.0
    last_code = None
    # first attempt with rate limiter
    success, code = send_one_email_with_rate_limiter(rate_limiter, smtp_config, to_emails, cc_email, subject, body_html, attachment_paths)
    if success:
        return True, None
    last_code = code
    attempt = 1
    while attempt < max_retries:
        # exponential backoff with jitter
        jitter = random.uniform(0.8, 1.2)
        time.sleep(delay * jitter)
        success, code = send_one_email(smtp_config, to_emails, cc_email, subject, body_html, attachment_paths)
        if success:
            return True, None
        last_code = code
        attempt += 1
        delay = min(delay * 2, 30)
    return False, last_code


def main(
    root_dir,
    project_names=None,
    progress_callback=None,
    stop_event=None,
    max_retries=3,
    initial_delay=1.0,
    min_delay=0.1,
    max_delay=10.0,
    cooldown_seconds=30,
    _421_threshold=3,
    _421_window=60.0,
    ema_alpha=0.3,
):
    """
    主流程：读取配置与数据，遍历项目，按供应商发送邮件，成功后移动文件并写日志。

    @param root_dir {str} TransportAgreement 根目录
    @param project_names {list|None} 仅处理这些项目文件夹名；None 表示处理所有以「项目」结尾的文件夹
    """
    root_dir = (root_dir or "").strip() or os.getcwd()
    _setup_logging(root_dir)

    smtp_config = load_smtp_config(root_dir)
    if not smtp_config or not smtp_config.get("host") or not smtp_config.get("username"):
        logging.error("SMTP 配置缺失或无效，退出。")
        return

    # 如果调用方（例如直接脚本运行）没有显式传入高级参数，
    # 则尝试从 smtp_config.ini 的 [smtp] 节读取这些值作为默认覆盖。
    try:
        cfg = configparser.ConfigParser()
        cfg_path = os.path.join(root_dir, "dev", CONFIG_FILENAME)
        if not os.path.isfile(cfg_path):
            cfg_path = os.path.join(root_dir, CONFIG_FILENAME)
        if os.path.isfile(cfg_path):
            cfg.read(cfg_path, encoding="utf-8")
            if cfg.has_section(SMTP_SECTION):
                s = cfg[SMTP_SECTION]
                # Only override when the parameter still equals its function default
                try:
                    if ema_alpha == 0.3:
                        ema_alpha = float(s.get("ema_alpha", ema_alpha))
                except Exception:
                    pass
                try:
                    if _421_threshold == 3:
                        _421_threshold = int(s.get("_421_threshold", _421_threshold))
                except Exception:
                    pass
                try:
                    if _421_window == 60.0:
                        _421_window = float(s.get("_421_window", _421_window))
                except Exception:
                    pass
                try:
                    if cooldown_seconds == 30:
                        cooldown_seconds = float(s.get("cooldown_seconds", cooldown_seconds))
                except Exception:
                    pass
                try:
                    if initial_delay == 1.0:
                        initial_delay = float(s.get("initial_delay", initial_delay))
                except Exception:
                    pass
                try:
                    if min_delay == 0.1:
                        min_delay = float(s.get("min_delay", min_delay))
                except Exception:
                    pass
                try:
                    if max_delay == 10.0:
                        max_delay = float(s.get("max_delay", max_delay))
                except Exception:
                    pass
    except Exception:
        # 不应阻塞主流程，忽略配置读取错误
        pass

    email_addresses = read_email_addresses(root_dir)
    signature = read_signature(root_dir)
    cc_email = smtp_config["username"]

    # 确定要处理的项目文件夹列表
    all_entries = []
    for name in os.listdir(root_dir):
        if name in SKIP_NAMES:
            continue
        path = os.path.join(root_dir, name)
        if os.path.isdir(path) and name.endswith("项目"):
            all_entries.append(name)
    if project_names:
        project_names_set = set(project_names)
        to_process = [n for n in sorted(all_entries) if n in project_names_set]
    else:
        to_process = sorted(all_entries)

    # Build task list (颗粒度：供应商号)，只包含有目标邮箱的任务
    tasks = []  # list of (project_folder, supplier_code, files)
    for project_folder in to_process:
        project_path = os.path.join(root_dir, project_folder)
        pending_path = os.path.join(project_path, "待外发")
        sent_path = os.path.join(project_path, "已外发")

        if not os.path.isdir(pending_path) or not os.path.isdir(sent_path):
            logging.warning(f"项目文件夹 '{project_folder}' 缺少「待外发」或「已外发」，跳过。")
            continue

        supplier_files = collect_supplier_files(pending_path)
        for supplier_code, files in supplier_files.items():
            to_email = email_addresses.get(supplier_code) or email_addresses.get(str(int(supplier_code)) if supplier_code.isdigit() else "")
            if not to_email:
                logging.warning(f"供应商 {supplier_code} ({project_folder}): 跳过 - 没有对应的邮箱地址。")
                continue
            tasks.append((project_folder, supplier_code, files))

    total = len(tasks)
    if total == 0:
        logging.info("没有需要发送的任务。")
        return

    rate_limiter = RateLimiter(initial_delay=initial_delay, max_delay=max_delay, min_delay=min_delay)
    start_time = time.time()
    completed = 0
    # exponential moving average for rate smoothing
    rate_ema = None
    # 421 tracking for global cooldown
    recent_421 = deque()
    cooldown_until = 0.0

    for project_folder, supplier_code, files in tasks:
        # allow cancellation via stop_event
        if stop_event is not None and getattr(stop_event, "is_set", lambda: False)():
            logging.info("发送已被用户取消，提前退出。")
            break
        to_email = email_addresses.get(supplier_code) or email_addresses.get(str(int(supplier_code)) if supplier_code.isdigit() else "")
        log_prefix = f"供应商 {supplier_code} ({project_folder}): "

        subject = f"{project_folder}零件供货方式确认_{supplier_code}"
        body_plain = (
            f"供应商，你好：\n\n"
            f"附件是{project_folder}零件《供货方式确认表》，烦请核对信息\n\n"
            f"如无问题，请在三个工作日内签字盖章回传，谢谢！\n\n"
            f"{signature}"
        )
        body_html = body_plain.replace("\n", "<br>")

        # check global cooldown
        now = time.time()
        if now < cooldown_until:
            wait_for = cooldown_until - now
            logging.info(f"全局冷却中，等待 {wait_for:.1f}s 后继续发送")
            time.sleep(wait_for)

        logging.info(f"{log_prefix}正在发送邮件到 {to_email}，附件数量: {len(files)}")
        success, last_code = send_with_retries(
            rate_limiter,
            smtp_config,
            to_email,
            cc_email,
            subject,
            body_html,
            files,
            max_retries=max_retries,
        )

        if success:
            logging.info(f"{log_prefix}成功 - 邮件发送成功。")
            for fp in files:
                try:
                    shutil.move(fp, os.path.join(os.path.join(root_dir, project_folder, "已外发"), os.path.basename(fp)))
                    logging.info(f"{log_prefix}已移动文件到「已外发」: {os.path.basename(fp)}")
                except Exception as e:
                    logging.error(f"{log_prefix}文件移动失败: {os.path.basename(fp)} - {e}")
        else:
            logging.error(f"{log_prefix}失败 - 邮件发送失败。")
            # If last_code is 421, append timestamp and possibly trigger global cooldown
            now = time.time()
            if last_code == 421:
                recent_421.append(now)
                # remove old
                while recent_421 and now - recent_421[0] > _421_window:
                    recent_421.popleft()
                if len(recent_421) >= _421_threshold:
                    cooldown_until = now + cooldown_seconds
                    logging.warning(f"检测到连续 {_421_threshold} 次 421，触发全局冷却 {cooldown_seconds}s")

            # 重试全部失败后，将这些文件移动到项目下的 failed/ 以便人工处理
            failed_dir = os.path.join(root_dir, project_folder, "failed")
            try:
                os.makedirs(failed_dir, exist_ok=True)
                for fp in files:
                    try:
                        shutil.move(fp, os.path.join(failed_dir, os.path.basename(fp)))
                        logging.info(f"{log_prefix}已移动失败文件到 failed/: {os.path.basename(fp)}")
                    except Exception as e:
                        logging.error(f"{log_prefix}failed 移动失败: {os.path.basename(fp)} - {e}")
            except Exception as e:
                logging.error(f"{log_prefix}创建 failed 目录失败: {e}")

        completed += 1
        # report progress via callback if provided; smooth rate with EMA
        if progress_callback:
            elapsed = time.time() - start_time
            inst_rate = completed / elapsed if elapsed > 0 else 0.0
            if rate_ema is None:
                rate_ema = inst_rate
            else:
                rate_ema = ema_alpha * inst_rate + (1 - ema_alpha) * rate_ema
            rate = rate_ema
            eta_seconds = (total - completed) / rate if rate > 0 else None
            percent = (completed / total) * 100.0
            try:
                progress_callback(percent, rate, eta_seconds, completed, total)
            except Exception:
                logging.exception("调用 progress_callback 时出错")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        root_directory = sys.argv[1]
    else:
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            root_directory = os.path.dirname(script_dir)
        except Exception:
            root_directory = os.getcwd()
    main(root_directory, project_names=None)
