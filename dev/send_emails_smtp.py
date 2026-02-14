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
import io
import logging
import shutil
import smtplib
from collections import deque
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time

# 确保 dev/ 目录在 sys.path 中以导入 config
_dev_dir = os.path.dirname(os.path.abspath(__file__))
if _dev_dir not in sys.path:
    sys.path.insert(0, _dev_dir)

from config import (
    LOG_FILENAME, CSV_FILENAME, SKIP_NAMES,
    load_smtp_config as _config_load_smtp,
    load_signature as _config_load_signature,
    signature_path,
)

# ---------------------------------------------------------------------------
# 限流配置（内置默认值，无需用户配置）
# ---------------------------------------------------------------------------
RATE_INITIAL_DELAY = 1.0    # 初始发送间隔（秒）
RATE_MAX_DELAY = 10.0       # 最大发送间隔（秒）
RATE_MIN_DELAY = 0.1        # 最小发送间隔（秒）
EMA_ALPHA = 0.3             # 速率平滑因子
COOLDOWN_SECONDS = 30.0     # 全局冷却时长（秒）
THRESHOLD_421 = 3           # 触发冷却的 421 错误次数阈值
WINDOW_421 = 60.0           # 421 滑动窗口（秒）
MAX_RETRIES = 3             # 最大重试次数


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
    从根目录下 smtp_config.ini 读取 SMTP 配置（委托给 config.py，增加日志）。
    @returns {dict|None}
    """
    cfg = _config_load_smtp(root_dir)
    if cfg is None:
        logging.error("SMTP 配置文件不存在或解析失败")
    return cfg


def read_email_addresses(root_dir):
    """
    从根目录下 EmailAddress.csv 读取供应商邮箱，按第一列分组，第三列用分号合并。
    自动检测编码：优先 UTF-8，失败则尝试 GBK（兼容中文 Windows Excel 导出）。
    @returns {dict} 供应商代码 -> 邮箱字符串(;分隔)
    """
    path = os.path.join(root_dir, CSV_FILENAME)
    if not os.path.isfile(path):
        logging.error(f"EmailAddress.csv 不存在: {path}")
        return {}

    content = None
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                content = f.read()
            break
        except (UnicodeDecodeError, LookupError):
            continue
    if content is None:
        logging.error(f"EmailAddress.csv 编码识别失败，请确认文件编码为 UTF-8 或 GBK: {path}")
        return {}

    try:
        by_code = {}
        reader = csv.reader(io.StringIO(content))
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
    except Exception as e:
        logging.error(f"读取 EmailAddress.csv 失败: {e}")
        return {}


def read_signature(root_dir):
    """从 Signature.txt 读取签名（委托给 config.py，增加日志）。"""
    content = _config_load_signature(root_dir)
    if not content:
        logging.error(f"Signature.txt 未找到或为空: {signature_path(root_dir)}")
    return content


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


def _build_message(from_addr, to_emails, cc_email, subject, body_html, attachment_paths):
    """
    构建 MIME 邮件消息。
    @returns {tuple} (MIMEMultipart, list[str]) 消息对象和收件人列表
    """
    to_list = [a.strip() for a in (to_emails or "").split(";") if a.strip()]
    if not to_list:
        return None, []

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
    return msg, recipients


def _create_smtp_connection(smtp_config):
    """
    创建并登录 SMTP 连接。
    @returns {smtplib.SMTP} 已登录的 SMTP 连接
    """
    host = smtp_config["host"]
    port = smtp_config["port"]
    use_ssl = smtp_config["use_ssl"]
    use_tls = smtp_config["use_tls"]
    username = smtp_config["username"]
    password = smtp_config["password"]

    if use_ssl:
        server = smtplib.SMTP_SSL(host, port, timeout=30)
    else:
        server = smtplib.SMTP(host, port, timeout=30)
    if use_tls and not use_ssl:
        server.starttls()
    if username or password:
        server.login(username, password)
    return server


def send_one_email(smtp_config, to_emails, cc_email, subject, body_html, attachment_paths, server=None):
    """
    发送一封邮件：To、CC、HTML 正文、附件。
    @param server {smtplib.SMTP|None} 可复用的 SMTP 连接；为 None 时自动创建并关闭
    @returns {tuple} (success: bool, error_code: Optional[int], server: smtplib.SMTP|None)
    """
    from_addr = smtp_config["username"]
    msg, recipients = _build_message(from_addr, to_emails, cc_email, subject, body_html, attachment_paths)
    if msg is None:
        return False, None, server

    own_server = False
    try:
        if server is None:
            server = _create_smtp_connection(smtp_config)
            own_server = True
        server.sendmail(from_addr, recipients, msg.as_string())
        if own_server:
            server.quit()
            server = None
        return True, None, server
    except smtplib.SMTPResponseException as e:
        code = getattr(e, 'smtp_code', None)
        logging.error(f"发送邮件失败: {e}")
        # 连接可能已损坏，标记为 None 以便下次重建
        try:
            if server:
                server.quit()
        except Exception:
            pass
        return False, code, None
    except Exception as e:
        logging.error(f"发送邮件失败: {e}")
        try:
            if server:
                server.quit()
        except Exception:
            pass
        return False, None, None


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

    def wait(self, stop_event=None):
        """等待当前延迟时长，支持 stop_event 提前中断。"""
        remaining = self.current_delay
        interval = 0.2
        while remaining > 0:
            if stop_event is not None and stop_event.is_set():
                return
            sleep_for = min(interval, remaining)
            time.sleep(sleep_for)
            remaining -= sleep_for


def send_one_email_with_rate_limiter(rate_limiter, smtp_config, to_emails, cc_email, subject, body_html, attachment_paths, server=None, stop_event=None):
    rate_limiter.wait(stop_event=stop_event)
    success, code, server = send_one_email(smtp_config, to_emails, cc_email, subject, body_html, attachment_paths, server=server)
    if success:
        rate_limiter.on_success()
    else:
        rate_limiter.on_failure()
    return success, code, server


def send_with_retries(rate_limiter, smtp_config, to_emails, cc_email, subject, body_html, attachment_paths, max_retries=3, server=None, stop_event=None):
    """Send with exponential backoff retries.

    Returns tuple (success: bool, last_error_code: Optional[int], server: Optional[smtplib.SMTP]).
    The first attempt respects rate_limiter; subsequent retries do not re-run the full rate_limiter.wait
    to avoid double waiting.
    """
    delay = 1.0
    last_code = None
    # first attempt with rate limiter
    success, code, server = send_one_email_with_rate_limiter(rate_limiter, smtp_config, to_emails, cc_email, subject, body_html, attachment_paths, server=server, stop_event=stop_event)
    if success:
        return True, None, server
    last_code = code
    attempt = 1
    while attempt < max_retries:
        # exponential backoff with jitter
        jitter = random.uniform(0.8, 1.2)
        time.sleep(delay * jitter)
        success, code, server = send_one_email(smtp_config, to_emails, cc_email, subject, body_html, attachment_paths, server=server)
        if success:
            return True, None, server
        last_code = code
        attempt += 1
        delay = min(delay * 2, 30)
    return False, last_code, server


def count_pending_tasks(root_dir, project_names=None):
    """
    预览待发送任务数量（不实际发送）。
    @returns {int} 待发送的供应商任务数
    """
    root_dir = (root_dir or "").strip() or os.getcwd()
    email_addresses = read_email_addresses(root_dir)

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

    count = 0
    for project_folder in to_process:
        pending_path = os.path.join(root_dir, project_folder, "待外发")
        if not os.path.isdir(pending_path):
            continue
        supplier_files = collect_supplier_files(pending_path)
        for supplier_code in supplier_files:
            to_email = email_addresses.get(supplier_code) or email_addresses.get(str(int(supplier_code)) if supplier_code.isdigit() else "")
            if to_email:
                count += 1
    return count


def main(
    root_dir,
    project_names=None,
    progress_callback=None,
    stop_event=None,
):
    """
    主流程：读取配置与数据，遍历项目，按供应商发送邮件，成功后移动文件并写日志。

    @param root_dir {str} TransportAgreement 根目录
    @param project_names {list|None} 仅处理这些项目文件夹名；None 表示处理所有以「项目」结尾的文件夹
    @returns {dict} {"success": int, "failed": int, "skipped": int, "cancelled": bool}
    """
    result = {"success": 0, "failed": 0, "skipped": 0, "cancelled": False}
    root_dir = (root_dir or "").strip() or os.getcwd()
    _setup_logging(root_dir)

    smtp_config = load_smtp_config(root_dir)
    if not smtp_config or not smtp_config.get("host") or not smtp_config.get("username"):
        logging.error("SMTP 配置缺失或无效，退出。")
        return result


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
        return result

    rate_limiter = RateLimiter(initial_delay=RATE_INITIAL_DELAY, max_delay=RATE_MAX_DELAY, min_delay=RATE_MIN_DELAY)
    start_time = time.time()
    completed = 0
    # exponential moving average for rate smoothing
    rate_ema = None
    # 421 tracking for global cooldown
    recent_421 = deque()
    cooldown_until = 0.0
    # 复用 SMTP 连接
    smtp_server = None

    try:
      for project_folder, supplier_code, files in tasks:
        # allow cancellation via stop_event
        if stop_event is not None and stop_event.is_set():
            logging.info("发送已被用户取消，提前退出。")
            result["cancelled"] = True
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

        # 确保 SMTP 连接可用
        if smtp_server is None:
            try:
                smtp_server = _create_smtp_connection(smtp_config)
            except Exception as e:
                logging.error(f"{log_prefix}SMTP 连接失败: {e}")
                result["failed"] += 1
                completed += 1
                continue

        logging.info(f"{log_prefix}正在发送邮件到 {to_email}，附件数量: {len(files)}")
        success, last_code, smtp_server = send_with_retries(
            rate_limiter,
            smtp_config,
            to_email,
            cc_email,
            subject,
            body_html,
            files,
            max_retries=MAX_RETRIES,
            server=smtp_server,
            stop_event=stop_event,
        )

        if success:
            result["success"] += 1
            logging.info(f"{log_prefix}成功 - 邮件发送成功。")
            for fp in files:
                try:
                    shutil.move(fp, os.path.join(os.path.join(root_dir, project_folder, "已外发"), os.path.basename(fp)))
                    logging.info(f"{log_prefix}已移动文件到「已外发」: {os.path.basename(fp)}")
                except Exception as e:
                    logging.error(f"{log_prefix}文件移动失败: {os.path.basename(fp)} - {e}")
        else:
            result["failed"] += 1
            logging.error(f"{log_prefix}失败 - 邮件发送失败。")
            # If last_code is 421, append timestamp and possibly trigger global cooldown
            now = time.time()
            if last_code == 421:
                recent_421.append(now)
                # remove old
                while recent_421 and now - recent_421[0] > WINDOW_421:
                    recent_421.popleft()
                if len(recent_421) >= THRESHOLD_421:
                    cooldown_until = now + COOLDOWN_SECONDS
                    logging.warning(f"检测到连续 {THRESHOLD_421} 次 421，触发全局冷却 {COOLDOWN_SECONDS}s")

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
                rate_ema = EMA_ALPHA * inst_rate + (1 - EMA_ALPHA) * rate_ema
            rate = rate_ema
            eta_seconds = (total - completed) / rate if rate > 0 else None
            percent = (completed / total) * 100.0
            try:
                progress_callback(percent, rate, eta_seconds, completed, total)
            except Exception:
                logging.exception("调用 progress_callback 时出错")

    finally:
        # 关闭 SMTP 连接
        if smtp_server is not None:
            try:
                smtp_server.quit()
            except Exception:
                pass

    return result


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
