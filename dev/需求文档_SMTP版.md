# 运输协议邮件批量外发 — SMTP 版需求文档

> 本文档描述基于 **SMTP** 的邮件外发实现需求，与 Outlook 版（`send_emails.py`）在业务规则上保持一致，仅在发信方式、运行环境与配置上不同。实现见 **send_emails_smtp.py**（本目录），配置示例见 **smtp_config.example.ini**（本目录）。

---

## 1. 功能概述

在指定根目录（TransportAgreement）下，按**项目文件夹**遍历，从各项目的 **待外发** 目录中识别按供应商命名的 Excel 文件，通过 **SMTP** 向 `EmailAddress.csv` 中对应供应商邮箱批量发送邮件（主题、正文、签名、附件按规则生成）；**每封邮件同时抄送（CC）至 SMTP 登录账号**；**发送成功后**将对应文件从 **待外发** 移动到 **已外发**，并写入 **SMTP 专用日志文件**。

---

## 2. 运行环境与依赖

| 项目 | 说明 |
|------|------|
| 运行环境 | **跨平台**（含 Linux），不依赖 Windows 或 Outlook |
| Python 库 | **仅使用标准库**：`smtplib`、`email`、`os`、`logging`、`shutil`、`csv`、`configparser`；不依赖 pandas 等第三方库 |
| 前置条件 | 本机可访问企业 SMTP 服务器（如 `smtp.csvw.com`），且已具备有效的 SMTP 配置文件 |

**说明**：`EmailAddress.csv` 使用标准库 `csv` 解析，保证列顺序、编码与现有格式一致，便于在无第三方库环境下运行。

---

## 3. SMTP 配置

### 3.1 配置文件

- **格式**：**INI**，使用 Python 标准库 `configparser` 解析。
- **文件名**：`smtp_config.ini`。
- **位置**：放在 TransportAgreement **根目录**下（即用户工作根目录）。
- **用途**：存放 SMTP 服务器与认证信息，由**用户自行创建与维护**，脚本不包含任何账号密码。
- **安全**：建议将 `smtp_config.ini` 加入 `.gitignore`；示例见本目录 **smtp_config.example.ini**。

### 3.2 配置项与 INI 结构

- **节（section）**：统一放在一节下，例如 `[smtp]`。
- **配置项**：host、port、use_ssl、use_tls、username、password。推荐端口 587 + STARTTLS。

### 3.3 发件人

- **发件人**：**仅使用邮箱地址**，与 SMTP 登录账号一致；不配置显示名。

---

## 4. 目录与文件结构

与《需求文档.md》一致：

- **根目录**：即“工作根目录”；用户仅需在根目录维护 **EmailAddress.csv**；smtp_config.ini、Signature.txt 可由 GUI 在工作根目录下保存生成。
- **根目录下**需存在：EmailAddress.csv；Signature.txt、smtp_config.ini 可在首次使用时通过 GUI 保存生成。
 - **根目录下**需存在：`EmailAddress.csv`；默认情况下 `smtp_config.ini` 与 `Signature.txt` 存放在 `dev/` 子目录中以减少根目录杂乱。GUI 在启动时会优先从 `dev/` 查找这两个文件，若未找到则回退到根目录以兼容旧版布局。
- **项目文件夹**：根目录下以「项目」结尾的子文件夹，各含「待外发」「已外发」。
- 脚本与文档位于 **dev/**，用户无需关注。

---

## 5. 数据文件规范

- **EmailAddress.csv**：UTF-8；至少 3 列（供应商代码、供应商名称、邮箱）；同一供应商多行邮箱按供应商代码分组，合并为分号 `;` 分隔；收件人查找键为 5 位数字供应商代码。
- **Signature.txt**：UTF-8；邮件签名全文；正文换行在邮件中转为 HTML `<br>`。

---

## 6. 待外发文件识别规则

- 仅处理各项目下 **待外发** 文件夹中的 `.xlsx` 文件；文件名按 `_` 分割，**倒数第二段**为供应商代码且须为**恰好 5 位数字**；同一项目、同一供应商代码下的文件合并为一封邮件的多个附件。

---

## 7. 邮件发送规则

- **To**：从 EmailAddress.csv 按供应商代码查找，全部放入 To。
- **CC**：每封邮件均将 SMTP 登录账号加入 CC。
- **主题**：`{项目名}零件供货方式确认_{供应商代码}`。
- **正文与签名**：固定正文 + Signature.txt 全文，HTML 发送（`\n`→`<br>`）。
- **附件**：该供应商在该项目「待外发」下识别到的所有 xlsx。

---

## 8. 文件迁移规则

- 仅当该封邮件 SMTP 发送成功后，才将对应附件从「待外发」移动到「已外发」；移动失败记 ERROR，其余照常移动。

---

## 9. 日志

- 根目录下 `email_smtp_log.log`，UTF-8；记录配置/发送/迁移等。

---

## 10. 主流程与入口

- **main(root_dir, project_names=None)**：project_names 为 None 时处理所有以「项目」结尾的文件夹，为列表时仅处理列表中项目。开发时运行 `python dev/gui.py`，默认 root_dir 为项目根。

---

## 11. 异常与边界行为

- SMTP 配置缺失/无效→ERROR 退出；无邮箱→WARNING 跳过；发送失败→ERROR 不迁移；发送成功、某文件移动失败→ERROR，其余照常移动。

---

## 12. 与 Outlook 版差异

- 发信方式：smtplib+email；配置：INI；抄送：每封 CC 登录账号；日志：email_smtp_log.log；脚本：send_emails_smtp.py。

---

## 13. 已确定项

- INI 配置；发件人仅邮箱；CSV 用标准库；脚本 send_emails_smtp.py，日志 email_smtp_log.log。

---

## 14. 图形界面（GUI）

- 入口 gui.py；工作根目录可选（打包后默认 exe 目录）。**项目管理**：展示以「项目」结尾的文件夹并勾选；新建项目可自动补全「项目」；创建成功后自动刷新列表；未选项目时「开始批量发送」提示并引导至项目管理。**开始批量发送**：底部居中；已选项目则执行批量发送。

本文档与当前实现（send_emails_smtp.py、gui.py）保持一致；需求变更时请同步更新本文档。

### 自适应限流策略

1. **单线程/单连接发送**：
   - 所有邮件通过单一线程和单一SMTP连接发送，避免并发导致的限流问题。

2. **失败即减速，成功可微加速**：
   - 遇到限流错误（如421）时，立即大幅降低发送速率（延迟时间翻倍）。
   - 连续成功发送3次时，小幅提速（延迟时间减少10%）。

3. **错误重试带退避**：
   - 发送失败的邮件采用指数退避策略（1秒、2秒、4秒……），最大重试次数为3次。

4. **全局状态管理**：
   - 使用全局速率控制器（RateLimiter）管理发送速率，所有任务共享同一控制器。

### 配置参数
- **初始延迟时间**：1秒
- **最大延迟时间**：10秒
- **最小延迟时间**：0.1秒
- **连续成功提速阈值**：3次
### 示例日志
```
2026-02-13 10:00:00 INFO Email sent to user@example.com
2026-02-13 10:01:02 INFO Retrying email to user2@example.com (Attempt 2)
```

### 15. 新增功能

- **功能**：
  - 显示当前发送进度（按供应商号计算百分比）。
   - 显示预估剩余时间（放置于进度条右侧）。
- **实现**：
  - 使用 `ttk.Progressbar` 实现进度条。
   - 使用 `ttk.Label` 在进度条中央显示百分比覆盖文本，`ttk.Progressbar` 使用绿色主题。  
   - 进度更新通过 `send_emails_smtp.main` 提供的回调接口驱动（见下文）。
- **位置**：GUI 新增「限流配置」选项卡。
- **功能**：
  - 用户可自定义最大发送速率（邮件/秒）。
  - 用户可自定义最小间隔时间（秒）。
- **实现**：
  - 配置保存后即时生效。

### 16. 进度回调接口

- **目的**：GUI 需要在邮件外发过程中实时显示进度（按供应商号）、发送速率与预估剩余时间。为此，`send_emails_smtp.main` 提供一个可选的进度回调参数 `progress_callback`。
- **函数签名**：`progress_callback(percent: float, rate: float, eta_seconds: Optional[float], completed: int, total: int) -> None`。
   - `percent`：0-100 的浮点数表示完成百分比。
   - `rate`：当前平均发送速率（单位：邮件/秒）。
   - `eta_seconds`：估计剩余秒数，若无法估计则为 `None`。
   - `completed` / `total`：已完成与总任务数（按供应商号计数）。
- **调用时机**：每处理完一个供应商号（成功或失败后）调用一次。

```python
def my_progress(percent, rate, eta_seconds, completed, total):
      # 在 GUI 线程中通过安全调度更新 UI
      pass

send_emails_smtp.main(root_dir, project_names=[...], progress_callback=my_progress)
```

### 19. 发送失败后的文件处理

- 若对某个供应商的所有重试均失败，脚本会将该供应商对应的附件文件从该项目的 `待外发/` 目录移动到该项目目录下的 `failed/` 子目录，便于人工后续处理与排查。日志中会记录移动操作与任何移动失败的错误信息。
