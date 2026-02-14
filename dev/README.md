# dev — 开发与配置（用户无需关注）

本目录存放**所有脚本、文档与配置模板**，项目根目录仅保留用户维护的 **EmailAddress.csv**。

| 文件 | 说明 |
|------|------|
| **gui.py** | 图形界面入口；打包后用户通过 exe 启动。 |
| **send_emails_smtp.py** | SMTP 批量发送逻辑，由 GUI「开始批量发送」调用。 |
| **build_gui.bat** | 打包脚本。在**项目根目录**执行 `dev\build_gui.bat`，生成 dist\TransportAgreement_GUI\TransportAgreement_GUI.exe。 |
| **requirements-dev.txt** | 打包依赖（PyInstaller），`pip install -r dev/requirements-dev.txt`。 |
| **smtp_config.example.ini** | SMTP 配置模板。可复制到工作根目录并改名为 `smtp_config.ini` 后填写；或由 GUI 首次保存自动生成。 |
| **打包说明.md** | 打包步骤与使用方式。 |
| **需求文档_SMTP版.md** | SMTP 版需求说明。 |

- **运行（开发）**：在项目根目录执行 `python dev/gui.py`，工作根目录默认为项目根。
- **打包**：在项目根目录执行 `dev\build_gui.bat`。打包后用户只需双击 **TransportAgreement_GUI.exe** 即可使用。
