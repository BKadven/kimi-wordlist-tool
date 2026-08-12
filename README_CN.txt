Kimi 雅思单词本整理器

一个 Windows 图形界面小工具：选择原始 .docx 或 .txt 英文单词记录，调用 Kimi API 自动整理，并生成适合 A4 打印复习的 Word 单词本。

主要功能
- 支持 .docx Word 文档和 .txt 纯文本文件。
- 自动读取完整整理规则。
- 调用 Kimi JSON Mode 输出结构化数据。
- 自动生成三列表格 Word 单词本。
- 保留 Kimi 原始 JSON 和结构化 JSON，便于核查。
- 可测试 API、刷新模型列表。
- 可选择是否启用深度思考；默认关闭以节省时间和费用。
- 可勾选在本机记住 Kimi API Key，保存到 Windows 凭据管理器，不写入源码目录。

运行源码版
1. 安装 Python。
2. 双击 run_gui.cmd。
3. 脚本会自动创建 .venv 并安装 requirements.txt 中的运行依赖。

打包 EXE
1. 双击 build_exe.cmd。
2. 脚本会自动安装打包依赖并调用 PyInstaller。
3. 打包完成后，EXE 位于 dist\Kimi雅思单词本整理器_太阳版.exe。

整理规则
软件首次运行会把默认规则复制到：
%APPDATA%\KimiWordlistTool\rules.txt

之后可以在软件中点击“打开整理规则”进行修改；点击“恢复默认规则”可还原内置规则。

注意
- 调用 Kimi API 会从你的开放平台余额中按量扣费。
- 第一次建议先用较短的 .txt 或 .docx 测试。
- .venv、build、dist、__pycache__、EXE 和 zip 文件不应提交到 GitHub。
