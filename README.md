# Wgen

一个 Windows 图形界面小工具：选择原始 `.docx` 或 `.txt` 英文单词记录，调用 BigModel / 智谱中国版、Ox Alpha / Z.ai GLM 或 Kimi API 自动整理，并生成适合 A4 打印复习的 Word 单词本，同时生成可在 iPhone 上阅读的离线 HTML。

## 功能

- 支持 `.docx` Word 文档和 `.txt` 纯文本文件
- 默认支持 BigModel / 智谱中国版，新模型码为 `glm-5.3-flash`
- 保留海外 Ox Alpha / Z.ai GLM 作为可选模型服务
- 保留 Kimi / Moonshot 作为可选模型服务
- 调用模型 JSON Mode 输出结构化数据
- 自动生成三列表格 Word 单词本
- 同时生成手机阅读版 HTML，可用 iPhone Safari、“文件”App 或邮件附件打开
- 可选择整理完成后自动归档到个人主页仓库，并提交推送到 GitHub
- 保留模型原始 JSON 和结构化 JSON，便于核查
- 支持测试 API、刷新模型列表
- 支持将 API Key 保存到 Windows 凭据管理器

## 下载

普通用户请到仓库右侧或顶部的 **Releases** 页面下载最新版 EXE：

```text
Wgen.exe
```

下载后双击运行即可，不需要安装 Python，也不需要打开 VS Code。

## 运行源码版

如果你想直接运行源码，可以下载仓库代码后双击：


```bat
run_gui.cmd
```

脚本会自动创建 `.venv` 并安装运行依赖。

## 打包 EXE

双击：

```bat
build_exe.cmd
```

打包完成后，EXE 位于：

```text
dist\Wgen.exe
```

## 注意

- 调用模型 API 会从所选平台余额中按量扣费。
- 中国版请选择 `BigModel / 智谱中国版`，模型默认使用 `glm-5.3-flash`，接口地址为 `https://open.bigmodel.cn/api/paas/v4`。
- Ox Alpha 是 GLM-5.3-Flash 的匿名测试名；海外 Z.ai API 可选择 `Ox Alpha / Z.ai GLM`。
- 可用环境变量 `BIGMODEL_API_KEY` 预填智谱中国版 Key；海外 Z.ai 支持 `ZAI_API_KEY`；旧 Kimi 服务仍支持 `MOONSHOT_API_KEY`。
- 第一次建议先用较短的 `.txt` 或 `.docx` 测试。
- iPhone 阅读建议使用生成的“手机版 HTML”：可以发邮件给自己作为附件，也可以放到 iCloud Drive 后在 iPhone 的“文件”App 或 Safari 中打开。
- 如果启用个人主页归档，软件会把手机版 HTML 写入 personal-homepage 仓库的 `notes/ielts/wordbooks/`，更新收录页，并自动 `git commit` / `git push`。
- `.venv`、`build`、`dist`、EXE 和 zip 文件不应提交到 GitHub。
