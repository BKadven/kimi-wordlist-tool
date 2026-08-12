# Kimi 雅思单词本整理器

一个 Windows 图形界面小工具：选择原始 `.docx` 或 `.txt` 英文单词记录，调用 Kimi API 自动整理，并生成适合 A4 打印复习的 Word 单词本。

## 功能

- 支持 `.docx` Word 文档和 `.txt` 纯文本文件
- 调用 Kimi JSON Mode 输出结构化数据
- 自动生成三列表格 Word 单词本
- 保留 Kimi 原始 JSON 和结构化 JSON，便于核查
- 支持测试 API、刷新模型列表
- 支持将 API Key 保存到 Windows 凭据管理器

## 运行源码版

双击：

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
dist\Kimi雅思单词本整理器_太阳版.exe
```

## 注意

- 调用 Kimi API 会从你的开放平台余额中按量扣费。
- 第一次建议先用较短的 `.txt` 或 `.docx` 测试。
- `.venv`、`build`、`dist`、EXE 和 zip 文件不应提交到 GitHub。
