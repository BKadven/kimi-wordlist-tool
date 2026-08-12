from __future__ import annotations

import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import keyring
from keyring.errors import KeyringError, PasswordDeleteError
from openai import AuthenticationError, BadRequestError, OpenAIError, RateLimitError

from wordlist_engine import (
    ProcessingResult,
    SUPPORTED_INPUT_SUFFIXES,
    list_available_models,
    process_wordlist,
    test_api,
)

APP_NAME = "Kimi 雅思单词本整理器"
APP_FOLDER = "KimiWordlistTool"
DEFAULT_MODELS = ["kimi-k2.6", "kimi-k2.5"]
KEYRING_SERVICE = "KimiWordlistTool"
KEYRING_USERNAME = "moonshot_api_key"


def bundled_resource_path(*parts: str) -> Path:
    return Path(__file__).resolve().parent.joinpath(*parts)


def user_data_dir() -> Path:
    if sys.platform == "win32":
        root = Path(os.getenv("APPDATA", Path.home()))
    elif sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support"
    else:
        root = Path(os.getenv("XDG_CONFIG_HOME", Path.home() / ".config"))
    path = root / APP_FOLDER
    path.mkdir(parents=True, exist_ok=True)
    return path


def open_path(path: Path) -> None:
    path = Path(path)
    if sys.platform == "win32":
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def friendly_error(exc: Exception) -> str:
    if isinstance(exc, AuthenticationError):
        return "API Key 无效或已经失效。请重新复制 Kimi 开放平台中的 API Key。"
    if isinstance(exc, RateLimitError):
        return "请求受到限流，或账户余额不足。请稍后重试并检查 Kimi API 余额。"
    if isinstance(exc, BadRequestError):
        return f"Kimi 拒绝了本次请求：{exc}"
    if isinstance(exc, OpenAIError):
        return f"Kimi API 调用失败：{exc}"
    return str(exc) or exc.__class__.__name__


class WordlistApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("860x700")
        self.minsize(760, 620)

        # 设置窗口和任务栏图标。失败时不影响程序启动。
        try:
            self.iconbitmap(
                str(bundled_resource_path("resources", "sun_app.ico"))
            )
        except (tk.TclError, OSError):
            pass

        self.app_dir = user_data_dir()
        self.rules_path = self.app_dir / "rules.txt"
        self.settings_path = self.app_dir / "settings.json"
        self.default_rules_path = bundled_resource_path("resources", "default_rules.txt")
        self._ensure_rules_file()

        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_running = False
        self.last_result: ProcessingResult | None = None

        settings = self._load_settings()
        saved_api_key = self._load_saved_api_key()
        environment_api_key = os.getenv("MOONSHOT_API_KEY", "").strip()

        self.input_var = tk.StringVar(value=settings.get("last_input", ""))
        self.output_var = tk.StringVar(value=settings.get("last_output", str(Path.home() / "Documents")))
        self.api_key_var = tk.StringVar(value=environment_api_key or saved_api_key)
        self.model_var = tk.StringVar(value=settings.get("model", DEFAULT_MODELS[0]))
        self.show_key_var = tk.BooleanVar(value=False)
        self.remember_key_var = tk.BooleanVar(value=bool(saved_api_key and not environment_api_key))
        self.thinking_var = tk.BooleanVar(value=False)
        self.progress_var = tk.IntVar(value=0)
        self.status_var = tk.StringVar(value="请选择原始单词 Word 文件。")

        self._configure_style()
        self._build_ui()
        self.after(120, self._poll_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Subtitle.TLabel", font=("Microsoft YaHei UI", 9))
        style.configure("Section.TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(14, 8))
        style.configure("TButton", font=("Microsoft YaHei UI", 9), padding=(8, 5))
        style.configure("TLabel", font=("Microsoft YaHei UI", 9))
        style.configure("TEntry", font=("Microsoft YaHei UI", 9))
        style.configure("TCheckbutton", font=("Microsoft YaHei UI", 9))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(3, weight=1)

        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(header, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="选择原始 Word 或 TXT，调用 Kimi 自动整理，并生成适合 A4 打印的雅思单词本。",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        files_frame = ttk.LabelFrame(outer, text="文件", style="Section.TLabelframe", padding=12)
        files_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        files_frame.columnconfigure(1, weight=1)

        ttk.Label(files_frame, text="原始文件：").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(files_frame, textvariable=self.input_var).grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Button(files_frame, text="选择文件…", command=self._choose_input).grid(row=0, column=2, padx=(8, 0), pady=5)

        ttk.Label(files_frame, text="输出目录：").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(files_frame, textvariable=self.output_var).grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Button(files_frame, text="选择目录…", command=self._choose_output).grid(row=1, column=2, padx=(8, 0), pady=5)

        api_frame = ttk.LabelFrame(outer, text="Kimi API", style="Section.TLabelframe", padding=12)
        api_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        api_frame.columnconfigure(1, weight=1)

        ttk.Label(api_frame, text="API Key：").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="●")
        self.key_entry.grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Checkbutton(api_frame, text="显示", variable=self.show_key_var, command=self._toggle_key).grid(row=0, column=2, padx=(8, 0), pady=5)

        credential_row = ttk.Frame(api_frame)
        credential_row.grid(row=1, column=1, columnspan=2, sticky="w", pady=(1, 5))
        ttk.Checkbutton(
            credential_row,
            text="在这台电脑上记住 API Key",
            variable=self.remember_key_var,
        ).pack(side="left")
        ttk.Button(
            credential_row,
            text="清除已保存 Key",
            command=self._clear_saved_api_key,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(
            credential_row,
            text="由 Windows 凭据管理器保存，不写入软件目录。",
            foreground="#666666",
        ).pack(side="left", padx=(10, 0))

        ttk.Label(api_frame, text="模型：").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        self.model_combo = ttk.Combobox(api_frame, textvariable=self.model_var, values=DEFAULT_MODELS, state="normal")
        self.model_combo.grid(row=2, column=1, sticky="ew", pady=5)
        model_buttons = ttk.Frame(api_frame)
        model_buttons.grid(row=2, column=2, padx=(8, 0), pady=5)
        ttk.Button(model_buttons, text="刷新列表", command=self._refresh_models).pack(side="left")
        ttk.Button(model_buttons, text="测试 API", command=self._test_api).pack(side="left", padx=(6, 0))

        ttk.Checkbutton(
            api_frame,
            text="启用深度思考（通常更慢、消耗更多；整理单词本一般不需要）",
            variable=self.thinking_var,
        ).grid(row=3, column=1, columnspan=2, sticky="w", pady=(4, 2))

        rules_row = ttk.Frame(api_frame)
        rules_row.grid(row=4, column=1, columnspan=2, sticky="w", pady=(5, 0))
        ttk.Button(rules_row, text="打开整理规则", command=self._open_rules).pack(side="left")
        ttk.Button(rules_row, text="恢复默认规则", command=self._restore_rules).pack(side="left", padx=(6, 0))
        ttk.Label(rules_row, text="规则保存在你的 AppData 中，可自行修改。", foreground="#666666").pack(side="left", padx=(10, 0))

        run_frame = ttk.LabelFrame(outer, text="运行状态", style="Section.TLabelframe", padding=12)
        run_frame.grid(row=3, column=0, sticky="nsew")
        run_frame.columnconfigure(0, weight=1)
        run_frame.rowconfigure(3, weight=1)

        controls = ttk.Frame(run_frame)
        controls.grid(row=0, column=0, sticky="ew")
        self.start_button = ttk.Button(controls, text="开始整理", style="Primary.TButton", command=self._start_processing)
        self.start_button.pack(side="left")
        self.open_result_button = ttk.Button(controls, text="打开结果", command=self._open_result, state="disabled")
        self.open_result_button.pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="打开输出目录", command=self._open_output_dir).pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(run_frame, variable=self.progress_var, maximum=100, mode="determinate")
        self.progress.grid(row=1, column=0, sticky="ew", pady=(12, 4))
        ttk.Label(run_frame, textvariable=self.status_var).grid(row=2, column=0, sticky="w", pady=(0, 7))

        self.log = ScrolledText(
            run_frame,
            height=13,
            wrap="word",
            font=("Microsoft YaHei UI", 9),
            state="disabled",
        )
        self.log.grid(row=3, column=0, sticky="nsew")
        if self.api_key_var.get().strip():
            self._append_log("准备就绪。已从 Windows 凭据管理器或环境变量读取 API Key。")
        else:
            self._append_log("准备就绪。可勾选“在这台电脑上记住 API Key”。")

        footer = ttk.Label(
            outer,
            text="提示：调用 Kimi API 会从你的开放平台余额中按量扣费。首次建议用较短文档测试。",
            foreground="#666666",
        )
        footer.grid(row=4, column=0, sticky="w", pady=(9, 0))

    def _load_saved_api_key(self) -> str:
        try:
            return keyring.get_password(KEYRING_SERVICE, KEYRING_USERNAME) or ""
        except KeyringError:
            return ""

    def _save_api_key_if_requested(self) -> None:
        if not self.remember_key_var.get():
            return

        api_key = self.api_key_var.get().strip()
        if not api_key:
            return

        try:
            keyring.set_password(KEYRING_SERVICE, KEYRING_USERNAME, api_key)
            self._append_log("API Key 已安全保存到 Windows 凭据管理器。")
        except KeyringError as exc:
            self._append_log(f"无法保存 API Key：{exc}")
            messagebox.showwarning(
                "无法记住 API Key",
                "API 调用本身不受影响，但 Windows 凭据管理器未能保存 Key。\n\n"
                f"详细信息：{exc}",
                parent=self,
            )

    def _clear_saved_api_key(self) -> None:
        confirmed = messagebox.askyesno(
            "清除已保存 Key",
            "确定从这台电脑的 Windows 凭据管理器中删除已保存的 Kimi API Key 吗？",
            parent=self,
        )
        if not confirmed:
            return

        try:
            keyring.delete_password(KEYRING_SERVICE, KEYRING_USERNAME)
        except PasswordDeleteError:
            pass
        except KeyringError as exc:
            messagebox.showerror(
                "清除失败",
                f"无法从 Windows 凭据管理器删除 Key：{exc}",
                parent=self,
            )
            return

        self.api_key_var.set("")
        self.remember_key_var.set(False)
        self._append_log("已从 Windows 凭据管理器清除 API Key。")
        messagebox.showinfo("已清除", "已保存的 API Key 已删除。", parent=self)

    def _ensure_rules_file(self) -> None:
        if self.rules_path.exists():
            return
        if not self.default_rules_path.exists():
            raise FileNotFoundError(f"程序内缺少默认规则文件：{self.default_rules_path}")
        shutil.copy2(self.default_rules_path, self.rules_path)

    def _load_settings(self) -> dict[str, str]:
        try:
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return {}

    def _save_settings(self) -> None:
        data = {
            "last_input": self.input_var.get().strip(),
            "last_output": self.output_var.get().strip(),
            "model": self.model_var.get().strip(),
        }
        try:
            self.settings_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _choose_input(self) -> None:
        initial = Path(self.input_var.get()).parent if self.input_var.get() else Path.home()
        path = filedialog.askopenfilename(
            title="选择原始单词文件",
            initialdir=str(initial),
            filetypes=[
                ("支持的文档", "*.docx *.txt"),
                ("Word 文档", "*.docx"),
                ("TXT 文本文档", "*.txt"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.input_var.set(path)
            if not self.output_var.get().strip():
                self.output_var.set(str(Path(path).parent))

    def _choose_output(self) -> None:
        initial = self.output_var.get().strip() or str(Path.home() / "Documents")
        path = filedialog.askdirectory(title="选择输出目录", initialdir=initial)
        if path:
            self.output_var.set(path)

    def _toggle_key(self) -> None:
        self.key_entry.configure(show="" if self.show_key_var.get() else "●")

    def _open_rules(self) -> None:
        try:
            open_path(self.rules_path)
        except Exception as exc:
            messagebox.showerror("无法打开规则", friendly_error(exc), parent=self)

    def _restore_rules(self) -> None:
        confirmed = messagebox.askyesno(
            "恢复默认规则",
            "这会覆盖你对规则文件所做的修改。确定继续吗？",
            parent=self,
        )
        if not confirmed:
            return
        try:
            shutil.copy2(self.default_rules_path, self.rules_path)
            self._append_log("已恢复默认整理规则。")
            messagebox.showinfo("完成", "默认规则已恢复。", parent=self)
        except Exception as exc:
            messagebox.showerror("恢复失败", friendly_error(exc), parent=self)

    def _validate_api_fields(self) -> tuple[str, str]:
        key = self.api_key_var.get().strip()
        model = self.model_var.get().strip()
        if not key:
            raise ValueError("请先填写 Kimi API Key。")
        if not model:
            raise ValueError("请填写模型名称。")
        return key, model

    def _set_busy(self, busy: bool) -> None:
        self.worker_running = busy
        self.start_button.configure(state="disabled" if busy else "normal")
        if busy:
            self.open_result_button.configure(state="disabled")

    def _start_processing(self) -> None:
        if self.worker_running:
            return
        try:
            input_path = Path(self.input_var.get().strip())
            output_dir = Path(self.output_var.get().strip())
            api_key, model = self._validate_api_fields()
            if input_path.suffix.lower() not in SUPPORTED_INPUT_SUFFIXES or not input_path.is_file():
                raise ValueError("请选择一个有效的 .docx 或 .txt 原始单词记录文件。")
            if not self.output_var.get().strip():
                raise ValueError("请选择输出目录。")
        except Exception as exc:
            messagebox.showwarning("信息不完整", friendly_error(exc), parent=self)
            return

        confirmed = messagebox.askyesno(
            "确认调用 API",
            "整理过程会调用 Kimi API，并从开放平台余额中按量扣费。\n\n确定开始吗？",
            parent=self,
        )
        if not confirmed:
            return

        self._save_settings()
        self.last_result = None
        self.progress_var.set(0)
        self.status_var.set("正在启动……")
        self._append_log("\n—— 开始新的整理任务 ——")
        self._append_log(f"输入：{input_path}")
        self._append_log(f"输出：{output_dir}")
        self._append_log(f"模型：{model}")
        self._set_busy(True)

        threading.Thread(
            target=self._processing_worker,
            kwargs={
                "input_path": input_path,
                "output_dir": output_dir,
                "api_key": api_key,
                "model": model,
                "thinking_enabled": self.thinking_var.get(),
            },
            daemon=True,
        ).start()

    def _processing_worker(
        self,
        *,
        input_path: Path,
        output_dir: Path,
        api_key: str,
        model: str,
        thinking_enabled: bool,
    ) -> None:
        try:
            result = process_wordlist(
                input_docx=input_path,
                output_dir=output_dir,
                rules_path=self.rules_path,
                api_key=api_key,
                model=model,
                thinking_enabled=thinking_enabled,
                progress_callback=lambda percent, message: self.events.put(("progress", (percent, message))),
            )
            self.events.put(("success", result))
        except Exception as exc:
            self.events.put(("error", (exc, traceback.format_exc())))

    def _test_api(self) -> None:
        if self.worker_running:
            return
        try:
            api_key, model = self._validate_api_fields()
        except Exception as exc:
            messagebox.showwarning("信息不完整", friendly_error(exc), parent=self)
            return

        self._append_log(f"正在测试模型 {model}……")
        self.status_var.set("正在测试 API……")
        self._set_busy(True)

        def worker() -> None:
            try:
                reply = test_api(api_key, model)
                self.events.put(("api_test_success", reply))
            except Exception as exc:
                self.events.put(("error", (exc, traceback.format_exc())))

        threading.Thread(target=worker, daemon=True).start()

    def _refresh_models(self) -> None:
        if self.worker_running:
            return
        try:
            api_key, _ = self._validate_api_fields()
        except Exception as exc:
            messagebox.showwarning("需要 API Key", friendly_error(exc), parent=self)
            return

        self._append_log("正在从 Kimi 开放平台获取可用模型列表……")
        self.status_var.set("正在刷新模型列表……")
        self._set_busy(True)

        def worker() -> None:
            try:
                models = list_available_models(api_key)
                self.events.put(("models", models))
            except Exception as exc:
                self.events.put(("error", (exc, traceback.format_exc())))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "progress":
                    percent, message = payload  # type: ignore[misc]
                    self.progress_var.set(int(percent))
                    self.status_var.set(str(message))
                    self._append_log(str(message))
                elif event == "success":
                    self._handle_success(payload)  # type: ignore[arg-type]
                elif event == "error":
                    exc, trace = payload  # type: ignore[misc]
                    self._handle_error(exc, trace)
                elif event == "models":
                    self._handle_models(payload)  # type: ignore[arg-type]
                elif event == "api_test_success":
                    self._set_busy(False)
                    self._save_api_key_if_requested()
                    self.status_var.set("API 连接正常。")
                    self._append_log(f"API 测试结果：{payload}")
                    messagebox.showinfo("测试成功", f"Kimi API 连接正常。\n\n返回：{payload}", parent=self)
        except queue.Empty:
            pass
        finally:
            self.after(120, self._poll_events)

    def _handle_success(self, result: ProcessingResult) -> None:
        self._set_busy(False)
        self._save_api_key_if_requested()
        self.last_result = result
        self.progress_var.set(100)
        self.status_var.set("整理完成。")
        self.open_result_button.configure(state="normal")
        self._append_log(f"Word 已生成：{result.output_docx}")
        self._append_log(f"结构化 JSON：{result.parsed_json}")
        if result.total_tokens is not None:
            self._append_log(
                f"Token 用量：输入 {result.prompt_tokens or 0:,}，"
                f"输出 {result.completion_tokens or 0:,}，总计 {result.total_tokens:,}。"
            )

        open_now = messagebox.askyesno(
            "整理完成",
            f"文件已生成：\n{result.output_docx}\n\n现在打开 Word 文件吗？",
            parent=self,
        )
        if open_now:
            self._open_result()

    def _handle_error(self, exc: Exception, trace: str) -> None:
        self._set_busy(False)
        self.status_var.set("处理失败。")
        self.progress_var.set(0)
        message = friendly_error(exc)
        self._append_log(f"错误：{message}")
        debug_path = self.app_dir / "last_error.log"
        try:
            debug_path.write_text(trace, encoding="utf-8")
            self._append_log(f"详细错误日志：{debug_path}")
        except OSError:
            pass
        messagebox.showerror("处理失败", message, parent=self)

    def _handle_models(self, models: list[str]) -> None:
        self._set_busy(False)
        if not models:
            self.status_var.set("没有获取到模型列表。")
            self._append_log("未获取到可用模型，保留默认列表。")
            return
        self.model_combo.configure(values=models)
        if self.model_var.get().strip() not in models:
            preferred = next((name for name in models if name == "kimi-k2.6"), models[0])
            self.model_var.set(preferred)
        self.status_var.set(f"已获取 {len(models)} 个可用模型。")
        self._append_log(f"已刷新模型列表，共 {len(models)} 个。")

    def _open_result(self) -> None:
        if not self.last_result:
            return
        try:
            open_path(self.last_result.output_docx)
        except Exception as exc:
            messagebox.showerror("无法打开结果", friendly_error(exc), parent=self)

    def _open_output_dir(self) -> None:
        path_text = self.output_var.get().strip()
        if not path_text:
            messagebox.showwarning("没有输出目录", "请先选择输出目录。", parent=self)
            return
        path = Path(path_text)
        try:
            path.mkdir(parents=True, exist_ok=True)
            open_path(path)
        except Exception as exc:
            messagebox.showerror("无法打开目录", friendly_error(exc), parent=self)

    def _on_close(self) -> None:
        if self.worker_running:
            confirmed = messagebox.askyesno(
                "任务仍在运行",
                "整理任务尚未结束。现在关闭会中断本次任务，确定退出吗？",
                parent=self,
            )
            if not confirmed:
                return
        self._save_settings()
        self.destroy()


if __name__ == "__main__":
    WordlistApp().mainloop()
