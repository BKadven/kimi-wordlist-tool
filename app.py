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
    API_PROVIDERS,
    DEFAULT_PROVIDER_ID,
    ProcessingResult,
    SUPPORTED_INPUT_SUFFIXES,
    get_api_provider,
    list_available_models,
    process_wordlist,
    test_api,
)

APP_NAME = "Wgen"
APP_FOLDER = "Wgen"
LEGACY_APP_FOLDER = "KimiWordlistTool"
DEFAULT_HOMEPAGE_REPO = Path.home() / "Documents" / "All_Program" / "personal-homepage"
KEYRING_SERVICE = "Wgen"
LEGACY_KEYRING_SERVICE = "KimiWordlistTool"


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


def legacy_user_data_dir() -> Path:
    if sys.platform == "win32":
        root = Path(os.getenv("APPDATA", Path.home()))
    elif sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support"
    else:
        root = Path(os.getenv("XDG_CONFIG_HOME", Path.home() / ".config"))
    return root / LEGACY_APP_FOLDER


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
        return "API Key 无效或已经失效。请重新复制所选平台中的 API Key。"
    if isinstance(exc, RateLimitError):
        return "请求受到限流，或账户余额不足。请稍后重试并检查所选平台的 API 余额。"
    if isinstance(exc, BadRequestError):
        return f"模型服务拒绝了本次请求：{exc}"
    if isinstance(exc, OpenAIError):
        return f"模型 API 调用失败：{exc}"
    return str(exc) or exc.__class__.__name__


class WordlistApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("860x760")
        self.minsize(780, 680)

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
        self._migrate_legacy_files()
        self._ensure_rules_file()

        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_running = False
        self.last_result: ProcessingResult | None = None

        settings = self._load_settings()
        provider_id = self._resolve_initial_provider(settings)
        provider = get_api_provider(provider_id)
        saved_api_key = self._load_saved_api_key(provider.id)
        environment_api_key = os.getenv(provider.env_var, "").strip()

        self.input_var = tk.StringVar(value=settings.get("last_input", ""))
        self.output_var = tk.StringVar(value=settings.get("last_output", str(Path.home() / "Documents")))
        self.provider_display_to_id = {provider.display_name: provider.id for provider in API_PROVIDERS.values()}
        self.provider_var = tk.StringVar(value=provider.id)
        self.provider_display_var = tk.StringVar(value=provider.display_name)
        self.api_key_var = tk.StringVar(value=environment_api_key or saved_api_key)
        model_setting = settings.get("model", "").strip()
        self.model_var = tk.StringVar(value=model_setting or provider.default_model)
        self.show_key_var = tk.BooleanVar(value=False)
        self.remember_key_var = tk.BooleanVar(value=bool(saved_api_key and not environment_api_key))
        self.thinking_var = tk.BooleanVar(value=False)
        homepage_repo_setting = settings.get("homepage_repo", str(DEFAULT_HOMEPAGE_REPO))
        homepage_repo_exists = Path(homepage_repo_setting).is_dir()
        self.publish_homepage_var = tk.BooleanVar(
            value=settings.get("publish_homepage", "1") != "0" and homepage_repo_exists
        )
        self.homepage_repo_var = tk.StringVar(value=homepage_repo_setting)
        self.push_homepage_var = tk.BooleanVar(value=settings.get("push_homepage", "1") != "0")
        self.progress_var = tk.IntVar(value=0)
        self.status_var = tk.StringVar(value="请选择原始单词 Word 或 TXT 文件。")

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
            text="选择原始 Word 或 TXT，调用所选模型自动整理，并生成打印版 Word 和手机阅读版 HTML。",
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

        ttk.Checkbutton(
            files_frame,
            text="整理完成后归档到个人主页",
            variable=self.publish_homepage_var,
        ).grid(row=2, column=1, columnspan=2, sticky="w", pady=(7, 2))

        ttk.Label(files_frame, text="主页仓库：").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(files_frame, textvariable=self.homepage_repo_var).grid(row=3, column=1, sticky="ew", pady=5)
        ttk.Button(files_frame, text="选择仓库…", command=self._choose_homepage_repo).grid(row=3, column=2, padx=(8, 0), pady=5)

        ttk.Checkbutton(
            files_frame,
            text="自动提交并推送到 GitHub",
            variable=self.push_homepage_var,
        ).grid(row=4, column=1, columnspan=2, sticky="w", pady=(1, 3))

        api_frame = ttk.LabelFrame(outer, text="模型 API", style="Section.TLabelframe", padding=12)
        api_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        api_frame.columnconfigure(1, weight=1)

        ttk.Label(api_frame, text="模型服务：").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.provider_combo = ttk.Combobox(
            api_frame,
            textvariable=self.provider_display_var,
            values=list(self.provider_display_to_id.keys()),
            state="readonly",
        )
        self.provider_combo.grid(row=0, column=1, sticky="ew", pady=5)
        self.provider_combo.bind("<<ComboboxSelected>>", self._on_provider_changed)

        ttk.Label(api_frame, text="API Key：").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="●")
        self.key_entry.grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Checkbutton(api_frame, text="显示", variable=self.show_key_var, command=self._toggle_key).grid(row=1, column=2, padx=(8, 0), pady=5)

        credential_row = ttk.Frame(api_frame)
        credential_row.grid(row=2, column=1, columnspan=2, sticky="w", pady=(1, 5))
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

        ttk.Label(api_frame, text="模型：").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=5)
        self.model_combo = ttk.Combobox(
            api_frame,
            textvariable=self.model_var,
            values=get_api_provider(self.provider_var.get()).default_models,
            state="normal",
        )
        self.model_combo.grid(row=3, column=1, sticky="ew", pady=5)
        model_buttons = ttk.Frame(api_frame)
        model_buttons.grid(row=3, column=2, padx=(8, 0), pady=5)
        ttk.Button(model_buttons, text="刷新列表", command=self._refresh_models).pack(side="left")
        ttk.Button(model_buttons, text="测试 API", command=self._test_api).pack(side="left", padx=(6, 0))

        ttk.Checkbutton(
            api_frame,
            text="启用高强度思考（通常更慢、消耗更多；整理单词本一般不需要）",
            variable=self.thinking_var,
        ).grid(row=4, column=1, columnspan=2, sticky="w", pady=(4, 2))

        rules_row = ttk.Frame(api_frame)
        rules_row.grid(row=5, column=1, columnspan=2, sticky="w", pady=(5, 0))
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
        self.open_result_button = ttk.Button(controls, text="打开 Word", command=self._open_result, state="disabled")
        self.open_result_button.pack(side="left", padx=(8, 0))
        self.open_html_button = ttk.Button(controls, text="打开手机 HTML", command=self._open_mobile_html, state="disabled")
        self.open_html_button.pack(side="left", padx=(8, 0))
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
            text="提示：调用模型 API 会从所选平台余额中按量扣费。首次建议用较短文档测试。",
            foreground="#666666",
        )
        footer.grid(row=4, column=0, sticky="w", pady=(9, 0))

    def _load_saved_api_key(self, provider_id: str) -> str:
        provider = get_api_provider(provider_id)
        try:
            saved = keyring.get_password(KEYRING_SERVICE, provider.keyring_username) or ""
            if saved:
                return saved
            if provider.id == "moonshot":
                return keyring.get_password(LEGACY_KEYRING_SERVICE, provider.keyring_username) or ""
            return ""
        except KeyringError:
            return ""

    def _save_api_key_if_requested(self) -> None:
        if not self.remember_key_var.get():
            return

        provider = get_api_provider(self.provider_var.get())
        api_key = self.api_key_var.get().strip()
        if not api_key:
            return

        try:
            keyring.set_password(KEYRING_SERVICE, provider.keyring_username, api_key)
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
        provider = get_api_provider(self.provider_var.get())
        confirmed = messagebox.askyesno(
            "清除已保存 Key",
            f"确定从这台电脑的 Windows 凭据管理器中删除已保存的 {provider.display_name} API Key 吗？",
            parent=self,
        )
        if not confirmed:
            return

        try:
            try:
                keyring.delete_password(KEYRING_SERVICE, provider.keyring_username)
            except PasswordDeleteError:
                pass
            if provider.id == "moonshot":
                try:
                    keyring.delete_password(LEGACY_KEYRING_SERVICE, provider.keyring_username)
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

    def _migrate_legacy_files(self) -> None:
        legacy_dir = legacy_user_data_dir()
        if not legacy_dir.is_dir():
            return
        for filename in ("rules.txt", "settings.json"):
            legacy_path = legacy_dir / filename
            target_path = self.app_dir / filename
            if legacy_path.is_file() and not target_path.exists():
                try:
                    shutil.copy2(legacy_path, target_path)
                except OSError:
                    pass

    def _resolve_initial_provider(self, settings: dict[str, str]) -> str:
        provider_id = settings.get("provider", "").strip()
        if provider_id in API_PROVIDERS:
            return provider_id
        model = settings.get("model", "").strip().lower()
        if model.startswith("kimi-"):
            return "moonshot"
        if model.startswith("glm-"):
            return "bigmodel"
        return DEFAULT_PROVIDER_ID

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
            "provider": self.provider_var.get().strip(),
            "model": self.model_var.get().strip(),
            "publish_homepage": "1" if self.publish_homepage_var.get() else "0",
            "homepage_repo": self.homepage_repo_var.get().strip(),
            "push_homepage": "1" if self.push_homepage_var.get() else "0",
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

    def _choose_homepage_repo(self) -> None:
        initial = self.homepage_repo_var.get().strip() or str(DEFAULT_HOMEPAGE_REPO.parent)
        path = filedialog.askdirectory(title="选择 personal-homepage 仓库目录", initialdir=initial)
        if path:
            self.homepage_repo_var.set(path)

    def _on_provider_changed(self, _event: object | None = None) -> None:
        provider_id = self.provider_display_to_id.get(self.provider_display_var.get(), DEFAULT_PROVIDER_ID)
        provider = get_api_provider(provider_id)
        self.provider_var.set(provider.id)
        self.model_combo.configure(values=provider.default_models)
        if self.model_var.get().strip() not in provider.default_models:
            self.model_var.set(provider.default_model)
        saved_api_key = self._load_saved_api_key(provider.id)
        environment_api_key = os.getenv(provider.env_var, "").strip()
        self.api_key_var.set(environment_api_key or saved_api_key)
        self.remember_key_var.set(bool(saved_api_key and not environment_api_key))
        self._append_log(f"已切换模型服务：{provider.display_name}。")

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

    def _validate_api_fields(self) -> tuple[str, str, str]:
        provider = get_api_provider(self.provider_var.get())
        key = self.api_key_var.get().strip()
        model = self.model_var.get().strip()
        if not key:
            raise ValueError(f"请先填写 {provider.display_name} API Key。")
        if not model:
            raise ValueError("请填写模型名称。")
        return provider.id, key, model

    def _set_busy(self, busy: bool) -> None:
        self.worker_running = busy
        self.start_button.configure(state="disabled" if busy else "normal")
        if busy:
            self.open_result_button.configure(state="disabled")
            self.open_html_button.configure(state="disabled")

    def _start_processing(self) -> None:
        if self.worker_running:
            return
        try:
            input_path = Path(self.input_var.get().strip())
            output_dir = Path(self.output_var.get().strip())
            provider_id, api_key, model = self._validate_api_fields()
            if input_path.suffix.lower() not in SUPPORTED_INPUT_SUFFIXES or not input_path.is_file():
                raise ValueError("请选择一个有效的 .docx 或 .txt 原始单词记录文件。")
            if not self.output_var.get().strip():
                raise ValueError("请选择输出目录。")
            if self.publish_homepage_var.get():
                homepage_repo = Path(self.homepage_repo_var.get().strip())
                if not homepage_repo.is_dir():
                    raise ValueError("已启用个人主页归档，请选择有效的 personal-homepage 仓库目录。")
        except Exception as exc:
            messagebox.showwarning("信息不完整", friendly_error(exc), parent=self)
            return

        confirmed = messagebox.askyesno(
            "确认调用 API",
            "整理过程会调用所选模型 API，并从对应平台余额中按量扣费。\n\n确定开始吗？",
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
        self._append_log(f"模型服务：{get_api_provider(provider_id).display_name}")
        self._append_log(f"模型：{model}")
        if self.publish_homepage_var.get():
            self._append_log(f"个人主页归档：{self.homepage_repo_var.get().strip()}")
            self._append_log("发布方式：自动提交并推送到 GitHub" if self.push_homepage_var.get() else "发布方式：只写入本地个人主页仓库")
        self._set_busy(True)

        threading.Thread(
            target=self._processing_worker,
            kwargs={
                "input_path": input_path,
                "output_dir": output_dir,
                "provider_id": provider_id,
                "api_key": api_key,
                "model": model,
                "thinking_enabled": self.thinking_var.get(),
                "publish_to_homepage": self.publish_homepage_var.get(),
                "homepage_repo": Path(self.homepage_repo_var.get().strip()) if self.publish_homepage_var.get() else None,
                "push_homepage": self.push_homepage_var.get(),
            },
            daemon=True,
        ).start()

    def _processing_worker(
        self,
        *,
        input_path: Path,
        output_dir: Path,
        provider_id: str,
        api_key: str,
        model: str,
        thinking_enabled: bool,
        publish_to_homepage: bool,
        homepage_repo: Path | None,
        push_homepage: bool,
    ) -> None:
        try:
            result = process_wordlist(
                input_docx=input_path,
                output_dir=output_dir,
                rules_path=self.rules_path,
                provider_id=provider_id,
                api_key=api_key,
                model=model,
                thinking_enabled=thinking_enabled,
                publish_to_homepage=publish_to_homepage,
                homepage_repo=homepage_repo,
                push_homepage=push_homepage,
                progress_callback=lambda percent, message: self.events.put(("progress", (percent, message))),
            )
            self.events.put(("success", result))
        except Exception as exc:
            self.events.put(("error", (exc, traceback.format_exc())))

    def _test_api(self) -> None:
        if self.worker_running:
            return
        try:
            provider_id, api_key, model = self._validate_api_fields()
        except Exception as exc:
            messagebox.showwarning("信息不完整", friendly_error(exc), parent=self)
            return

        self._append_log(f"正在测试 {get_api_provider(provider_id).display_name} 模型 {model}……")
        self.status_var.set("正在测试 API……")
        self._set_busy(True)

        def worker() -> None:
            try:
                reply = test_api(api_key, model, provider_id)
                self.events.put(("api_test_success", reply))
            except Exception as exc:
                self.events.put(("error", (exc, traceback.format_exc())))

        threading.Thread(target=worker, daemon=True).start()

    def _refresh_models(self) -> None:
        if self.worker_running:
            return
        try:
            provider_id, api_key, _ = self._validate_api_fields()
        except Exception as exc:
            messagebox.showwarning("需要 API Key", friendly_error(exc), parent=self)
            return

        self._append_log(f"正在从 {get_api_provider(provider_id).display_name} 获取可用模型列表……")
        self.status_var.set("正在刷新模型列表……")
        self._set_busy(True)

        def worker() -> None:
            try:
                models = list_available_models(api_key, provider_id)
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
                    provider = get_api_provider(self.provider_var.get())
                    messagebox.showinfo("测试成功", f"{provider.display_name} API 连接正常。\n\n返回：{payload}", parent=self)
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
        self.open_html_button.configure(state="normal")
        self._append_log(f"Word 已生成：{result.output_docx}")
        self._append_log(f"手机 HTML 已生成：{result.output_html}")
        if result.homepage_publish:
            self._append_log(f"个人主页归档：{result.homepage_publish.archive_dir}")
            if result.homepage_publish.commit_hash:
                self._append_log(f"Git 提交：{result.homepage_publish.commit_hash}")
            if result.homepage_publish.url:
                self._append_log(f"网站地址：{result.homepage_publish.url}")
        self._append_log(f"结构化 JSON：{result.parsed_json}")
        if result.total_tokens is not None:
            self._append_log(
                f"Token 用量：输入 {result.prompt_tokens or 0:,}，"
                f"输出 {result.completion_tokens or 0:,}，总计 {result.total_tokens:,}。"
            )

        open_now = messagebox.askyesno(
            "整理完成",
            f"Word 已生成：\n{result.output_docx}\n\n手机 HTML 已生成：\n{result.output_html}"
            + (
                f"\n\n已归档到个人主页：\n{result.homepage_publish.url or result.homepage_publish.archive_dir}"
                if result.homepage_publish
                else ""
            )
            + "\n\n现在打开手机版 HTML 吗？",
            parent=self,
        )
        if open_now:
            self._open_mobile_html()

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
            provider = get_api_provider(self.provider_var.get())
            preferred = next((name for name in models if name == provider.default_model), models[0])
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

    def _open_mobile_html(self) -> None:
        if not self.last_result:
            return
        try:
            open_path(self.last_result.output_html)
        except Exception as exc:
            messagebox.showerror("无法打开手机 HTML", friendly_error(exc), parent=self)

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
