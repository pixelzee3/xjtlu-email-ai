"""图形化启动助手：环境检查、半自动安装、报错粘贴诊断。"""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

from startup_helper_core import (
    diagnose_error_text,
    iter_subprocess_lines,
    repo_root,
    run_all_checks,
    suggested_pip_install_command,
    suggested_playwright_install_command,
    suggested_venv_create_command,
)

ROOT = repo_root()


class HelperApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("邮件助手 · 安装与环境诊断")
        self.geometry("900x640")
        self.minsize(760, 520)

        self._log_q: queue.Queue[str] = queue.Queue()

        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        tab_env = ttk.Frame(nb)
        tab_err = ttk.Frame(nb)
        nb.add(tab_env, text="环境检查")
        nb.add(tab_err, text="粘贴报错分析")

        self._build_env_tab(tab_env)
        self._build_err_tab(tab_err)

        self.after(120, self._poll_log)
        self._run_checks_async()

    def _build_env_tab(self, parent: ttk.Frame) -> None:
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, pady=(0, 6))

        ttk.Button(bar, text="重新检查", command=self._run_checks_async).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="创建 .venv", command=self._on_create_venv).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="安装 pip 依赖", command=self._on_pip_install).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="安装 Playwright 浏览器", command=self._on_playwright_install).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(bar, text="启动主程序", command=self._on_launch_main).pack(side=tk.LEFT, padx=12)
        ttk.Button(bar, text="打开说明", command=self._open_readme).pack(side=tk.RIGHT, padx=2)

        tree_fr = ttk.Frame(parent)
        tree_fr.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(
            tree_fr,
            columns=("status", "detail"),
            show="tree headings",
            height=9,
        )
        self.tree.heading("#0", text="检查项")
        self.tree.heading("status", text="状态")
        self.tree.heading("detail", text="说明")
        self.tree.column("#0", width=200)
        self.tree.column("status", width=72)
        self.tree.column("detail", width=580)
        sb = ttk.Scrollbar(tree_fr, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("ok", foreground="#0a6b0a")
        self.tree.tag_configure("warn", foreground="#8a6d1a")
        self.tree.tag_configure("err", foreground="#a01010")

        ttk.Label(parent, text="安装/命令输出（滚动查看）").pack(anchor=tk.W, pady=(8, 0))
        self.log = scrolledtext.ScrolledText(parent, height=11, state=tk.DISABLED, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True, pady=4)

    def _build_err_tab(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="将黑窗口或网页里的报错全文粘贴到下方，点击「分析」。").pack(anchor=tk.W)
        self.err_in = scrolledtext.ScrolledText(parent, height=14, font=("Consolas", 9))
        self.err_in.pack(fill=tk.BOTH, expand=True, pady=6)
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="分析", command=self._on_analyze).pack(side=tk.LEFT)
        ttk.Button(bar, text="清空", command=lambda: self.err_in.delete("1.0", tk.END)).pack(side=tk.LEFT, padx=6)

        ttk.Label(parent, text="诊断与建议").pack(anchor=tk.W, pady=(10, 0))
        self.err_out = scrolledtext.ScrolledText(
            parent, height=12, state=tk.DISABLED, font=("Microsoft YaHei UI", 10)
        )
        self.err_out.pack(fill=tk.BOTH, expand=True, pady=4)

    def _append_log(self, line: str) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, line + "\n")
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def _poll_log(self) -> None:
        try:
            while True:
                line = self._log_q.get_nowait()
                self._append_log(line)
        except queue.Empty:
            pass
        self.after(120, self._poll_log)

    def _run_cmd_logged(self, argv: list[str], cwd: Path | None = None) -> None:
        def worker() -> None:
            try:
                for line in iter_subprocess_lines(argv, cwd=cwd):
                    self._log_q.put(line)
            except Exception as e:
                self._log_q.put(f"[异常] {e}")
            self._log_q.put("--- 完成，可点「重新检查」---")

        threading.Thread(target=worker, daemon=True).start()

    def _run_checks_async(self) -> None:
        def worker() -> None:
            try:
                items = run_all_checks()
                self.after(0, lambda: self._apply_checks(items))
            except Exception as e:
                err = str(e)
                self.after(0, lambda msg=err: messagebox.showerror("检查失败", msg))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_checks(self, items) -> None:
        for ch in self.tree.get_children():
            self.tree.delete(ch)
        for it in items:
            if not it.ok:
                st = "失败"
                tag = "err"
            elif it.severity == "warn":
                st = "警告"
                tag = "warn"
            else:
                st = "通过"
                tag = "ok"
            self.tree.insert("", tk.END, text=it.title, values=(st, it.detail), tags=(tag,))

    def _on_create_venv(self) -> None:
        cmd = suggested_venv_create_command()
        if not messagebox.askyesno(
            "确认",
            "将创建项目根目录下的 .venv：\n\n" + subprocess.list2cmdline(cmd) + "\n\n是否继续？",
        ):
            return
        self._append_log(subprocess.list2cmdline(cmd))
        self._run_cmd_logged(cmd, cwd=ROOT)

    def _on_pip_install(self) -> None:
        cmd = suggested_pip_install_command()
        if not messagebox.askyesno(
            "确认",
            "将执行 pip 安装 requirements.txt：\n\n" + subprocess.list2cmdline(cmd) + "\n\n是否继续？",
        ):
            return
        self._append_log(subprocess.list2cmdline(cmd))
        self._run_cmd_logged(cmd, cwd=ROOT)

    def _on_playwright_install(self) -> None:
        cmd = suggested_playwright_install_command()
        if not messagebox.askyesno(
            "确认",
            "将下载 Playwright 的 Edge 组件（可能较慢）：\n\n"
            + subprocess.list2cmdline(cmd)
            + "\n\n是否继续？",
        ):
            return
        self._append_log(subprocess.list2cmdline(cmd))
        self._run_cmd_logged(cmd, cwd=ROOT)

    def _on_launch_main(self) -> None:
        bat = ROOT / "run_app.bat"
        if not bat.is_file():
            messagebox.showerror("未找到", str(bat))
            return
        if not messagebox.askokcancel(
            "启动主程序",
            "将打开新的命令行窗口运行 run_app.bat（Web 服务）。\n本窗口可保留或关闭。",
        ):
            return
        os.startfile(str(bat))  # type: ignore[attr-defined]

    def _open_readme(self) -> None:
        p = ROOT / "docs" / "README.md"
        if p.is_file():
            os.startfile(str(p))  # type: ignore[attr-defined]
        else:
            messagebox.showinfo("说明", f"未找到 {p}")

    def _on_analyze(self) -> None:
        text = self.err_in.get("1.0", tk.END)
        diag = diagnose_error_text(text)
        self.err_out.configure(state=tk.NORMAL)
        self.err_out.delete("1.0", tk.END)
        if not diag:
            self.err_out.insert(tk.END, "请先粘贴报错内容。")
            self.err_out.configure(state=tk.DISABLED)
            return
        lines = [
            f"类型: {diag.title} ({diag.code})",
            "",
            diag.summary,
            "",
            "建议:",
        ]
        for h in diag.hints:
            lines.append(f"· {h}")
        if diag.commands:
            lines.append("")
            lines.append("可复制命令:")
            for c in diag.commands:
                lines.append(c)
        self.err_out.insert(tk.END, "\n".join(lines))
        self.err_out.configure(state=tk.DISABLED)


def _crash_log_path() -> str:
    try:
        return str(repo_root() / "helper_last_error.log")
    except Exception:
        return str(os.path.expanduser("~/practice2_helper_error.log"))


def main() -> None:
    try:
        app = HelperApp()
    except tk.TclError as e:
        print("无法启动图形界面（tkinter 不可用）:", e, file=sys.stderr)
        print("请安装带 Tcl/Tk 的官方 Python，或使用命令行安装依赖。", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        log = _crash_log_path()
        try:
            Path(log).parent.mkdir(parents=True, exist_ok=True)
            Path(log).write_text(traceback.format_exc(), encoding="utf-8")
        except OSError:
            log = "(无法写入日志文件)"
        try:
            messagebox.showerror(
                "启动助手失败",
                f"{e}\n\n详情已保存到:\n{log}",
            )
        except Exception:
            traceback.print_exc()
        sys.exit(1)
    try:
        app.mainloop()
    except Exception:
        try:
            Path(_crash_log_path()).write_text(traceback.format_exc(), encoding="utf-8")
        except OSError:
            pass
        raise


if __name__ == "__main__":
    main()
