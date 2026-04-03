"""
环境与启动预检、报错文本诊断（不 import app/main，避免拉起重依赖后才失败）。
"""
from __future__ import annotations

import re
import socket
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator


def src_dir() -> Path:
    return Path(__file__).resolve().parent


def repo_root() -> Path:
    return src_dir().parent


def requirements_path() -> Path:
    return src_dir() / "requirements.txt"


PIP_TO_IMPORT: dict[str, str] = {
    "fastapi": "fastapi",
    "uvicorn": "uvicorn",
    "jinja2": "jinja2",
    "playwright": "playwright",
    "requests": "requests",
    "beautifulsoup4": "bs4",
    "python-dotenv": "dotenv",
    "streamlit": "streamlit",
    "bcrypt": "bcrypt",
    "itsdangerous": "itsdangerous",
}


def _parse_requirement_names(req_file: Path) -> list[str]:
    if not req_file.is_file():
        return []
    names: list[str] = []
    for line in req_file.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        m = re.match(r"^([a-zA-Z0-9_.-]+)", line)
        if m:
            names.append(m.group(1).lower().replace("_", "-"))
    return names


def _import_name_for_pip(pkg: str) -> str:
    return PIP_TO_IMPORT.get(pkg.lower(), pkg.lower().replace("-", "_"))


@dataclass
class CheckItem:
    id: str
    ok: bool
    title: str
    detail: str
    severity: str
    fix_commands: list[str] = field(default_factory=list)


@dataclass
class Diagnosis:
    code: str
    title: str
    summary: str
    hints: list[str]
    commands: list[str]


def diagnose_error_text(text: str) -> Diagnosis | None:
    if not (text or "").strip():
        return None
    t = text

    def dm(
        code: str,
        title: str,
        summary: str,
        hints: list[str],
        commands: list[str],
    ) -> Diagnosis:
        return Diagnosis(code, title, summary, hints, commands)

    m = re.search(r"No module named\s+['\"]([^'\"]+)['\"]", t, re.I)
    if m:
        mod = m.group(1)
        return dm(
            "MISSING_MODULE",
            "缺少 Python 包",
            f"当前环境缺少模块（疑似: {mod}）。",
            [
                "使用助手「安装 pip 依赖」安装 requirements.txt。",
                "确认本窗口使用的 Python 与启动主程序时一致（建议始终用 run_helper.bat）。",
            ],
            [f'{sys.executable} -m pip install -r "{requirements_path()}"'],
        )

    rules: list[tuple[re.Pattern[str], Diagnosis]] = [
        (
            re.compile(
                r"Python not found|不是内部或外部命令|\'python\' 不是|python不是内部或外部命令",
                re.I,
            ),
            dm(
                "NO_PYTHON",
                "未找到 Python",
                "系统里没有可用的 python 命令，或尚未安装 Python。",
                [
                    "从 python.org 安装 Python 3.10+，勾选 Add Python to PATH。",
                    "安装后重新打开本助手。",
                ],
                [],
            ),
        ),
        (
            re.compile(r"ModuleNotFoundError", re.I),
            dm(
                "MISSING_MODULE_GENERIC",
                "导入失败",
                "某个 Python 模块找不到，多为依赖未装全。",
                ["点击「安装 pip 依赖」。", "若使用虚拟环境，请先激活再启动助手。"],
                [],
            ),
        ),
        (
            re.compile(
                r"playwright install|Executable doesn\'t exist|BrowserType\.launch|install msedge",
                re.I,
            ),
            dm(
                "PLAYWRIGHT_EDGE",
                "Playwright / Edge",
                "Playwright 浏览器组件未就绪或无法启动 Edge 通道。",
                [
                    "点击「安装 Playwright 浏览器」。",
                    "确认本机已安装 Microsoft Edge。",
                ],
                [f'"{sys.executable}" -m playwright install msedge'],
            ),
        ),
        (
            re.compile(
                r"Address already in use|10048|Only one usage of each socket address|通常每个套接字地址",
                re.I,
            ),
            dm(
                "PORT_IN_USE",
                "端口 8001 被占用",
                "8001 已被占用，无法再起一个 Web 服务。",
                ["关掉之前运行主程序的黑窗口，或结束占用端口的进程。"],
                [],
            ),
        ),
        (
            re.compile(
                r"sqlite3\.OperationalError|database is locked|unable to open database file",
                re.I,
            ),
            dm(
                "SQLITE",
                "数据库文件异常",
                "无法写入或锁定 SQLite（user.db）。",
                ["确认 src 目录可写。", "关闭其他正在运行本程序的实例。"],
                [],
            ),
        ),
        (
            re.compile(
                r"无法定位搜索框|Cookie 已过期|cookie_expired|debug_searchbox_final\.png",
                re.I,
            ),
            dm(
                "OWA_COOKIE",
                "邮箱或 Cookie",
                "Cookie 失效或页面结构变化。",
                ["在网页「设置」里重新粘贴 Cookie JSON。", "参见维护指南。"],
                [],
            ),
        ),
        (
            re.compile(r"缺少 base_url 或 api_key", re.I),
            dm(
                "LLM_CONFIG",
                "大模型未配置",
                "未配置 API base_url 或 api_key。",
                ["在网页「设置」中填写密钥与接口地址。"],
                [],
            ),
        ),
        (
            re.compile(r"Failed to launch browser after", re.I),
            dm(
                "BROWSER_LAUNCH",
                "浏览器启动失败",
                "Playwright 无法拉起浏览器。",
                ["先安装 Playwright 的 msedge 组件。", "检查 Edge 是否能手动打开。"],
                [f'"{sys.executable}" -m playwright install msedge'],
            ),
        ),
    ]
    for cre, diag in rules:
        if cre.search(t):
            return diag
    return dm(
        "UNKNOWN",
        "未能精确定位",
        "未命中内置规则。可把完整报错发给维护者，或查看 docs/README.md。",
        [
            "依次尝试：安装 pip 依赖 → 安装 Playwright 浏览器 → 重新检查。",
            "确认用 run_helper.bat 启动，且未破坏项目文件夹结构。",
        ],
        [],
    )


def check_project_layout() -> CheckItem:
    root = repo_root()
    src = src_dir()
    need = [src / "app.py", src / "main.py", requirements_path()]
    missing = [p.relative_to(root) for p in need if not p.is_file()]
    if missing:
        return CheckItem(
            id="layout",
            ok=False,
            title="项目文件不完整",
            detail="缺少: " + ", ".join(str(m) for m in missing),
            severity="error",
            fix_commands=[],
        )
    return CheckItem(
        id="layout",
        ok=True,
        title="项目结构",
        detail=f"根目录 {root}",
        severity="ok",
        fix_commands=[],
    )


def check_python_version() -> CheckItem:
    v = sys.version_info
    if v.major < 3 or (v.major == 3 and v.minor < 10):
        return CheckItem(
            id="python_ver",
            ok=False,
            title="Python 版本",
            detail=f"当前 {v.major}.{v.minor}，建议 3.10+。",
            severity="warn",
            fix_commands=[],
        )
    return CheckItem(
        id="python_ver",
        ok=True,
        title="Python 版本",
        detail=f"{v.major}.{v.minor}.{v.micro} — {sys.executable}",
        severity="ok",
        fix_commands=[],
    )


def check_venv() -> CheckItem:
    root = repo_root()
    p_venv = root / ".venv" / "Scripts" / "python.exe"
    p_alt = root / "venv" / "Scripts" / "python.exe"
    if p_venv.is_file():
        same = p_venv.resolve() == Path(sys.executable).resolve()
        return CheckItem(
            id="venv",
            ok=True,
            title="虚拟环境",
            detail=".venv 已存在"
            + ("，当前解释器已在使用。" if same else "；当前未使用 .venv，建议用 run_helper.bat 启动。"),
            severity="ok" if same else "warn",
            fix_commands=[],
        )
    if p_alt.is_file():
        same = p_alt.resolve() == Path(sys.executable).resolve()
        return CheckItem(
            id="venv",
            ok=True,
            title="虚拟环境",
            detail="venv 已存在"
            + ("，当前解释器已在使用。" if same else "；建议用 run_helper.bat 启动。"),
            severity="ok" if same else "warn",
            fix_commands=[],
        )
    return CheckItem(
        id="venv",
        ok=True,
        title="虚拟环境",
        detail="未检测到 .venv / venv，正使用系统 Python。",
        severity="warn",
        fix_commands=[f'"{sys.executable}" -m venv .venv'],
    )


def check_src_writable() -> CheckItem:
    p = src_dir() / ".write_test_startup_helper"
    try:
        p.write_text("ok", encoding="utf-8")
        p.unlink(missing_ok=True)
        return CheckItem(
            id="writable",
            ok=True,
            title="src 目录可写",
            detail="可创建 user.db。",
            severity="ok",
            fix_commands=[],
        )
    except OSError as e:
        return CheckItem(
            id="writable",
            ok=False,
            title="src 目录不可写",
            detail=str(e),
            severity="error",
            fix_commands=[],
        )


def check_pip_imports() -> CheckItem:
    req = requirements_path()
    if not req.is_file():
        return CheckItem(
            id="pip_imports",
            ok=False,
            title="依赖检查",
            detail="找不到 requirements.txt",
            severity="error",
            fix_commands=[],
        )
    pkgs = _parse_requirement_names(req)
    missing: list[str] = []
    for pkg in pkgs:
        im = _import_name_for_pip(pkg)
        try:
            __import__(im)
        except ImportError:
            missing.append(f"{pkg} (import {im})")
    if missing:
        return CheckItem(
            id="pip_imports",
            ok=False,
            title="Python 依赖包",
            detail="未安装: " + "; ".join(missing[:8]) + (" …" if len(missing) > 8 else ""),
            severity="error",
            fix_commands=[
                f'"{sys.executable}" -m pip install -U pip',
                f'"{sys.executable}" -m pip install -r "{requirements_path()}"',
            ],
        )
    return CheckItem(
        id="pip_imports",
        ok=True,
        title="Python 依赖包",
        detail="requirements 中的包均可 import。",
        severity="ok",
        fix_commands=[],
    )


def check_playwright_edge() -> CheckItem:
    """
    在独立子进程里做 Playwright 探测。避免在 Tk 后台线程里直接调用 sync_playwright
    （非线程安全，可能导致整个进程闪退、窗口秒关）。
    """
    script = (
        "import sys\n"
        "try:\n"
        "    from playwright.sync_api import sync_playwright\n"
        "except ImportError as e:\n"
        "    print('ImportError:', e, file=sys.stderr)\n"
        "    sys.exit(2)\n"
        "try:\n"
        "    with sync_playwright() as p:\n"
        "        b = p.chromium.launch(channel='msedge', headless=True)\n"
        "        b.close()\n"
        "except Exception as e:\n"
        "    print(str(e), file=sys.stderr)\n"
        "    sys.exit(1)\n"
    )
    try:
        r = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True,
            text=True,
            timeout=120,
            cwd=str(repo_root()),
            encoding="utf-8",
            errors="replace",
        )
    except subprocess.TimeoutExpired:
        return CheckItem(
            id="playwright",
            ok=False,
            title="Playwright / Edge",
            detail="检测超时（>120s），请检查网络或稍后重试。",
            severity="error",
            fix_commands=[f'"{sys.executable}" -m playwright install msedge'],
        )
    except OSError as e:
        return CheckItem(
            id="playwright",
            ok=False,
            title="Playwright / Edge",
            detail=f"无法启动子进程检测: {e}",
            severity="error",
            fix_commands=[],
        )
    if r.returncode == 0:
        return CheckItem(
            id="playwright",
            ok=True,
            title="Playwright / Edge",
            detail="msedge 通道 headless 启动成功（子进程检测）。",
            severity="ok",
            fix_commands=[],
        )
    err = (r.stderr or r.stdout or "").strip()[:500] or f"exit {r.returncode}"
    if r.returncode == 2:
        return CheckItem(
            id="playwright",
            ok=False,
            title="Playwright / Edge",
            detail=f"未安装 playwright 包: {err}",
            severity="error",
            fix_commands=[],
        )
    return CheckItem(
        id="playwright",
        ok=False,
        title="Playwright / Edge",
        detail=err,
        severity="error",
        fix_commands=[f'"{sys.executable}" -m playwright install msedge'],
    )


def check_port_8001() -> CheckItem:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.settimeout(0.4)
        r = sock.connect_ex(("127.0.0.1", 8001))
        if r == 0:
            return CheckItem(
                id="port",
                ok=True,
                title="端口 8001",
                detail="已有程序监听 8001（若是上次主程序，可先关闭再开新的）。",
                severity="warn",
                fix_commands=[],
            )
        return CheckItem(
            id="port",
            ok=True,
            title="端口 8001",
            detail="当前空闲。",
            severity="ok",
            fix_commands=[],
        )
    finally:
        sock.close()


def run_all_checks() -> list[CheckItem]:
    items = [
        check_project_layout(),
        check_python_version(),
        check_venv(),
        check_src_writable(),
        check_pip_imports(),
    ]
    if items[-1].ok:
        items.append(check_playwright_edge())
    else:
        items.append(
            CheckItem(
                id="playwright",
                ok=False,
                title="Playwright / Edge",
                detail="已跳过（先装 pip 依赖）。",
                severity="warn",
                fix_commands=[f'"{sys.executable}" -m playwright install msedge'],
            )
        )
    items.append(check_port_8001())
    return items


def suggested_pip_install_command() -> list[str]:
    return [sys.executable, "-m", "pip", "install", "-r", str(requirements_path())]


def suggested_playwright_install_command() -> list[str]:
    return [sys.executable, "-m", "playwright", "install", "msedge"]


def suggested_venv_create_command() -> list[str]:
    return [sys.executable, "-m", "venv", str(repo_root() / ".venv")]


def iter_subprocess_lines(
    argv: list[str],
    *,
    cwd: Path | None = None,
) -> Iterator[str]:
    cwd = cwd or repo_root()
    proc = subprocess.Popen(
        argv,
        cwd=str(cwd),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    assert proc.stdout
    for line in proc.stdout:
        yield line.rstrip("\n\r")
    proc.wait()
    yield f"[exit code {proc.returncode}]"

