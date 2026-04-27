# -*- coding: utf-8 -*-
"""
Stata MCP Server - 通过 Stata Automation COM 控制已打开的 Stata GUI
用法：python stata_mcp.py
版本：3.0
"""

import os
import asyncio
import threading
import queue
import concurrent.futures
import time
import hashlib
import re
from datetime import datetime
from pathlib import Path
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# ── 配置 ──────────────────────────────────────────────────────
MCP_DIR = Path(os.environ.get("STATA_MCP_DIR", Path(__file__).resolve().parents[2]))
STATA_DIR = Path(os.environ.get("STATA_DIR", MCP_DIR.parent))
STATA_EXE = Path(os.environ.get("STATA_EXE", STATA_DIR / "StataMP-64.exe"))
DOFILE_DIR = MCP_DIR / "runtime" / "dofiles"
LOG_DIR = MCP_DIR / "runtime" / "logs"
COM_PROG_ID = os.environ.get("STATA_COM_PROG_ID", "stata.StataOLEApp")

# 危险命令黑名单（安全防护）
DANGEROUS_COMMANDS = ["!", "shell", "erase", "rm ", "del ", "rmdir", "rd "]

# ── 初始化目录 ────────────────────────────────────────────────
DOFILE_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

try:
    import pythoncom
    import win32com.client
    PYWIN32_READY = True
    PYWIN32_ERROR = ""
except Exception as e:
    win32com = None
    PYWIN32_READY = False
    PYWIN32_ERROR = str(e)

STATA_ERROR = ""
SESSIONS = {}
LAST_SESSION_KEY = None


class StataSession:
    def __init__(self, key: str):
        self.key = key
        self.requests = queue.Queue()
        self.ready = threading.Event()
        self.error = ""
        self.thread = threading.Thread(target=self._run, name=f"StataSession-{Path(key).name}", daemon=True)
        self.thread.start()

    def _run(self):
        pythoncom.CoInitialize()
        self.app = None
        try:
            self.app = win32com.client.Dispatch(COM_PROG_ID)
            time.sleep(1.5)
            self.app.DoCommand('display "Stata MCP session ready"')
            self.error = ""
        except Exception as e:
            self.error = str(e)
        finally:
            self.ready.set()

        while True:
            command, future = self.requests.get()
            if command is None:
                break
            try:
                if self.app is None:
                    raise RuntimeError(self.error or "Stata COM session is not ready")
                self.app.DoCommand(command)
                future.set_result(True)
            except Exception as e:
                future.set_exception(e)

    def execute(self, command: str):
        if not self.ready.wait(timeout=20):
            raise RuntimeError("Stata COM session startup timed out")
        future = concurrent.futures.Future()
        self.requests.put((command, future))
        return future.result(timeout=180)

    def is_ready(self) -> bool:
        return self.ready.is_set() and self.app is not None


def _normalize_session_key(path: str) -> str:
    return str(Path(path).resolve())


def _get_session(session_key: str = None):
    global LAST_SESSION_KEY, STATA_ERROR
    if not PYWIN32_READY:
        STATA_ERROR = f"pywin32 未安装或无法导入：{PYWIN32_ERROR}"
        return None

    key = session_key or LAST_SESSION_KEY
    if not key:
        STATA_ERROR = "还没有 do 文件会话；请先运行一个 do 文件"
        return None

    if key not in SESSIONS:
        SESSIONS[key] = StataSession(key)
    LAST_SESSION_KEY = key

    session = SESSIONS[key]
    if not session.ready.wait(timeout=20) or not session.is_ready():
        STATA_ERROR = session.error or "Stata COM session is not ready"
        return None
    STATA_ERROR = ""
    return session


def _get_stata_app():
    return _get_session()

# ── MCP Server ────────────────────────────────────────────────
server = Server("stata-mcp")


@server.list_tools()
async def list_tools():
    return [
        Tool(
            name="stata_run",
            description="在已打开的 Stata GUI 中执行一条或多条 Stata 命令",
            inputSchema={
                "type": "object",
                "properties": {
                    "commands": {
                        "type": "string",
                        "description": "要执行的 Stata 命令，多行用换行符分隔"
                    }
                },
                "required": ["commands"]
            }
        ),
        Tool(
            name="stata_run_dofile",
            description="在已打开的 Stata GUI 中运行一个 do 文件",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "do 文件的完整路径，例如 D:/project/analysis.do"
                    }
                },
                "required": ["path"]
            }
        ),
        Tool(
            name="stata_write_dofile",
            description="将 Stata 代码写入 do 文件并保存到磁盘，返回文件路径",
            inputSchema={
                "type": "object",
                "properties": {
                    "content": {
                        "type": "string",
                        "description": "do 文件的完整内容"
                    },
                    "filename": {
                        "type": "string",
                        "description": "文件名（不含扩展名），留空则自动生成时间戳文件名"
                    }
                },
                "required": ["content"]
            }
        ),
        Tool(
            name="stata_append_dofile",
            description="向已有 do 文件追加内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "已有 do 文件的完整路径"
                    },
                    "content": {
                        "type": "string",
                        "description": "要追加的 Stata 代码内容"
                    }
                },
                "required": ["path", "content"]
            }
        ),
        Tool(
            name="stata_read_dofile",
            description="读取 do 文件内容，返回给 Claude 检查、解释或继续修改",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "do 文件的完整路径"
                    }
                },
                "required": ["path"]
            }
        ),
        Tool(
            name="stata_read_log",
            description="读取 Stata text log 文件内容，返回给 Claude 分析",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "log 文件的完整路径，留空则读取最新的 log 文件"
                    },
                    "tail_lines": {
                        "type": "integer",
                        "description": "只读取最后 N 行，留空则读取全部"
                    }
                }
            }
        ),
        Tool(
            name="stata_install_package",
            description="在已打开的 Stata GUI 中安装外部包，如 estout、ivreg2、rdrobust 等",
            inputSchema={
                "type": "object",
                "properties": {
                    "package": {
                        "type": "string",
                        "description": "包名称，例如 estout、ivreg2、rdrobust"
                    },
                    "source": {
                        "type": "string",
                        "description": "安装来源：ssc（默认）或 net",
                        "enum": ["ssc", "net"]
                    }
                },
                "required": ["package"]
            }
        ),
        Tool(
            name="stata_get_results",
            description="在 Stata GUI 中显示上一次命令的 r() 或 e() 存储结果",
            inputSchema={
                "type": "object",
                "properties": {
                    "type": {
                        "type": "string",
                        "description": "结果类型：r 或 e",
                        "enum": ["r", "e"]
                    }
                },
                "required": ["type"]
            }
        ),
        Tool(
            name="stata_get_data_info",
            description="在 Stata GUI 中显示当前数据集的基本信息",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="stata_get_data_schema",
            description="读取当前 Stata 数据集结构、缺失摘要和样本预览，返回给 Claude 分析",
            inputSchema={
                "type": "object",
                "properties": {
                    "sample_rows": {
                        "type": "integer",
                        "description": "样本预览行数，默认 20"
                    },
                    "include_codebook": {
                        "type": "boolean",
                        "description": "是否包含 codebook, compact，默认 true"
                    },
                    "include_sample": {
                        "type": "boolean",
                        "description": "是否包含样本预览，默认 true"
                    },
                    "include_missing": {
                        "type": "boolean",
                        "description": "是否包含缺失值摘要，默认 true"
                    }
                }
            }
        ),
        Tool(
            name="stata_status",
            description="检查 Stata MCP 服务器状态及目录配置",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        )
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict):
    try:
        loop = asyncio.get_event_loop()

        if name == "stata_run":
            commands = arguments.get("commands", "")
            safety_result = _check_safety(commands)
            if safety_result:
                return [TextContent(type="text", text=safety_result)]
            output = await loop.run_in_executor(None, lambda: _run_stata_commands(commands))
            return [TextContent(type="text", text=output)]

        elif name == "stata_run_dofile":
            path = arguments.get("path", "").replace("/", "\\")
            if not os.path.exists(path):
                return [TextContent(type="text", text=f"❌ 文件不存在：{path}")]
            session_key = _normalize_session_key(path)
            output = await loop.run_in_executor(None, lambda: _run_stata_commands(f'do "{path}"', session_key))
            return [TextContent(type="text", text=output)]

        elif name == "stata_write_dofile":
            content = arguments.get("content", "")
            filename = arguments.get("filename", "")
            output = _write_dofile(content, filename)
            return [TextContent(type="text", text=output)]

        elif name == "stata_append_dofile":
            path = arguments.get("path", "").replace("/", "\\")
            content = arguments.get("content", "")
            output = _append_dofile(path, content)
            return [TextContent(type="text", text=output)]

        elif name == "stata_read_dofile":
            path = arguments.get("path", "").replace("/", "\\")
            output = _read_dofile(path)
            return [TextContent(type="text", text=output)]

        elif name == "stata_read_log":
            path = arguments.get("path", "")
            tail_lines = arguments.get("tail_lines", None)
            output = _read_log(path, tail_lines)
            return [TextContent(type="text", text=output)]

        elif name == "stata_install_package":
            package = arguments.get("package", "")
            source = arguments.get("source", "ssc")
            output = await loop.run_in_executor(None, lambda: _install_package(package, source))
            return [TextContent(type="text", text=output)]

        elif name == "stata_get_results":
            result_type = arguments.get("type", "r")
            output = await loop.run_in_executor(None, lambda: _get_stored_results(result_type))
            return [TextContent(type="text", text=output)]

        elif name == "stata_get_data_info":
            output = await loop.run_in_executor(None, lambda: _get_data_info())
            return [TextContent(type="text", text=output)]

        elif name == "stata_get_data_schema":
            sample_rows = arguments.get("sample_rows", 20)
            include_codebook = arguments.get("include_codebook", True)
            include_sample = arguments.get("include_sample", True)
            include_missing = arguments.get("include_missing", True)
            output = await loop.run_in_executor(
                None,
                lambda: _get_data_schema(sample_rows, include_codebook, include_sample, include_missing)
            )
            return [TextContent(type="text", text=output)]

        elif name == "stata_status":
            return [TextContent(type="text", text=_status_text())]

        else:
            return [TextContent(type="text", text=f"❌ 未知工具：{name}")]

    except Exception as e:
        return [TextContent(type="text", text=f"❌ 执行出错：{str(e)}")]


# ── 工具函数 ──────────────────────────────────────────────────

def _stata_unavailable_message() -> str:
    if not PYWIN32_READY:
        return f"❌ pywin32 未安装或无法导入：{PYWIN32_ERROR}\n请运行：pip install pywin32"
    return ("❌ 当前没有可用的 Stata do 文件会话。\n"
            "请先用 stata_run_dofile 运行一个 do 文件；之后无路径命令会发送到最近的 do 会话。\n"
            f"COM ProgID：{COM_PROG_ID}\n"
            f"错误信息：{STATA_ERROR}")


def _status_text() -> str:
    pywin32_status = "✅ 可用" if PYWIN32_READY else f"❌ 不可用：{PYWIN32_ERROR}"
    session_lines = []
    for key, session in SESSIONS.items():
        state = "ready" if session.is_ready() else f"not ready: {session.error}"
        marker = " (最近)" if key == LAST_SESSION_KEY else ""
        session_lines.append(f"- {key}{marker}: {state}")
    if not session_lines:
        session_lines.append("- 暂无 do 文件会话")

    return (f"Stata MCP 状态：pywin32 {pywin32_status}\n"
            "后端模式：COM / Stata Automation\n"
            f"COM ProgID：{COM_PROG_ID}\n"
            f"Stata 路径：{STATA_DIR}\n"
            f"Do 文件目录：{DOFILE_DIR}\n"
            f"Log 目录：{LOG_DIR}\n"
            f"会话数量：{len(SESSIONS)}\n"
            f"最近会话：{LAST_SESSION_KEY or '无'}\n"
            "会话列表：\n" + "\n".join(session_lines))


def _check_safety(commands: str) -> str:
    for line in commands.strip().splitlines():
        line_lower = line.strip().lower()
        for dangerous in DANGEROUS_COMMANDS:
            if line_lower.startswith(dangerous):
                return (f"⚠️ 安全防护：检测到危险命令 `{line.strip().split()[0]}`，已拒绝执行。\n"
                        "如确实需要此操作，请直接在 Stata 界面手动执行。")
    return ""


def _run_stata_commands(commands: str, session_key: str = None) -> str:
    if not PYWIN32_READY:
        return _stata_unavailable_message()

    try:
        session = _get_session(session_key)
        if session is None:
            return _stata_unavailable_message()

        sent = 0
        for line in commands.splitlines():
            command = line.strip()
            if command:
                session.execute(command)
                sent += 1

        if sent == 0:
            return "❌ 没有可执行的 Stata 命令"
        return f"✅ 已发送 {sent} 条命令到 Stata 会话执行：{session.key}"
    except Exception as e:
        return f"❌ 命令发送失败：{str(e)}"


def _write_dofile(content: str, filename: str = "") -> str:
    try:
        if not filename:
            filename = f"analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        path = Path(filename)
        if not path.is_absolute():
            if not filename.endswith(".do"):
                filename += ".do"
            path = DOFILE_DIR / filename
        elif path.suffix.lower() != ".do":
            path = path.with_suffix(".do")
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
        return f"✅ do 文件已保存：{path}"
    except Exception as e:
        return f"❌ 写入 do 文件失败：{str(e)}"


def _append_dofile(path: str, content: str) -> str:
    try:
        p = Path(path)
        if not p.exists():
            return f"❌ 文件不存在：{path}"
        with open(p, "a", encoding="utf-8") as f:
            f.write("\n" + content)
        return f"✅ 内容已追加到：{path}"
    except Exception as e:
        return f"❌ 追加失败：{str(e)}"


def _read_dofile(path: str) -> str:
    try:
        p = Path(path)
        if not p.exists():
            return f"❌ do 文件不存在：{path}"
        content = p.read_text(encoding="utf-8", errors="ignore")
        return f"=== {p.name} ===\n{content}"
    except Exception as e:
        return f"❌ 读取 do 文件失败：{str(e)}"


def _read_log(path: str = "", tail_lines: int = None) -> str:
    try:
        if not path:
            logs = sorted(LOG_DIR.glob("*.log"), key=os.path.getmtime, reverse=True)
            if not logs:
                return "❌ 没有找到 log 文件"
            p = logs[0]
        else:
            p = Path(path)
            if not p.exists():
                return f"❌ log 文件不存在：{path}"

        content = p.read_text(encoding="utf-8", errors="ignore")

        if tail_lines:
            lines = content.splitlines()
            content = "\n".join(lines[-tail_lines:])

        return f"=== {p.name} ===\n{content}"
    except Exception as e:
        return f"❌ 读取 log 失败：{str(e)}"


def _install_package(package: str, source: str = "ssc") -> str:
    if source == "ssc":
        cmd = f"ssc install {package}, replace"
    else:
        cmd = f"net install {package}, replace"
    return _run_stata_commands(cmd)


def _get_stored_results(result_type: str) -> str:
    if result_type == "r":
        cmd = "return list"
    else:
        cmd = "ereturn list"
    result = _run_stata_commands(cmd)
    return f"{result}\n结果已显示在 Stata GUI；如需 Claude 分析，请使用 text log 后调用 stata_read_log。"


def _get_data_info() -> str:
    result = _run_stata_commands("describe")
    return f"{result}\n数据信息已显示在 Stata GUI；如需 Claude 分析，请使用 stata_get_data_schema。"


def _safe_log_stem(value: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_.-]+", "_", Path(value).stem).strip("_") or "stata_session"
    digest = hashlib.sha1(value.encode("utf-8", errors="ignore")).hexdigest()[:10]
    return f"{name}_{digest}"


def _get_data_schema(
    sample_rows: int = 20,
    include_codebook: bool = True,
    include_sample: bool = True,
    include_missing: bool = True,
) -> str:
    if not LAST_SESSION_KEY:
        return _stata_unavailable_message()

    session = _get_session()
    if session is None:
        return _stata_unavailable_message()

    try:
        sample_rows = int(sample_rows)
    except Exception:
        sample_rows = 20
    sample_rows = max(1, min(sample_rows, 1000))

    log_path = LOG_DIR / f"schema_{_safe_log_stem(session.key)}.log"
    stata_log_path = str(log_path).replace("\\", "/")
    commands = [
        "capture log close _all",
        f'log using "{stata_log_path}", text replace',
        "display \"=== Stata data schema snapshot ===\"",
        "describe",
    ]
    if include_codebook:
        commands.append("codebook, compact")
    if include_missing:
        commands.append("misstable summarize")
    if include_sample:
        commands.append(f"list in 1/{sample_rows}, abbreviate(20)")
    commands.extend([
        "notes",
        "label dir",
        "log close",
    ])

    try:
        for command in commands:
            session.execute(command)
        return _read_log(str(log_path))
    except Exception as e:
        return f"❌ 读取数据结构失败：{str(e)}"


# ── 启动 ──────────────────────────────────────────────────────
async def main():
    async with stdio_server() as streams:
        await server.run(
            streams[0],
            streams[1],
            server.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())
