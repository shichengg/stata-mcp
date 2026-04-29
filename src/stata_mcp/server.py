# -*- coding: utf-8 -*-
"""
Stata MCP Server - 通过 Stata Automation COM 控制已打开的 Stata GUI
用法：python stata_mcp.py
版本：3.0
"""

import os
import json
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
COM_PROG_ID = os.environ.get("STATA_COM_PROG_ID", "stata.StataOLEApp")

LOCAL_MCP_DIR_NAME = ".stata-mcp"
CACHE_DIR_NAME = "cache"
DOFILES_DIR_NAME = "dofiles"
TASK_REGISTRY_FILE = "task_registry.json"

# 危险命令黑名单（安全防护）
DANGEROUS_COMMANDS = ["!", "shell", "erase", "rm ", "del ", "rmdir", "rd "]

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
SESSION_REGISTRIES = {}
KNOWN_REGISTRY_PATHS = set()
COM_DISCONNECT_PATTERNS = (
    "RPC 服务器不可用",
    "RPC server is unavailable",
    "The RPC server is unavailable",
    "对象已断开",
    "object disconnected",
    "server threw an exception",
    "被调用的对象已与其客户端断开连接",
)


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


# ── MCP Server ────────────────────────────────────────────────
server = Server("stata-mcp")


@server.list_tools()
async def list_tools():
    return [
        Tool(
            name="stata_run",
            description="在最近的 Stata MCP session_id 对应的 Stata GUI 中执行一条或多条 Stata 命令；可先用 stata_session(action='set_recent') 切换目标 session。",
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
            description=(
                "在 Stata GUI 中运行一个 do 文件。该工具以 session_id 表示一个复现任务/同一个 Stata GUI；"
                "同一任务的原始 do、检查 do、续跑 do 应使用同一个 session_id。工具会自动在 do 文件所在目录创建 text log，"
                "默认每个 do 一个日志，文件名包含 do 名和 session_id；不需要在用户 do 文件里手写 log using，"
                "也不会生成临时 wrapper do 文件。返回 session_id、source_do 和本次 do 的 log_path。"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "do 文件的完整路径，例如 D:/project/analysis.do"
                    },
                    "session_id": {
                        "type": "string",
                        "description": "可选；同一复现任务固定使用同一个 session_id。不传时兼容旧行为：用 do 文件路径作为 session key"
                    },
                    "role": {
                        "type": "string",
                        "description": "该 do 在 session 中的角色：entry/source/auxiliary，默认 entry",
                        "enum": ["entry", "source", "auxiliary"]
                    },
                    "log_mode": {
                        "type": "string",
                        "description": "日志写入模式：replace 或 append，默认 replace",
                        "enum": ["replace", "append"]
                    }
                },
                "required": ["path"]
            }
        ),
        Tool(
            name="stata_session",
            description="管理 Stata MCP session：列出、查询、销毁、切换最近 session。session_id 代表同一复现任务/同一个 Stata GUI，并关联 entry/source/current do、每个 do 的 log_paths 和 last_log_path。",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "description": "操作：list/get/destroy/set_recent",
                        "enum": ["list", "get", "destroy", "set_recent"]
                    },
                    "session_id": {
                        "type": "string",
                        "description": "get/destroy/set_recent 需要指定的 session_id"
                    }
                },
                "required": ["action"]
            }
        ),
        Tool(
            name="stata_write_dofile",
            description="将 Stata 代码写入 do 文件并保存到磁盘，返回文件路径；相对文件名会写入最近 session 项目的 .stata-mcp/dofiles/，不会使用 MCP runtime。",
            inputSchema={
                "type": "object",
                "properties": {
                    "content": {
                        "type": "string",
                        "description": "do 文件的完整内容"
                    },
                    "filename": {
                        "type": "string",
                        "description": "文件名（不含扩展名）或绝对路径；留空则自动生成时间戳文件名"
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
            description="读取 Stata text log。path 留空时读取最近 session 的 last_log_path；推荐 output_format='dict'，它会把日志解析成命令-结果对，便于 AI 判断报错位置。",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "log 文件的完整路径；留空则读取最近 session 的 last_log_path"
                    },
                    "tail_lines": {
                        "type": "integer",
                        "description": "只读取最后 N 行，留空则读取全部"
                    },
                    "output_format": {
                        "type": "string",
                        "description": "输出格式：full=完整文本，core=去除日志框架行，dict=JSON 命令-结果对",
                        "enum": ["full", "core", "dict"]
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
            description="检查 Stata MCP 服务器状态、内存 session 和项目局部 .stata-mcp/cache/task_registry.json 中记录的 session/do/log 关系",
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
            path = _normalize_path_arg(arguments.get("path", ""))
            if not os.path.exists(path):
                return [TextContent(type="text", text=f"❌ 文件不存在：{path}")]
            session_id = arguments.get("session_id", "")
            role = arguments.get("role", "entry")
            log_mode = arguments.get("log_mode", "replace")
            output = await loop.run_in_executor(None, lambda: _run_dofile_session(path, session_id, role, log_mode))
            return [TextContent(type="text", text=output)]

        elif name == "stata_session":
            action = arguments.get("action", "list")
            session_id = arguments.get("session_id", "")
            output = _session_tool(action, session_id)
            return [TextContent(type="text", text=output)]

        elif name == "stata_write_dofile":
            content = arguments.get("content", "")
            filename = arguments.get("filename", "")
            output = _write_dofile(content, filename)
            return [TextContent(type="text", text=output)]

        elif name == "stata_append_dofile":
            path = _normalize_path_arg(arguments.get("path", ""))
            content = arguments.get("content", "")
            output = _append_dofile(path, content)
            return [TextContent(type="text", text=output)]

        elif name == "stata_read_dofile":
            path = _normalize_path_arg(arguments.get("path", ""))
            output = _read_dofile(path)
            return [TextContent(type="text", text=output)]

        elif name == "stata_read_log":
            path = arguments.get("path", "")
            tail_lines = arguments.get("tail_lines", None)
            output_format = arguments.get("output_format", "full")
            output = _read_log(path, tail_lines, output_format)
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

def _normalize_path_arg(path: str) -> str:
    return str(path or "").replace("/", "\\")


def _stata_path(path: Path) -> str:
    return str(path).replace("\\", "/")


def _normalize_session_key(path: str) -> str:
    return str(Path(path).resolve())


def _safe_session_id(value: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "")).strip("_")
    return safe or "stata_session"


def _safe_file_stem(value: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "")).strip("_")
    return safe or "stata_file"


def _project_mcp_dir(do_file_path: Path) -> Path:
    return do_file_path.parent / LOCAL_MCP_DIR_NAME


def _task_registry_path(do_file_path: Path) -> Path:
    return _project_mcp_dir(do_file_path) / CACHE_DIR_NAME / TASK_REGISTRY_FILE


def _project_dofiles_dir(do_file_path: Path) -> Path:
    return _project_mcp_dir(do_file_path) / DOFILES_DIR_NAME


def _load_task_registry(registry_path: Path) -> dict:
    KNOWN_REGISTRY_PATHS.add(str(registry_path))
    if not registry_path.exists():
        return {"sessions": {}}
    try:
        data = json.loads(registry_path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {"sessions": {}}
        data.setdefault("sessions", {})
        return data
    except Exception:
        return {"sessions": {}}


def _save_task_registry(registry_path: Path, registry: dict) -> None:
    registry_path.parent.mkdir(parents=True, exist_ok=True)
    registry_path.write_text(json.dumps(registry, ensure_ascii=False, indent=2), encoding="utf-8")
    KNOWN_REGISTRY_PATHS.add(str(registry_path))


def _get_log_file_path(do_file_path: Path, session_id: str) -> Path:
    do_stem = _safe_file_stem(do_file_path.stem)
    safe_session = _safe_session_id(session_id)
    return do_file_path.parent / f"{do_stem}_{safe_session}_mcp.log"


def _session_key_for_dofile(path: Path, session_id: str) -> str:
    if session_id:
        return _safe_session_id(session_id)
    return _normalize_session_key(str(path))


def _get_session(session_key: str = None, force_new: bool = False):
    global LAST_SESSION_KEY, STATA_ERROR
    if not PYWIN32_READY:
        STATA_ERROR = f"pywin32 未安装或无法导入：{PYWIN32_ERROR}"
        return None

    key = session_key or LAST_SESSION_KEY
    if not key:
        STATA_ERROR = "还没有 do 文件会话；请先运行一个 do 文件"
        return None

    if force_new and key in SESSIONS:
        del SESSIONS[key]
    if key not in SESSIONS:
        SESSIONS[key] = StataSession(key)
    LAST_SESSION_KEY = key

    session = SESSIONS[key]
    if not session.ready.wait(timeout=20) or not session.is_ready():
        STATA_ERROR = session.error or "Stata COM session is not ready"
        return None
    STATA_ERROR = ""
    return session


def _is_com_disconnect_error(error: Exception) -> bool:
    text = str(error)
    return any(pattern.lower() in text.lower() for pattern in COM_DISCONNECT_PATTERNS)


def _cleanup_disconnected_session(session_id: str, delete_logs: bool = True, extra_log_paths: list[Path] = None) -> None:
    global LAST_SESSION_KEY
    if session_id in SESSIONS:
        del SESSIONS[session_id]

    registry_path, registry = _registry_for_session(session_id)
    session_meta = registry.get("sessions", {}).get(session_id, {}) if registry_path else {}
    if delete_logs:
        log_paths = set(session_meta.get("log_paths", {}).values())
        if session_meta.get("last_log_path"):
            log_paths.add(session_meta["last_log_path"])
        for log_path in extra_log_paths or []:
            log_paths.add(str(log_path))

        safe_session = _safe_session_id(session_id)
        search_dirs = {Path(log_path).parent for log_path in log_paths if log_path}
        working_dir = session_meta.get("working_dir")
        if working_dir:
            search_dirs.add(Path(working_dir))
        for directory in search_dirs:
            try:
                for historical_log in directory.glob(f"*_{safe_session}_mcp.log"):
                    log_paths.add(str(historical_log))
            except Exception:
                pass

        for log_path in log_paths:
            try:
                Path(log_path).unlink(missing_ok=True)
            except Exception:
                pass

    if registry_path and session_id in registry.get("sessions", {}):
        del registry["sessions"][session_id]
        _save_task_registry(registry_path, registry)
    if LAST_SESSION_KEY == session_id:
        LAST_SESSION_KEY = None


def _execute_with_reconnect(session_key: str, command: str, extra_log_paths: list[Path] = None) -> bool:
    session = _get_session(session_key)
    if session is None:
        raise RuntimeError(STATA_ERROR or "Stata COM session is not ready")
    try:
        session.execute(command)
        return False
    except Exception as e:
        if not _is_com_disconnect_error(e):
            raise
        _cleanup_disconnected_session(session_key, extra_log_paths=extra_log_paths)
        session = _get_session(session_key, force_new=True)
        if session is None:
            raise RuntimeError(STATA_ERROR or "Stata COM session reconnect failed") from e
        return True


def _get_stata_app():
    return _get_session()


def _stata_unavailable_message() -> str:
    if not PYWIN32_READY:
        return f"❌ pywin32 未安装或无法导入：{PYWIN32_ERROR}\n请运行：pip install pywin32"
    return ("❌ 当前没有可用的 Stata do 文件会话。\n"
            "请先用 stata_run_dofile 运行一个 do 文件；之后无路径命令会发送到最近的 session。\n"
            f"COM ProgID：{COM_PROG_ID}\n"
            f"错误信息：{STATA_ERROR}")


def _registry_for_session(session_id: str):
    registry_path = SESSION_REGISTRIES.get(session_id)
    if registry_path:
        path = Path(registry_path)
        return path, _load_task_registry(path)

    for item in sorted(KNOWN_REGISTRY_PATHS):
        path = Path(item)
        registry = _load_task_registry(path)
        if session_id in registry.get("sessions", {}):
            SESSION_REGISTRIES[session_id] = str(path)
            return path, registry
    return None, {"sessions": {}}


def _session_metadata(session_id: str) -> dict:
    _, registry = _registry_for_session(session_id)
    return registry.get("sessions", {}).get(session_id, {})


def _update_task_registry(session_id: str, registry_path: Path, **metadata) -> dict:
    registry = _load_task_registry(registry_path)
    sessions = registry.setdefault("sessions", {})
    session = sessions.setdefault(session_id, {"session_id": session_id})
    session.update(metadata)
    session["updated_at"] = datetime.now().isoformat(timespec="seconds")
    SESSION_REGISTRIES[session_id] = str(registry_path)
    _save_task_registry(registry_path, registry)
    return session


def _register_dofile_run(path: Path, session_id: str, role: str, log_path: Path) -> dict:
    registry_path = _task_registry_path(path)
    registry = _load_task_registry(registry_path)
    sessions = registry.setdefault("sessions", {})
    session = sessions.setdefault(session_id, {"session_id": session_id})

    if role == "entry" or not session.get("entry_do"):
        session["entry_do"] = str(path)
    if role == "source" or not session.get("source_do"):
        session["source_do"] = str(path)
    session["current_do"] = str(path)
    session["working_dir"] = str(path.parent)
    session["last_log_path"] = str(log_path)
    session["status"] = "ready"
    session.setdefault("log_paths", {})[path.name] = str(log_path)
    session.setdefault("role_history", []).append({
        "role": role,
        "path": str(path),
        "log_path": str(log_path),
        "at": datetime.now().isoformat(timespec="seconds")
    })
    SESSION_REGISTRIES[session_id] = str(registry_path)
    _save_task_registry(registry_path, registry)
    return session


def _status_text() -> str:
    pywin32_status = "✅ 可用" if PYWIN32_READY else f"❌ 不可用：{PYWIN32_ERROR}"
    session_lines = []
    for key, session in SESSIONS.items():
        state = "ready" if session.is_ready() else f"not ready: {session.error}"
        marker = " (最近)" if key == LAST_SESSION_KEY else ""
        meta = _session_metadata(key)
        if meta:
            session_lines.append(
                f"- {key}{marker}: {state}\n"
                f"  entry_do: {meta.get('entry_do', '无')}\n"
                f"  source_do: {meta.get('source_do', '无')}\n"
                f"  current_do: {meta.get('current_do', '无')}\n"
                f"  last_log_path: {meta.get('last_log_path', '无')}\n"
                f"  log_paths: {json.dumps(meta.get('log_paths', {}), ensure_ascii=False)}"
            )
        else:
            session_lines.append(f"- {key}{marker}: {state}")
    if not session_lines:
        session_lines.append("- 暂无内存 do 文件会话")

    registry_lines = []
    for item in sorted(KNOWN_REGISTRY_PATHS):
        path = Path(item)
        registry = _load_task_registry(path)
        session_ids = sorted(registry.get("sessions", {}).keys())
        registry_lines.append(f"- {path}: {', '.join(session_ids) if session_ids else '无 session'}")
    if not registry_lines:
        registry_lines.append("- 暂无已知项目局部注册表；运行 stata_run_dofile 后会在 do 目录创建 .stata-mcp/cache/task_registry.json")

    return (f"Stata MCP 状态：pywin32 {pywin32_status}\n"
            "后端模式：COM / Stata Automation\n"
            f"COM ProgID：{COM_PROG_ID}\n"
            f"Stata 路径：{STATA_DIR}\n"
            f"MCP 程序目录：{MCP_DIR}\n"
            "缓存策略：项目/do 文件目录下的 .stata-mcp/cache/task_registry.json；不使用全局缓存或 MCP runtime\n"
            f"会话数量：{len(SESSIONS)}\n"
            f"最近会话：{LAST_SESSION_KEY or '无'}\n"
            "会话列表：\n" + "\n".join(session_lines) + "\n"
            "已知任务注册表：\n" + "\n".join(registry_lines))


def _session_tool(action: str, session_id: str = "") -> str:
    global LAST_SESSION_KEY
    action = action or "list"
    session_id = _safe_session_id(session_id) if session_id else ""

    if action == "list":
        data = {
            "recent_session": LAST_SESSION_KEY,
            "memory_sessions": {
                key: {
                    "state": "ready" if session.is_ready() else f"not ready: {session.error}",
                    "metadata": _session_metadata(key)
                }
                for key, session in SESSIONS.items()
            },
            "registries": {}
        }
        for item in sorted(KNOWN_REGISTRY_PATHS):
            path = Path(item)
            data["registries"][str(path)] = _load_task_registry(path)
        return json.dumps(data, ensure_ascii=False, indent=2)

    if not session_id:
        return f"❌ action={action} 需要 session_id"

    if action == "get":
        meta = _session_metadata(session_id)
        state = "not in memory"
        if session_id in SESSIONS:
            session = SESSIONS[session_id]
            state = "ready" if session.is_ready() else f"not ready: {session.error}"
        return json.dumps({"session_id": session_id, "state": state, "metadata": meta}, ensure_ascii=False, indent=2)

    if action == "set_recent":
        if session_id not in SESSIONS and not _session_metadata(session_id):
            return f"❌ 未找到 session_id：{session_id}"
        LAST_SESSION_KEY = session_id
        return f"✅ 最近 session 已切换为：{session_id}"

    if action == "destroy":
        if session_id in SESSIONS:
            del SESSIONS[session_id]
        registry_path, registry = _registry_for_session(session_id)
        if registry_path and session_id in registry.get("sessions", {}):
            registry["sessions"][session_id]["status"] = "destroyed"
            registry["sessions"][session_id]["updated_at"] = datetime.now().isoformat(timespec="seconds")
            _save_task_registry(registry_path, registry)
        if LAST_SESSION_KEY == session_id:
            LAST_SESSION_KEY = None
        return f"✅ 已从 MCP 内存移除 session：{session_id}；未关闭用户的 Stata GUI"

    return f"❌ 未知 session action：{action}"


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
                reconnected = _execute_with_reconnect(session.key, command)
                if reconnected:
                    return ("⚠️ 检测到 Stata GUI/COM 断开；已删除该 session 的旧日志和缓存，"
                            f"并重新初始化空 Stata GUI：{session.key}。未重放中断命令。")
                sent += 1

        if sent == 0:
            return "❌ 没有可执行的 Stata 命令"
        return f"✅ 已发送 {sent} 条命令到 Stata 会话执行：{session.key}"
    except Exception as e:
        return f"❌ 命令发送失败：{str(e)}"


def _run_dofile_session(path: str, session_id: str = "", role: str = "entry", log_mode: str = "replace") -> str:
    if not PYWIN32_READY:
        return _stata_unavailable_message()

    do_path = Path(path).resolve()
    role = role if role in {"entry", "source", "auxiliary"} else "entry"
    log_mode = log_mode if log_mode in {"replace", "append"} else "replace"
    session_key = _session_key_for_dofile(do_path, session_id)
    log_path = _get_log_file_path(do_path, session_key)

    output = _run_dofile_with_log(do_path, session_key, log_path, log_mode)
    if output.startswith("❌") or output.startswith("⚠️"):
        return output

    _register_dofile_run(do_path, session_key, role, log_path)
    return ("✅ 已发送 do 文件到 Stata 会话执行\n"
            f"session_id: {session_key}\n"
            f"role: {role}\n"
            f"source_do: {do_path}\n"
            f"log_path: {log_path}\n"
            f"log_mode: {log_mode}\n"
            "notes: 未生成 wrapper do 文件；日志由 MCP 直接发送 log/do/log close 命令创建")


def _run_dofile_with_log(path: Path, session_key: str, log_path: Path, log_mode: str) -> str:
    try:
        session = _get_session(session_key)
        if session is None:
            return _stata_unavailable_message()

        log_path.parent.mkdir(parents=True, exist_ok=True)
        commands = [
            "capture log close _all",
            f'log using "{_stata_path(log_path)}", text {log_mode}',
            f'cd "{_stata_path(path.parent)}"',
            f'do "{_stata_path(path)}"',
            "log close",
        ]
        for command in commands:
            reconnected = _execute_with_reconnect(session_key, command, extra_log_paths=[log_path])
            if reconnected:
                return ("⚠️ 检测到 Stata GUI/COM 断开；已删除该 session 的旧日志和缓存，"
                        f"并重新初始化空 Stata GUI：{session_key}。未重放中断 do 文件。")
        return "✅ do 文件命令已发送"
    except Exception as e:
        return f"❌ do 文件发送失败：{str(e)}"


def _write_dofile(content: str, filename: str = "") -> str:
    try:
        if not filename:
            filename = f"analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        path = Path(_normalize_path_arg(filename))
        if not path.is_absolute():
            if not filename.endswith(".do"):
                filename += ".do"
            base_dir = _recent_project_dofiles_dir()
            if base_dir is None:
                return "❌ 未指定绝对路径，且没有最近 session；请传入绝对路径或先运行 stata_run_dofile 建立项目 session"
            path = base_dir / filename
        elif path.suffix.lower() != ".do":
            path = path.with_suffix(".do")
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
        return f"✅ do 文件已保存：{path}"
    except Exception as e:
        return f"❌ 写入 do 文件失败：{str(e)}"


def _recent_project_dofiles_dir():
    if not LAST_SESSION_KEY:
        return None
    meta = _session_metadata(LAST_SESSION_KEY)
    current_do = meta.get("current_do") or meta.get("entry_do") or meta.get("source_do")
    if not current_do:
        return None
    return _project_dofiles_dir(Path(current_do))


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


def _read_log(path: str = "", tail_lines: int = None, output_format: str = "full") -> str:
    try:
        p = Path(_normalize_path_arg(path)) if path else _recent_log_path()
        if p is None:
            return "❌ 没有最近 session 日志；请先运行 stata_run_dofile，或显式传入 log 路径"
        if not p.exists():
            return f"❌ log 文件不存在：{p}"

        content = p.read_text(encoding="utf-8", errors="ignore")

        if tail_lines:
            lines = content.splitlines()
            content = "\n".join(lines[-int(tail_lines):])

        output_format = output_format or "full"
        if output_format == "core":
            content = _core_log_text(content)
        elif output_format == "dict":
            return json.dumps(_parse_log_to_command_results(content), ensure_ascii=False, indent=2)

        return f"=== {p.name} ===\n{content}"
    except Exception as e:
        return f"❌ 读取 log 失败：{str(e)}"


def _recent_log_path():
    if not LAST_SESSION_KEY:
        return None
    meta = _session_metadata(LAST_SESSION_KEY)
    last_log_path = meta.get("last_log_path")
    if not last_log_path:
        return None
    return Path(last_log_path)


def _core_log_text(content: str) -> str:
    kept = []
    skip_prefixes = (
        "log using ",
        "log close",
        "capture log close",
        "--------------------------------------------------------------------------------",
    )
    for line in content.splitlines():
        stripped = line.strip()
        lowered = stripped.lower()
        command_text = lowered[2:] if lowered.startswith(". ") else lowered
        if not stripped:
            continue
        if any(command_text.startswith(prefix) or lowered.startswith(prefix) for prefix in skip_prefixes):
            continue
        if stripped.startswith("name:") or stripped.startswith("log:"):
            continue
        if stripped.startswith("opened on:") or stripped.startswith("closed on:"):
            continue
        kept.append(line)
    return "\n".join(kept)


def _parse_log_to_command_results(content: str) -> list[dict]:
    parsed = []
    current = None
    def finish_current():
        if not current or not current["command"]:
            return
        current["result"] = "\n".join(current.pop("_result_lines")).strip()
        current["return_code"] = _extract_return_code(current["result"])
        parsed.append(current)

    for line in _core_log_text(content).splitlines():
        if line.startswith(". "):
            finish_current()
            command = line[2:].strip()
            current = {"command": command, "_result_lines": []} if command else None
        elif current:
            current["_result_lines"].append(line)
    finish_current()
    return parsed


def _extract_return_code(text: str):
    match = re.search(r"r\((\d+)\);", text or "")
    return match.group(1) if match else None


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

    log_path = _schema_log_path()
    if log_path is None:
        return "❌ 没有最近 session 的项目目录；请先运行 stata_run_dofile，或使用 Stata GUI 手动 describe"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    stata_log_path = _stata_path(log_path)
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


def _schema_log_path():
    meta = _session_metadata(LAST_SESSION_KEY) if LAST_SESSION_KEY else {}
    current_do = meta.get("current_do") or meta.get("entry_do") or meta.get("source_do")
    if not current_do:
        return None
    current_path = Path(current_do)
    return _project_mcp_dir(current_path) / CACHE_DIR_NAME / f"schema_{_safe_log_stem(LAST_SESSION_KEY)}.log"


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
