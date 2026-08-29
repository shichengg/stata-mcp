import asyncio
import json
import threading
from pathlib import Path
from types import SimpleNamespace

from stata_mcp import server


def prepare_schema_session(tmp_path: Path):
    source = tmp_path / "analysis.do"
    source.write_text("display 1", encoding="utf-8")
    registry = tmp_path / ".stata-mcp" / "cache" / "task_registry.json"
    registry.parent.mkdir(parents=True)
    registry.write_text(
        json.dumps(
            {
                "sessions": {
                    "main": {
                        "session_id": "main",
                        "current_do": str(source),
                    }
                }
            }
        ),
        encoding="utf-8",
    )
    server.SESSION_REGISTRIES["main"] = str(registry)
    server.KNOWN_REGISTRY_PATHS.add(str(registry))
    server.LAST_SESSION_KEY = "main"


def test_schema_uses_named_log_without_closing_all(monkeypatch, tmp_path):
    prepare_schema_session(tmp_path)
    commands = []
    fake_session = SimpleNamespace(
        key="main",
        transaction_lock=threading.RLock(),
        is_ready=lambda: True,
    )

    monkeypatch.setattr(
        server,
        "_get_session",
        lambda key=None, force_new=False: fake_session,
    )

    def fake_execute(session_key, command, extra_log_paths=None, session=None):
        commands.append(command)
        if command == f"capture log close {server.MCP_SCHEMA_LOG_NAME}":
            schema_path = server._schema_log_path()
            schema_path.parent.mkdir(parents=True, exist_ok=True)
            schema_path.write_text(
                "=== Stata data schema snapshot ===\nprice float\n",
                encoding="utf-8",
            )
        return False

    monkeypatch.setattr(server, "_execute_with_reconnect", fake_execute)

    result = server._get_data_schema(20, True, True, True)

    assert "price float" in result
    assert all("log close _all" not in command for command in commands)
    assert any(
        f"name({server.MCP_SCHEMA_LOG_NAME})" in command
        for command in commands
    )
    assert "describe" in commands
    assert "codebook, compact" in commands
    assert "misstable summarize" in commands
    assert "list in 1/20, abbreviate(20)" in commands


def test_tool_list_keeps_all_twelve_names():
    tools = asyncio.run(server.list_tools())
    names = {tool.name for tool in tools}
    assert names == {
        "stata_run",
        "stata_run_dofile",
        "stata_session",
        "stata_write_dofile",
        "stata_append_dofile",
        "stata_read_dofile",
        "stata_read_log",
        "stata_install_package",
        "stata_get_results",
        "stata_get_data_info",
        "stata_get_data_schema",
        "stata_status",
    }


def test_result_helpers_do_not_tell_ai_to_open_another_log(monkeypatch):
    monkeypatch.setattr(
        server,
        "_run_stata_commands",
        lambda command: f"captured: {command}",
    )
    assert server._get_stored_results("e") == "captured: ereturn list"
    assert server._get_data_info() == "captured: describe"


def test_schema_waits_for_session_transaction_lock(monkeypatch, tmp_path):
    prepare_schema_session(tmp_path)
    transaction_lock = threading.RLock()
    fake_session = SimpleNamespace(
        key="main",
        transaction_lock=transaction_lock,
        is_ready=lambda: True,
        execute=lambda command: None,
    )
    command_started = threading.Event()

    monkeypatch.setattr(
        server,
        "_get_session",
        lambda key=None, force_new=False: fake_session,
    )

    def fake_execute(session_key, command, extra_log_paths=None, session=None):
        command_started.set()
        if command == f"capture log close {server.MCP_SCHEMA_LOG_NAME}":
            schema_path = server._schema_log_path()
            schema_path.parent.mkdir(parents=True, exist_ok=True)
            schema_path.write_text("price float\n", encoding="utf-8")
        return False

    monkeypatch.setattr(server, "_execute_with_reconnect", fake_execute)

    transaction_lock.acquire()
    worker = threading.Thread(
        target=server._get_data_schema,
        args=(5, True, True, True),
    )
    worker.start()
    started_while_locked = command_started.wait(timeout=0.1)
    transaction_lock.release()
    worker.join(timeout=2)

    assert started_while_locked is False
    assert not worker.is_alive()


def test_schema_timeout_quarantines_session_without_queuing_close(monkeypatch, tmp_path):
    prepare_schema_session(tmp_path)
    executed = []
    close_calls = []
    fake_session = SimpleNamespace(
        key="main",
        transaction_lock=threading.RLock(),
        is_ready=lambda: True,
        execute=lambda command: executed.append(command),
        close=lambda timeout=5: close_calls.append(timeout) or False,
    )

    monkeypatch.setattr(
        server,
        "_get_session",
        lambda key=None, force_new=False: fake_session,
    )

    def timeout_execute(session_key, command, extra_log_paths=None, session=None):
        if command == "describe":
            raise TimeoutError("synthetic schema timeout")
        return False

    monkeypatch.setattr(server, "_execute_with_reconnect", timeout_execute)

    result = server._get_data_schema(5, True, True, True)

    assert "快照执行超时" in result
    assert close_calls == [0]
    assert executed == []
