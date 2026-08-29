import asyncio
import json
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from types import SimpleNamespace

from stata_mcp import server


def set_session(tmp_path: Path, session_id: str = "main") -> Path:
    project = tmp_path / "project"
    project.mkdir()
    source = project / "analysis.do"
    source.write_text("display 1", encoding="utf-8")
    registry = project / ".stata-mcp" / "cache" / "task_registry.json"
    registry.parent.mkdir(parents=True)
    registry.write_text(
        json.dumps(
            {
                "sessions": {
                    session_id: {
                        "session_id": session_id,
                        "current_do": str(source),
                        "working_dir": str(project),
                    }
                }
            }
        ),
        encoding="utf-8",
    )
    server.SESSION_REGISTRIES[session_id] = str(registry)
    server.KNOWN_REGISTRY_PATHS.add(str(registry))
    server.LAST_SESSION_KEY = session_id
    return project


def test_selection_runs_as_one_temporary_dofile_and_is_removed(monkeypatch, tmp_path):
    project = set_session(tmp_path)
    fake_session = SimpleNamespace(
        key="main",
        transaction_lock=threading.RLock(),
        is_ready=lambda: True,
    )
    observed = {}

    monkeypatch.setattr(server, "PYWIN32_READY", True)
    monkeypatch.setattr(
        server,
        "_get_session",
        lambda key=None, force_new=False: fake_session,
    )

    def fake_captured_run(path, session_key, log_path, working_dir, log_mode="replace", session=None):
        observed["path"] = Path(path)
        observed["content"] = Path(path).read_text(encoding="utf-8")
        observed["session_key"] = session_key
        observed["working_dir"] = Path(working_dir)
        Path(log_path).parent.mkdir(parents=True, exist_ok=True)
        Path(log_path).write_text(
            ". foreach var in price mpg {\nsummary output\n",
            encoding="utf-8",
        )
        return ""

    monkeypatch.setattr(server, "_run_dofile_with_log", fake_captured_run)

    code = "foreach var in price mpg {\n    summarize `var'\n}"
    result = server._run_stata_commands(code)

    assert observed["content"] == code + "\n"
    assert observed["session_key"] == "main"
    assert observed["working_dir"] == project
    assert observed["path"].suffix == ".do"
    assert not observed["path"].exists()
    assert "Stata 执行完成" in result
    assert "summary output" in result


def test_empty_selection_is_rejected(monkeypatch, tmp_path):
    set_session(tmp_path)
    monkeypatch.setattr(server, "PYWIN32_READY", True)
    monkeypatch.setattr(
        server,
        "_get_session",
        lambda key=None, force_new=False: SimpleNamespace(
        key="main",
        transaction_lock=threading.RLock(),
        is_ready=lambda: True,
    ),
    )
    assert server._run_stata_commands("  \n") == "❌ 没有可执行的 Stata 命令"


def test_temp_selection_is_removed_when_execution_fails(monkeypatch, tmp_path):
    set_session(tmp_path)
    fake_session = SimpleNamespace(
        key="main",
        transaction_lock=threading.RLock(),
        is_ready=lambda: True,
    )
    observed = {}

    monkeypatch.setattr(server, "PYWIN32_READY", True)
    monkeypatch.setattr(
        server,
        "_get_session",
        lambda key=None, force_new=False: fake_session,
    )

    def fake_failure(path, session_key, log_path, working_dir, log_mode="replace", session=None):
        observed["path"] = Path(path)
        return "❌ do 文件发送失败：synthetic failure"

    monkeypatch.setattr(server, "_run_dofile_with_log", fake_failure)

    result = server._run_stata_commands("display 1")

    assert result == "❌ do 文件发送失败：synthetic failure"
    assert not observed["path"].exists()


def test_full_dofile_returns_latest_log(monkeypatch, tmp_path):
    source = tmp_path / "analysis.do"
    source.write_text("display 24680", encoding="utf-8")
    monkeypatch.setattr(server, "PYWIN32_READY", True)
    monkeypatch.setattr(
        server,
        "_get_session",
        lambda key=None, force_new=False: SimpleNamespace(
            key=key,
            transaction_lock=threading.RLock(),
            is_ready=lambda: True,
        ),
    )

    def fake_captured_run(path, session_key, log_path, working_dir, log_mode="replace", session=None):
        assert Path(path) == source
        assert Path(working_dir) == tmp_path
        assert log_mode == "replace"
        Path(log_path).parent.mkdir(parents=True, exist_ok=True)
        Path(log_path).write_text(". display 24680\n24680\n", encoding="utf-8")
        return ""

    monkeypatch.setattr(server, "_run_dofile_with_log", fake_captured_run)

    result = server._run_dofile_session(
        str(source),
        "release",
        "entry",
        "replace",
    )

    assert "Stata 执行完成" in result
    assert "source_do:" in result
    assert "24680" in result
    assert "run_release_" in result


def test_same_session_keeps_original_registry_directory(tmp_path):
    first_dir = tmp_path / "first"
    second_dir = tmp_path / "second"
    first_dir.mkdir()
    second_dir.mkdir()
    first = first_dir / "entry.do"
    second = second_dir / "aux.do"
    first.write_text("display 1", encoding="utf-8")
    second.write_text("display 2", encoding="utf-8")
    first_log = server._run_log_path("main", first)
    first_log.parent.mkdir(parents=True, exist_ok=True)

    server._register_dofile_run(first, "main", "entry", first_log)
    second_log = server._run_log_path("main", second)
    server._register_dofile_run(second, "main", "auxiliary", second_log)

    assert second_log.parent == first_log.parent
    assert server.SESSION_REGISTRIES["main"] == str(
        first_log.parent / "task_registry.json"
    )


def test_disconnect_preserves_log_and_marks_registry(tmp_path):
    project = set_session(tmp_path)
    log_path = project / ".stata-mcp" / "cache" / "run.log"
    log_path.write_text("valuable output", encoding="utf-8")
    registry_path = Path(server.SESSION_REGISTRIES["main"])
    registry = server._load_task_registry(registry_path)
    registry["sessions"]["main"]["last_log_path"] = str(log_path)
    server._save_task_registry(registry_path, registry)
    server.SESSIONS["main"] = SimpleNamespace(
        transaction_lock=threading.RLock(),
        close=lambda timeout=5: True,
    )

    server._cleanup_disconnected_session("main", extra_log_paths=[log_path])

    updated = server._load_task_registry(registry_path)
    assert log_path.read_text(encoding="utf-8") == "valuable output"
    assert updated["sessions"]["main"]["status"] == "disconnected"
    assert "main" not in server.SESSIONS


def test_dofile_tool_rejects_directory_path(monkeypatch, tmp_path):
    called = []
    monkeypatch.setattr(
        server,
        "_run_dofile_session",
        lambda *args: called.append(args) or "unexpected",
    )

    result = asyncio.run(
        server.call_tool(
            "stata_run_dofile",
            {"path": str(tmp_path)},
        )
    )

    assert called == []
    assert "do 文件不存在或不是文件" in result[0].text


def test_distinct_raw_session_ids_do_not_collide(tmp_path):
    source = tmp_path / "analysis.do"
    source.write_text("display 1", encoding="utf-8")

    slash_id = server._session_key_for_dofile(source, "a/b")
    underscore_id = server._session_key_for_dofile(source, "a_b")

    assert slash_id == "a/b"
    assert underscore_id == "a_b"
    assert slash_id != underscore_id

    server.SESSIONS["a/b"] = SimpleNamespace(is_ready=lambda: True, error="")
    details = json.loads(server._session_tool("get", "a/b"))
    assert details["session_id"] == "a/b"
    assert details["state"] == "ready"


def test_parallel_selection_transactions_are_serialized(monkeypatch, tmp_path):
    set_session(tmp_path)
    transaction_lock = threading.RLock()
    fake_session = SimpleNamespace(
        key="main",
        transaction_lock=transaction_lock,
        is_ready=lambda: True,
    )
    state_lock = threading.Lock()
    overlap_barrier = threading.Barrier(2)
    state = {"active": 0, "max_active": 0}

    monkeypatch.setattr(server, "PYWIN32_READY", True)
    monkeypatch.setattr(
        server,
        "_get_session",
        lambda key=None, force_new=False: fake_session,
    )

    def fake_captured_run(path, session_key, log_path, working_dir, log_mode="replace", session=None):
        marker = Path(path).read_text(encoding="utf-8").strip()
        with state_lock:
            state["active"] += 1
            state["max_active"] = max(state["max_active"], state["active"])
        try:
            overlap_barrier.wait(timeout=0.2)
        except threading.BrokenBarrierError:
            pass
        Path(log_path).write_text(marker, encoding="utf-8")
        with state_lock:
            state["active"] -= 1
        return ""

    monkeypatch.setattr(server, "_run_dofile_with_log", fake_captured_run)

    commands = ["display 111", "display 222"]
    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(server._run_stata_commands, commands))

    assert state["max_active"] == 1
    assert commands[0] in results[0]
    assert commands[1] in results[1]


def test_parallel_dofile_transactions_are_serialized(monkeypatch, tmp_path):
    sources = []
    for index in (111, 222):
        source = tmp_path / f"analysis_{index}.do"
        source.write_text(f"display {index}", encoding="utf-8")
        sources.append(source)

    fake_session = SimpleNamespace(
        key="main",
        transaction_lock=threading.RLock(),
        is_ready=lambda: True,
    )
    state_lock = threading.Lock()
    overlap_barrier = threading.Barrier(2)
    state = {"active": 0, "max_active": 0}

    monkeypatch.setattr(server, "PYWIN32_READY", True)
    monkeypatch.setattr(
        server,
        "_get_session",
        lambda key=None, force_new=False: fake_session,
    )

    def fake_captured_run(path, session_key, log_path, working_dir, log_mode="replace", session=None):
        marker = Path(path).read_text(encoding="utf-8").strip()
        with state_lock:
            state["active"] += 1
            state["max_active"] = max(state["max_active"], state["active"])
        try:
            overlap_barrier.wait(timeout=0.2)
        except threading.BrokenBarrierError:
            pass
        Path(log_path).parent.mkdir(parents=True, exist_ok=True)
        Path(log_path).write_text(marker, encoding="utf-8")
        with state_lock:
            state["active"] -= 1
        return ""

    monkeypatch.setattr(server, "_run_dofile_with_log", fake_captured_run)

    def run_source(source):
        return server._run_dofile_session(
            str(source),
            "main",
            "entry",
            "replace",
        )

    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(run_source, sources))

    assert state["max_active"] == 1
    assert "display 111" in results[0]
    assert "display 222" in results[1]


def test_captured_run_quarantines_session_after_command_timeout(tmp_path):
    source = tmp_path / "analysis.do"
    source.write_text("display 1", encoding="utf-8")
    log_path = tmp_path / "run.log"

    class TimeoutSession:
        def __init__(self):
            self.commands = []
            self.close_calls = []

        def execute(self, command):
            self.commands.append(command)
            if command.startswith("do "):
                raise TimeoutError("synthetic command timeout")
            return True

        def close(self, timeout=5):
            self.close_calls.append(timeout)
            return False

    session = TimeoutSession()

    result = server._run_dofile_with_log(
        source,
        "main",
        log_path,
        tmp_path,
        "replace",
        session=session,
    )

    assert "执行超时" in result
    assert session.close_calls == [0]
    assert session.commands[-1].startswith("do ")
