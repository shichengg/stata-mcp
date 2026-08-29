import threading
from types import SimpleNamespace

from stata_mcp import server


class ReadyFakeSession:
    def __init__(self, key, close_result=True):
        self.key = key
        self.ready = threading.Event()
        self.ready.set()
        self.app = object()
        self.error = ""
        self.closed = False
        self.close_result = close_result
        self.transaction_lock = threading.RLock()

    def is_ready(self):
        return not self.closed

    def close(self, timeout=5):
        self.closed = True
        return self.close_result


def test_stata_session_close_stops_thread_and_uninitializes_com(monkeypatch):
    events = []

    class FakeApp:
        def DoCommand(self, command):
            events.append(("command", command))

    monkeypatch.setattr(
        server.pythoncom,
        "CoInitialize",
        lambda: events.append(("com", "initialize")),
    )
    monkeypatch.setattr(
        server.pythoncom,
        "CoUninitialize",
        lambda: events.append(("com", "uninitialize")),
    )
    monkeypatch.setattr(server, "_start_configured_stata_process", lambda: None)
    monkeypatch.setattr(
        server.win32com.client,
        "Dispatch",
        lambda prog_id: FakeApp(),
    )
    monkeypatch.setattr(server.time, "sleep", lambda seconds: None)

    session = server.StataSession("lifecycle")
    assert session.ready.wait(timeout=2)
    assert session.close(timeout=2) is True

    assert not session.thread.is_alive()
    assert session.app is None
    assert events[0] == ("com", "initialize")
    assert events[-1] == ("com", "uninitialize")


def test_destroy_closes_removed_session():
    session = ReadyFakeSession("main")
    server.SESSIONS["main"] = session
    server.LAST_SESSION_KEY = "main"

    result = server._session_tool("destroy", "main")

    assert session.closed is True
    assert "main" not in server.SESSIONS
    assert server.LAST_SESSION_KEY is None
    assert "已从 MCP 内存移除" in result


def test_force_new_closes_old_session(monkeypatch):
    old_session = ReadyFakeSession("main")
    replacement = ReadyFakeSession("main")
    server.SESSIONS["main"] = old_session
    monkeypatch.setattr(server, "StataSession", lambda key: replacement)

    result = server._get_session("main", force_new=True)

    assert old_session.closed is True
    assert result is replacement
    assert server.SESSIONS["main"] is replacement


def test_dispose_does_not_wait_for_busy_transaction_lock():
    session = ReadyFakeSession("main", close_result=False)
    session.transaction_lock = threading.Lock()
    session.transaction_lock.acquire()
    server.SESSIONS["main"] = session
    finished = threading.Event()

    def dispose():
        server._dispose_session("main", timeout=0.01)
        finished.set()

    worker = threading.Thread(target=dispose)
    worker.start()
    completed_while_busy = finished.wait(timeout=0.1)
    session.transaction_lock.release()
    worker.join(timeout=2)

    assert completed_while_busy is True
    assert server.SESSIONS["main"] is session


def test_force_new_does_not_replace_worker_that_failed_to_close(monkeypatch):
    old_session = ReadyFakeSession("main", close_result=False)
    server.SESSIONS["main"] = old_session
    replacements = []
    monkeypatch.setattr(
        server,
        "StataSession",
        lambda key: replacements.append(ReadyFakeSession(key)) or replacements[-1],
    )

    result = server._get_session("main", force_new=True)

    assert result is None
    assert replacements == []
    assert server.SESSIONS["main"] is old_session
    assert "无法安全停止" in server.STATA_ERROR


def test_destroy_reports_busy_worker_instead_of_false_success():
    session = ReadyFakeSession("main", close_result=False)
    server.SESSIONS["main"] = session

    result = server._session_tool("destroy", "main")

    assert server.SESSIONS["main"] is session
    assert "仍在结束当前命令" in result
    assert "已从 MCP 内存移除" not in result
