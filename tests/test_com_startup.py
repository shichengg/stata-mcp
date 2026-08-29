import pytest

from stata_mcp import server


def test_no_executable_override_skips_process_launch(monkeypatch):
    monkeypatch.setattr(server, "STATA_EXE", None)

    def unexpected_launch(*args, **kwargs):
        raise AssertionError("process launch must not run without STATA_EXE")

    monkeypatch.setattr(server.subprocess, "Popen", unexpected_launch)
    assert server._start_configured_stata_process() is None


def test_executable_override_must_exist(monkeypatch, tmp_path):
    missing = tmp_path / "StataMP-64.exe"
    monkeypatch.setattr(server, "STATA_EXE", missing)
    monkeypatch.setattr(server, "STATA_PROCESS", None)

    with pytest.raises(FileNotFoundError, match="Stata executable not found"):
        server._start_configured_stata_process()


def test_executable_override_restarts_exact_path_after_process_exits(monkeypatch, tmp_path):
    executable = tmp_path / "StataMP-64.exe"
    executable.write_bytes(b"")
    launched = []

    class FakeProcess:
        def __init__(self):
            self.returncode = None

        def poll(self):
            return self.returncode

    def fake_popen(args, close_fds):
        process = FakeProcess()
        launched.append((args, close_fds, process))
        return process

    monkeypatch.setattr(server, "STATA_EXE", executable)
    monkeypatch.setattr(server, "STATA_PROCESS", None)
    monkeypatch.setattr(server.subprocess, "Popen", fake_popen)
    monkeypatch.setattr(server.time, "sleep", lambda seconds: None)

    assert server._start_configured_stata_process() is None
    assert server._start_configured_stata_process() is None
    assert len(launched) == 1
    assert launched[0][:2] == ([str(executable)], True)

    launched[0][2].returncode = 0
    assert server._start_configured_stata_process() is None
    assert len(launched) == 2
    assert launched[1][:2] == ([str(executable)], True)


def test_status_explains_com_default(monkeypatch):
    monkeypatch.setattr(server, "STATA_EXE", None)
    text = server._status_text()
    assert "由 COM 注册信息定位" in text
