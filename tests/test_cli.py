import inspect

import stata_mcp
from stata_mcp import server


def test_release_version_is_1_1_0():
    assert stata_mcp.__version__ == "1.1.0"


def test_main_is_sync_and_serve_is_async():
    assert not inspect.iscoroutinefunction(server.main)
    assert inspect.iscoroutinefunction(server.serve)


def test_main_runs_serve(monkeypatch):
    calls = []

    async def fake_serve():
        return None

    def fake_run(coroutine):
        calls.append(coroutine)
        coroutine.close()

    monkeypatch.setattr(server, "serve", fake_serve)
    monkeypatch.setattr(server.asyncio, "run", fake_run)

    assert server.main() is None
    assert len(calls) == 1
