import json
import multiprocessing
import os
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from stata_mcp import server


def update_registry_in_process(registry_path: str, session_id: str) -> None:
    server._update_task_registry(
        session_id,
        Path(registry_path),
        status="ready",
    )


def test_registry_save_is_atomic_when_replace_fails(monkeypatch, tmp_path):
    registry_path = tmp_path / "task_registry.json"
    original = {"sessions": {"existing": {"status": "ready"}}}
    registry_path.write_text(json.dumps(original), encoding="utf-8")

    def fail_replace(source, destination):
        raise OSError("synthetic replace failure")

    monkeypatch.setattr(server.os, "replace", fail_replace)

    with pytest.raises(OSError, match="synthetic replace failure"):
        server._save_task_registry(
            registry_path,
            {"sessions": {"replacement": {"status": "ready"}}},
        )

    assert json.loads(registry_path.read_text(encoding="utf-8")) == original
    assert list(tmp_path.glob("*.tmp")) == []


def test_parallel_registry_updates_do_not_lose_sessions(tmp_path):
    registry_path = tmp_path / "cache" / "task_registry.json"
    session_ids = [f"session-{index}" for index in range(40)]

    def update(session_id):
        server._update_task_registry(
            session_id,
            registry_path,
            status="ready",
        )

    with ThreadPoolExecutor(max_workers=16) as executor:
        list(executor.map(update, session_ids))

    registry = json.loads(registry_path.read_text(encoding="utf-8"))
    assert set(registry["sessions"]) == set(session_ids)


@pytest.mark.skipif(os.name != "nt", reason="cross-process registry locking is Windows-only")
def test_cross_process_registry_updates_do_not_lose_sessions(tmp_path):
    registry_path = tmp_path / "cache" / "task_registry.json"
    session_ids = [f"process-session-{index}" for index in range(20)]
    context = multiprocessing.get_context("spawn")
    processes = [
        context.Process(
            target=update_registry_in_process,
            args=(str(registry_path), session_id),
        )
        for session_id in session_ids
    ]

    for process in processes:
        process.start()
    for process in processes:
        process.join(timeout=20)
        assert process.exitcode == 0

    registry = json.loads(registry_path.read_text(encoding="utf-8"))
    assert set(registry["sessions"]) == set(session_ids)
