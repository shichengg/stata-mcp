from pathlib import Path

from stata_mcp import server


def register_session_cache(tmp_path: Path, session_id: str = "main") -> Path:
    registry = tmp_path / ".stata-mcp" / "cache" / "task_registry.json"
    registry.parent.mkdir(parents=True)
    registry.write_text('{"sessions": {}}', encoding="utf-8")
    server.SESSION_REGISTRIES[session_id] = str(registry)
    return registry.parent


def test_run_and_schema_logs_are_separate_and_stable(tmp_path):
    cache = register_session_cache(tmp_path)
    server.LAST_SESSION_KEY = "main"

    first = server._run_log_path("main")
    second = server._run_log_path("main")
    schema = server._schema_log_path()

    assert first == second
    assert first.parent == cache
    assert first.name.startswith("run_main_")
    assert schema.parent == cache
    assert schema.name.startswith("schema_main_")
    assert schema != first


def test_schema_log_path_uses_explicit_session_not_recent_global(tmp_path):
    cache_a = register_session_cache(tmp_path / "project-a", "session-a")
    cache_b = register_session_cache(tmp_path / "project-b", "session-b")
    server.LAST_SESSION_KEY = "session-b"

    schema_a = server._schema_log_path("session-a")

    assert schema_a.parent == cache_a
    assert schema_a.parent != cache_b
    assert schema_a.name.startswith("schema_session-a_")


def test_initial_run_path_uses_dofile_project(tmp_path):
    do_path = tmp_path / "analysis.do"
    do_path.write_text("display 1", encoding="utf-8")

    result = server._run_log_path("new-session", do_path)

    assert result.parent == tmp_path / ".stata-mcp" / "cache"
    assert result.name.startswith("run_new-session_")


def test_extract_last_return_code_uses_last_error():
    assert server._extract_last_return_code("first r(111);\nlast r(198);") == 198
    assert server._extract_last_return_code("no Stata error") == 0


def test_truncate_output_keeps_tail(monkeypatch):
    monkeypatch.setenv("STATA_MCP_MAX_OUTPUT_CHARS", "20")
    text, truncated = server._truncate_output("0123456789abcdefghijTAIL")
    assert truncated is True
    assert text == "456789abcdefghijTAIL"


def test_format_success_includes_log_and_output(tmp_path):
    log_path = tmp_path / "run.log"
    log_path.write_text('. display 12345\n12345\n', encoding="utf-8")

    result = server._format_run_result("main", "selection", log_path)

    assert "✅ Stata 执行完成" in result
    assert "session_id: main" in result
    assert "source: selection" in result
    assert "return_code: 0" in result
    assert str(log_path) in result
    assert "12345" in result


def test_format_error_includes_return_code(tmp_path):
    log_path = tmp_path / "run.log"
    log_path.write_text("variable y not found\nr(111);\n", encoding="utf-8")

    result = server._format_run_result(
        "main",
        "dofile",
        log_path,
        tmp_path / "analysis.do",
    )

    assert "❌ Stata 执行失败" in result
    assert "return_code: 111" in result
    assert "source_do:" in result
