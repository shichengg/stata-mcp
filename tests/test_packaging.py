import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_pyproject_release_contract():
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    project = data["project"]
    assert project["name"] == "stata-gui-mcp"
    assert project["version"] == "1.1.0"
    assert project["license"] == "MIT"
    assert project["scripts"]["stata-gui-mcp"] == "stata_mcp.server:main"
    assert "mcp>=1,<2" in project["dependencies"]
    assert any("pywin32" in dependency for dependency in project["dependencies"])


def test_readme_has_required_client_examples_and_no_removed_sections():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    required = [
        "pip install stata-gui-mcp",
        "Claude Code",
        "claude mcp add --transport stdio --scope user stata -- stata-gui-mcp",
        "Codex",
        "codex mcp add stata -- stata-gui-mcp",
        "OpenCode",
        "Claude Desktop",
        "Cursor",
        "VS Code",
        '"command": "stata-gui-mcp"',
        '"args": ["-m", "stata_mcp"]',
        "SepineTam/mcp-for-stata",
        "hanlulong/stata-mcp",
        "run_<session",
        "schema_<session",
    ]
    for text in required:
        assert text in readme
    assert "pip install --upgrade stata-gui-mcp" not in readme
    assert "stata-gui-mcp --version" not in readme


def test_mit_license_exists():
    license_text = (ROOT / "LICENSE").read_text(encoding="utf-8")
    assert "MIT License" in license_text
    assert "Permission is hereby granted" in license_text
    assert "shichengg" in license_text
