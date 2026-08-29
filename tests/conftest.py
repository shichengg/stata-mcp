import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


@pytest.fixture(autouse=True)
def reset_server_state():
    from stata_mcp import server

    server.SESSIONS.clear()
    server.SESSION_REGISTRIES.clear()
    server.KNOWN_REGISTRY_PATHS.clear()
    server.LAST_SESSION_KEY = None
    server.STATA_ERROR = ""
    yield
    server.SESSIONS.clear()
    server.SESSION_REGISTRIES.clear()
    server.KNOWN_REGISTRY_PATHS.clear()
    server.LAST_SESSION_KEY = None
    server.STATA_ERROR = ""
