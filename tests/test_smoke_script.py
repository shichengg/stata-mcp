import io
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import smoke_mcp


def test_configure_stdout_uses_utf8_for_unicode_status_text():
    raw = io.BytesIO()
    stream = io.TextIOWrapper(raw, encoding="gbk")

    smoke_mcp.configure_stdout(stream)
    stream.write("✅ Stata 执行完成")
    stream.flush()

    assert raw.getvalue().decode("utf-8") == "✅ Stata 执行完成"
