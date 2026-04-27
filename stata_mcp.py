# -*- coding: utf-8 -*-
"""Compatibility launcher for the Stata MCP server."""

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from stata_mcp.server import main


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
