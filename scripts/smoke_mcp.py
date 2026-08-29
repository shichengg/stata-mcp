import argparse
import asyncio
import json
import sys
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def configure_stdout(stream) -> None:
    if hasattr(stream, "reconfigure"):
        stream.reconfigure(encoding="utf-8", errors="backslashreplace")


EXPECTED_TOOLS = {
    "stata_run",
    "stata_run_dofile",
    "stata_session",
    "stata_write_dofile",
    "stata_append_dofile",
    "stata_read_dofile",
    "stata_read_log",
    "stata_install_package",
    "stata_get_results",
    "stata_get_data_info",
    "stata_get_data_schema",
    "stata_status",
}


def result_text(result) -> str:
    return "\n".join(
        block.text
        for block in result.content
        if hasattr(block, "text")
    )


async def run(command: str, smoke_do: Path, exercise_stata: bool) -> None:
    parameters = StdioServerParameters(command=command, args=[])
    async with stdio_client(parameters) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            listed = await session.list_tools()
            names = {tool.name for tool in listed.tools}
            if names != EXPECTED_TOOLS:
                raise RuntimeError(f"Unexpected tools: {sorted(names)}")
            print(json.dumps({"tools": sorted(names)}, ensure_ascii=False))

            if not exercise_stata:
                return

            do_result = await session.call_tool(
                "stata_run_dofile",
                {
                    "path": str(smoke_do.resolve()),
                    "session_id": "release_smoke",
                    "role": "entry",
                    "log_mode": "replace",
                },
            )
            do_text = result_text(do_result)
            if "STATA_GUI_MCP_SMOKE_OK" not in do_text or "return_code: 0" not in do_text:
                raise RuntimeError(do_text)
            print(do_text)

            selection_result = await session.call_tool(
                "stata_run",
                {
                    "commands": (
                        "ereturn list\n"
                        "predict double price_hat\n"
                        "summarize price_hat"
                    )
                },
            )
            selection_text = result_text(selection_result)
            if "price_hat" not in selection_text or "return_code: 0" not in selection_text:
                raise RuntimeError(selection_text)
            print(selection_text)

            schema_result = await session.call_tool(
                "stata_get_data_schema",
                {
                    "sample_rows": 5,
                    "include_codebook": True,
                    "include_sample": True,
                    "include_missing": True,
                },
            )
            schema_text = result_text(schema_result)
            if "price" not in schema_text or "mpg" not in schema_text:
                raise RuntimeError(schema_text)
            print(schema_text)


if __name__ == "__main__":
    configure_stdout(sys.stdout)
    parser = argparse.ArgumentParser()
    parser.add_argument("--command", required=True)
    parser.add_argument(
        "--smoke-do",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "tests" / "manual_smoke.do",
    )
    parser.add_argument("--exercise-stata", action="store_true")
    arguments = parser.parse_args()
    asyncio.run(run(arguments.command, arguments.smoke_do, arguments.exercise_stata))
