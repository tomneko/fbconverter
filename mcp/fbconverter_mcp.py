# /// script
# requires-python = ">=3.11"
# dependencies = ["mcp==2.3.0", "anyio>=4,<5", "pydantic>=2,<3"]
# ///
"""Local stdio MCP server for the fbconverter CLI."""

import json
import os
import subprocess
from pathlib import Path
from typing import Annotated, Literal

from anyio import run_process
from mcp.server import MCPServer
from mcp.types import CallToolResult, TextContent, ToolAnnotations
from pydantic import Field


server = MCPServer("fbconverter")

SourceVersion = Literal["1.0", "1.5", "2.0", "2.1", "2.5", "3.0", "4.0", "5.0"]
DestinationVersion = Literal["1.5", "2.0", "2.1", "2.5", "3.0", "4.0", "5.0"]


def conversion_result(destination: Path, exit_code: int | None, stdout: str, stderr: str) -> CallToolResult:
    log_dir = next(
        (line.removeprefix("Logs: ") for line in stdout.splitlines() if line.startswith("Logs: ")),
        None,
    )
    data = {
        "destination": str(destination),
        "exit_code": exit_code,
        "log_dir": log_dir,
        "stdout": stdout,
        "stderr": stderr,
    }
    return CallToolResult(
        content=[TextContent(type="text", text=json.dumps(data, ensure_ascii=False))],
        structured_content=data,
        is_error=exit_code != 0,
    )


@server.tool(annotations=ToolAnnotations(read_only_hint=False, destructive_hint=False, open_world_hint=True))
async def convert_database(
    source: Annotated[str, Field(min_length=1, description="Database path as seen by the source Firebird server.")],
    destination: Annotated[str, Field(min_length=1, description="Path of the new local database. Must not exist.")],
    source_version: SourceVersion,
    destination_version: DestinationVersion,
    source_host: str | None = None,
    source_port: Annotated[int, Field(ge=1, le=65535)] | None = None,
    source_user: str | None = None,
    source_charset: str | None = None,
    destination_charset: str | None = None,
    work_dir: str | None = None,
    single_thread: bool = False,
    source_local: bool = False,
) -> CallToolResult:
    """Copy a Firebird database to a new local Embedded database and return CLI output.

    The source host defaults to localhost. For a direct local connection, set
    source_local instead of supplying a host or port. Credentials come from the
    server process environment via FBCONVERTER_SOURCE_PASSWORD.
    """
    output_path = Path(destination).absolute()
    if source_local and (source_host is not None or source_port is not None):
        return conversion_result(output_path, None, "", "source_local cannot be combined with source_host or source_port.")

    configured_exe = os.environ.get("FBCONVERTER_EXE")
    exe = Path(configured_exe).absolute() if configured_exe else Path(__file__).with_name("fbconverter.exe")
    if not exe.is_file():
        return conversion_result(output_path, None, "", f"fbconverter.exe not found at {exe}. Set FBCONVERTER_EXE.")

    args = [
        str(exe),
        "--convert",
        "--source", str(Path(source).absolute()) if source_local else source,
        "--destination", str(output_path),
        "--source-version", source_version,
        "--destination-version", destination_version,
    ]
    if source_local:
        args.append("--source-local")
    else:
        args.extend(("--source-host", "localhost" if source_host is None else source_host))
    for flag, value in (
        ("--source-port", source_port),
        ("--source-user", source_user),
        ("--source-charset", source_charset),
        ("--destination-charset", destination_charset),
        ("--work-dir", Path(work_dir).absolute() if work_dir else work_dir),
    ):
        if value is not None:
            args.extend((flag, str(value)))
    if single_thread:
        args.append("--single-thread")

    try:
        process = await run_process(
            args,
            cwd=exe.parent,
            stdin=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            check=False,
        )
    except (OSError, ValueError) as exc:
        return conversion_result(output_path, None, "", f"Could not launch fbconverter: {exc}")

    return conversion_result(
        output_path,
        process.returncode,
        process.stdout.decode("utf-8", errors="replace").strip(),
        process.stderr.decode("utf-8", errors="replace").strip(),
    )


if __name__ == "__main__":
    server.run(transport="stdio")
