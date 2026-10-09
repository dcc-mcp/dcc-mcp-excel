"""Launch the released DCC-MCP server for the Excel skill packs.

The Rust ``dcc-mcp-server`` owns HTTP transport, FileRegistry registration,
heartbeats, and shutdown. This module only supplies the Excel-specific
application identity, bundled skill path, and Python script runner — the
same contract `dcc_mcp_powerpoint.server_launcher` implements.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

from . import __version__

SERVER_ENV = "DCC_MCP_SERVER"
PYTHON_EXECUTABLE_ENV = "DCC_MCP_PYTHON_EXECUTABLE"
SERVER_NAMES = ("dcc-mcp-server.exe", "dcc-mcp-server")


def find_server_binary(explicit: str | None = None) -> str | None:
    """Resolve the official server without downloading or guessing a path."""
    candidates = [explicit, os.environ.get(SERVER_ENV)]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(Path(candidate).resolve())
    for name in SERVER_NAMES:
        resolved = shutil.which(name)
        if resolved:
            return resolved
    return None


def bundled_skills_dir() -> Path:
    """Return the filesystem-backed skill directory shipped with the adapter."""
    return Path(__file__).resolve().parent / "skills"


@dataclass(frozen=True)
class ServeConfig:
    """Inputs for one registered Excel server process."""

    server: str | None = None
    mcp_port: int = 0
    registry_dir: str | None = None
    pid_file: str | None = None

    def command(self) -> list[str]:
        server = find_server_binary(self.server)
        if server is None:
            raise FileNotFoundError(
                "dcc-mcp-server was not found; install the official DCC-MCP "
                f"runtime or set {SERVER_ENV} to its absolute path"
            )
        skills = bundled_skills_dir()
        if not skills.is_dir():
            raise FileNotFoundError(f"bundled Excel skills not found: {skills}")
        command = [
            server,
            "serve",
            "--app",
            "excel",
            "--skill-paths",
            str(skills),
            "--server-name",
            "dcc-mcp-excel",
            "--app-version",
            __version__,
            "--mcp-port",
            str(self.mcp_port),
            "--no-bridge",
        ]
        if self.registry_dir:
            command.extend(("--registry-dir", str(Path(self.registry_dir).resolve())))
        if self.pid_file:
            command.extend(("--pid-file", str(Path(self.pid_file).resolve())))
        return command


def serve(config: ServeConfig) -> int:
    """Run until the registered server exits, forwarding its exit status."""
    environment = os.environ.copy()
    environment[PYTHON_EXECUTABLE_ENV] = sys.executable
    return subprocess.call(config.command(), env=environment)
