"""Run a command in a child process with a timeout so a hang or crash cannot stop the benchmark."""

import os
import signal
import subprocess
import time
from typing import Any, Dict, List, Optional

TAIL = 2000


def kill_tree(process: "subprocess.Popen[str]") -> None:
    """Kill a child process and everything it started."""
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(process.pid)], capture_output=True, check=False
        )
    else:
        os.killpg(process.pid, signal.SIGKILL)
    process.kill()


def run_isolated(
    argv: List[str], timeout: float, cwd: Optional[str] = None, env: Optional[Dict[str, str]] = None
) -> Dict[str, Any]:
    """Run argv; return status ok, crash (non-zero exit) or timeout, with the output."""
    started = time.perf_counter()
    process = subprocess.Popen(
        argv,
        cwd=cwd,
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        stdin=subprocess.DEVNULL,
        universal_newlines=True,
        encoding="utf-8",
        errors="replace",
        start_new_session=os.name != "nt",
    )
    try:
        stdout, stderr = process.communicate(timeout=timeout)
        status = "ok" if process.returncode == 0 else "crash"
    except subprocess.TimeoutExpired:
        kill_tree(process)
        stdout, stderr = process.communicate()
        status = "timeout"
    return {
        "status": status,
        "returncode": process.returncode,
        "seconds": time.perf_counter() - started,
        "stdout": stdout or "",
        "stderr": (stderr or "")[-TAIL:],
    }
