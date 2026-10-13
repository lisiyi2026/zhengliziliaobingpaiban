"""Run MinerU under the interactive Windows user account.

The main pipeline can run inside the Codex sandbox, but Torch DLL loading is
restricted there on some Windows installations. This worker consumes jobs from
tmp/mineru-bridge/queue-v1 and executes the project's .venv-mineru interpreter
under the normal user account.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BRIDGE = ROOT / "tmp" / "mineru-bridge"
QUEUE = BRIDGE / "queue-v1"
DONE = BRIDGE / "done-v1"
ERRORS = BRIDGE / "errors-v1"


def clean_env() -> dict[str, str]:
    env = os.environ.copy()
    for name in ("PYTHONHOME", "PYTHONPATH", "PYTHONSTARTUP", "VIRTUAL_ENV"):
        env.pop(name, None)
    return env


def find_content_list(output_dir: Path, stem: str) -> Path | None:
    for name in (f"{stem}_content_list_v2.json", f"{stem}_content_list.json"):
        matches = sorted(output_dir.rglob(name))
        if matches:
            return matches[0]
    return None


def process_job(job_path: Path) -> None:
    job = json.loads(job_path.read_text(encoding="utf-8"))
    job_id = str(job["id"])
    input_path = Path(job["input"]).resolve()
    output_dir = Path(job["output_dir"]).resolve()
    python = Path(job["python"]).resolve()
    if not input_path.exists():
        raise RuntimeError(f"输入文件不存在: {input_path}")
    if not python.exists():
        raise RuntimeError(f"MinerU Python 不存在: {python}")
    output_dir.mkdir(parents=True, exist_ok=True)
    command = [
        str(python), "-m", "mineru.cli.client",
        "-p", str(input_path), "-o", str(output_dir),
        "-b", str(job.get("backend", "pipeline")),
        "-m", str(job.get("method", "auto")),
        "-l", "ch", "--client-side-output-generation=True",
    ]
    timeout = max(60, int(job.get("timeout", 1800)))
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            timeout=timeout,
            env=clean_env(),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(f"MinerU 转换超时（{timeout} 秒）: {exc}") from exc
    content = find_content_list(output_dir, input_path.stem)
    if result.returncode != 0 or content is None:
        detail = (result.stderr or result.stdout or f"退出码 {result.returncode}").strip()
        raise RuntimeError(f"MinerU 退出码 {result.returncode}: {detail[-6000:]}")
    (DONE / f"{job_id}.json").write_text(
        json.dumps(
            {
                "output_dir": str(output_dir),
                "content_list": str(content),
                "stdout": result.stdout[-2000:],
                "stderr": result.stderr[-2000:],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def main() -> int:
    if os.environ.get("USERNAME", "").lower().startswith("codexsandbox"):
        raise RuntimeError("MinerU 桥接器必须在普通 Windows PowerShell 中启动")
    for folder in (QUEUE, DONE, ERRORS, BRIDGE / "work"):
        folder.mkdir(parents=True, exist_ok=True)
    print(f"MinerU 用户账户桥接器已启动: {ROOT}")
    print("保持此窗口运行；按 Ctrl+C 停止。")
    while True:
        jobs = sorted(QUEUE.glob("*.json"))
        if not jobs:
            time.sleep(1)
            continue
        for job_path in jobs:
            job_id = job_path.stem
            try:
                process_job(job_path)
            except Exception as exc:
                (ERRORS / f"{job_id}.json").write_text(str(exc), encoding="utf-8")
            finally:
                job_path.unlink(missing_ok=True)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nMinerU 桥接器已停止。")
