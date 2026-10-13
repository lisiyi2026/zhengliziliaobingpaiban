"""Run LibreOffice conversions under the interactive Windows user account.

Start this script from a normal PowerShell owned by the user. The project
queues jobs in tmp/lo-bridge/queue; this worker writes results to done/errors.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BRIDGE = ROOT / "tmp" / "lo-bridge"
# v2 isolates current jobs from stale bridge processes started under the
# Codex sandbox account.  The interactive user bridge is the only worker that
# watches this queue.
QUEUE = BRIDGE / "queue-v2"
DONE = BRIDGE / "done-v2"
ERRORS = BRIDGE / "errors-v2"


def find_soffice() -> str:
    found = shutil.which("soffice.com") or shutil.which("soffice.exe")
    if found:
        return found
    for name in ("soffice.com", "soffice.exe"):
        candidate = Path("C:/Program Files/LibreOffice/program") / name
        if candidate.exists():
            return str(candidate)
    raise RuntimeError("找不到 LibreOffice，请确认安装在 C:\\Program Files\\LibreOffice")


def clean_env() -> dict[str, str]:
    env = os.environ.copy()
    for name in ("PYTHONHOME", "PYTHONPATH", "PYTHONSTARTUP", "VIRTUAL_ENV"):
        env.pop(name, None)
    return env


def kill_tree(pid: int) -> None:
    subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True, check=False)


def process_job(job_path: Path) -> None:
    job = json.loads(job_path.read_text(encoding="utf-8"))
    job_id = str(job["id"])
    input_path = Path(job["input"]).resolve()
    output_dir = Path(job["output_dir"]).resolve()
    output_format = str(job.get("format", "pdf")).lower()
    if output_format not in {"pdf", "docx"}:
        raise RuntimeError(f"不支持的输出格式: {output_format}")
    if not input_path.exists():
        raise RuntimeError(f"输入文件不存在: {input_path}")
    output_dir.mkdir(parents=True, exist_ok=True)
    profile = BRIDGE / "profiles" / job_id
    profile.mkdir(parents=True, exist_ok=True)
    output = output_dir / f"{input_path.stem}.{output_format}"
    # LibreOffice has no PDF -> DOCX export filter. Use the installed
    # pdf2docx converter for that direction; keep LibreOffice for Office/PDF.
    if input_path.suffix.lower() == ".pdf" and output_format == "docx":
        from pdf2docx import Converter  # type: ignore
        converter = Converter(str(input_path))
        try:
            converter.convert(str(output), start=0, end=None)
        finally:
            converter.close()
        if not output.exists():
            raise RuntimeError(f"pdf2docx 未生成输出文件: {output}")
        (DONE / f"{job_id}.json").write_text(json.dumps({"output": str(output)}, ensure_ascii=False), encoding="utf-8")
        return
    command = [
        find_soffice(),
        "--headless", "--invisible", "--norestore", "--nofirststartwizard", "--nodefault",
        f"-env:UserInstallation={profile.as_uri()}",
        "--convert-to", output_format,
        "--outdir", str(output_dir),
        str(input_path),
    ]
    proc = subprocess.Popen(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=clean_env(),
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    try:
        stdout, stderr = proc.communicate(timeout=600)
    except subprocess.TimeoutExpired:
        kill_tree(proc.pid)
        raise RuntimeError("LibreOffice 转换超过 600 秒，已结束进程")
    if proc.returncode != 0 or not output.exists():
        detail = (stderr or stdout or f"退出码 {proc.returncode}").strip()
        raise RuntimeError(f"退出码 {proc.returncode}; 输出文件: {output}; LibreOffice: {detail}")
    (DONE / f"{job_id}.json").write_text(json.dumps({"output": str(output)}, ensure_ascii=False), encoding="utf-8")


def main() -> int:
    if os.environ.get("USERNAME", "").lower().startswith("codexsandbox"):
        raise RuntimeError("桥接器不能从 Codex 隔离账户启动；请在普通 Windows PowerShell 中双击 scripts\\start_lo_bridge.cmd")
    for folder in (QUEUE, DONE, ERRORS, BRIDGE / "profiles"):
        folder.mkdir(parents=True, exist_ok=True)
    print(f"LibreOffice 用户账户桥接器已启动: {ROOT}")
    print("保持此窗口运行；按 Ctrl+C 停止。")
    while True:
        jobs = sorted(QUEUE.glob("*.json"))
        if not jobs:
            time.sleep(1)
            continue
        for job_path in jobs:
            try:
                process_job(job_path)
            except Exception as exc:
                job_id = job_path.stem
                (ERRORS / f"{job_id}.json").write_text(str(exc), encoding="utf-8")
            finally:
                job_path.unlink(missing_ok=True)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n桥接器已停止。")
