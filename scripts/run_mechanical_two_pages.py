"""One-click two-page end-to-end preview for the mechanical principles book."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "scripts" / "quark_pdf_to_word.py"
SOURCE = ROOT / "input" / "word" / "机械原理习题集.docx"
REFERENCE = ROOT / "input" / "pdfs" / "机械原理习题集.pdf"
OUTPUT = ROOT / "output" / "word" / "机械原理习题集_前两页.docx"


def main() -> int:
    env = os.environ.copy()
    env["CODEX_LO_BRIDGE"] = "1"
    print("请先在普通 Windows PowerShell 中运行 scripts\\start_lo_bridge.cmd，并保持窗口打开。")
    print("本入口不再自动启动桥接器，避免 Codex 隔离账户触发 0xC0000022。")
    command = [
        sys.executable,
        str(MAIN),
        "--input", str(SOURCE),
        "--output", str(OUTPUT),
        "--reference", str(REFERENCE),
        "--max-pages", "2",
        "--compare", "full",
        "--bridge",
    ]
    return subprocess.call(command, cwd=str(ROOT), env=env)


if __name__ == "__main__":
    raise SystemExit(main())
