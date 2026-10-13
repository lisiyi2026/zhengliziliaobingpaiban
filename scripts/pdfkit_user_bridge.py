"""Run PDF-Extract-Kit formula detection in the interactive user account."""

from __future__ import annotations

import json
import os
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BRIDGE = ROOT / "tmp" / "pdfkit-bridge"
QUEUE = BRIDGE / "queue-v1"
DONE = BRIDGE / "done-v1"
ERRORS = BRIDGE / "errors-v1"
READY = BRIDGE / "ready-v1.json"
DEFAULT_MODEL = ROOT / "models" / "pdf-extract-kit" / "models" / "MFD" / "weights.pt"


def _boxes(result):
    boxes = getattr(result, "boxes", None)
    if boxes is None:
        return []
    xyxy = boxes.xyxy.detach().cpu().tolist() if getattr(boxes, "xyxy", None) is not None else []
    conf = boxes.conf.detach().cpu().tolist() if getattr(boxes, "conf", None) is not None else []
    cls = boxes.cls.detach().cpu().tolist() if getattr(boxes, "cls", None) is not None else []
    output = []
    for index, bbox in enumerate(xyxy):
        output.append({
            "bbox_pixels": [round(float(value), 2) for value in bbox],
            "confidence": round(float(conf[index]), 6) if index < len(conf) else None,
            "class_id": int(cls[index]) if index < len(cls) else None,
            "kind": "inline" if index < len(cls) and int(cls[index]) == 0 else "isolated",
        })
    return output


def process_job(job_path: Path) -> None:
    job = json.loads(job_path.read_text(encoding="utf-8"))
    job_id = str(job["id"])
    input_path = Path(job["input"]).resolve()
    output_path = Path(job["output"]).resolve()
    model_path = Path(job.get("model_path") or DEFAULT_MODEL).resolve()
    if not input_path.exists():
        raise RuntimeError(f"输入文件不存在: {input_path}")
    if not model_path.exists():
        raise RuntimeError(f"公式检测权重不存在: {model_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        import fitz  # type: ignore
        from pdf_extract_kit.tasks.formula_detection.models.yolo import FormulaDetectionYOLO  # type: ignore
    except Exception as exc:
        raise RuntimeError(f"PDF-Extract-Kit 导入失败: {exc}") from exc

    dpi = max(72, int(job.get("dpi", 144)))
    max_pages = job.get("max_pages")
    max_pages = int(max_pages) if max_pages is not None else None
    detector = FormulaDetectionYOLO({
        "model_path": str(model_path),
        "img_size": int(job.get("img_size", 1280)),
        "pdf_dpi": dpi,
        "conf_thres": float(job.get("conf_thres", 0.25)),
        "iou_thres": float(job.get("iou_thres", 0.45)),
        "device": "cpu",
        "visualize": False,
    })
    document = fitz.open(str(input_path))
    limit = min(len(document), max_pages) if max_pages else len(document)
    images = []
    page_meta = []
    for page_number in range(limit):
        page = document[page_number]
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0), alpha=False)
        from PIL import Image  # type: ignore
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(image)
        page_meta.append({
            "page": page_number + 1,
            "pdf_size": [round(float(page.rect.width), 2), round(float(page.rect.height), 2)],
            "pixel_size": [pix.width, pix.height],
        })
    results = detector.predict(images, str(BRIDGE / "work"), [f"page_{n:04d}" for n in range(1, limit + 1)]) if images else []
    candidates = []
    for meta, result in zip(page_meta, results):
        sx = meta["pdf_size"][0] / max(meta["pixel_size"][0], 1)
        sy = meta["pdf_size"][1] / max(meta["pixel_size"][1], 1)
        for item in _boxes(result):
            x0, y0, x1, y1 = item.pop("bbox_pixels")
            item.update({
                "page": meta["page"],
                "bbox": [round(x0 * sx, 2), round(y0 * sy, 2), round(x1 * sx, 2), round(y1 * sy, 2)],
                "detector": "pdf-extract-kit-yolov8",
            })
            candidates.append(item)
    document.close()
    output_path.write_text(json.dumps({
        "engine": "pdf-extract-kit",
        "model": str(model_path),
        "pages": limit,
        "candidates": candidates,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    (DONE / f"{job_id}.json").write_text(json.dumps({"output": str(output_path), "count": len(candidates)}, ensure_ascii=False), encoding="utf-8")


def main() -> int:
    if os.environ.get("USERNAME", "").lower().startswith("codexsandbox"):
        raise RuntimeError("PDF-Extract-Kit 桥接器必须在普通 Windows PowerShell 中启动")
    for folder in (QUEUE, DONE, ERRORS, BRIDGE / "work"):
        folder.mkdir(parents=True, exist_ok=True)
    READY.write_text(json.dumps({"pid": os.getpid(), "started": time.time()}, ensure_ascii=False), encoding="utf-8")
    print(f"PDF-Extract-Kit 公式检测桥接器已启动: {ROOT}")
    print("保持此窗口运行；按 Ctrl+C 停止。")
    try:
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
    finally:
        READY.unlink(missing_ok=True)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nPDF-Extract-Kit 桥接器已停止。")
