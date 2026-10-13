"""Deterministic OOXML cleanup for Word/WPS-compatible DOCX output.

This module intentionally operates on the DOCX package after python-docx has
finished formatting.  It does not reflow content or rewrite images; it only
removes converter-specific positioning metadata and unsafe hidden breaks.
"""

from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS = {"w": W_NS, "wp": WP_NS, "a": A_NS, "pic": PIC_NS}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _q(namespace: str, local: str) -> str:
    return "{" + namespace + "}" + local


def _remove_children(parent: ET.Element, names: set[str]) -> int:
    removed = 0
    for child in list(parent):
        if _local(child.tag) in names:
            parent.remove(child)
            removed += 1
    return removed


def _sync_picture_extents(inline: ET.Element) -> int:
    extents = inline.findall(_q(WP_NS, "extent"))
    if not extents:
        return 0
    extent = extents[0]
    cx, cy = extent.get("cx"), extent.get("cy")
    if not cx or not cy:
        return 0
    changed = 0
    for node in inline.findall(".//" + _q(A_NS, "ext")):
        if node.get("cx") != cx:
            node.set("cx", cx)
            changed += 1
        if node.get("cy") != cy:
            node.set("cy", cy)
            changed += 1
    return changed


def _sanitize_drawing(drawing: ET.Element) -> dict[str, int]:
    stats = {"anchors_inlined": 0, "position_nodes_removed": 0, "picture_attrs_removed": 0, "extent_repairs": 0}
    anchor_tag = _q(WP_NS, "anchor")
    inline_tag = _q(WP_NS, "inline")
    for node in list(drawing.iter()):
        if node.tag == anchor_tag:
            node.tag = inline_tag
            stats["anchors_inlined"] += 1
        if node.tag != inline_tag:
            continue
        for attr in ("behindDoc", "allowOverlap", "relativeHeight", "layoutInCell", "locked", "simplePos"):
            if attr in node.attrib:
                del node.attrib[attr]
                stats["picture_attrs_removed"] += 1
        stats["position_nodes_removed"] += _remove_children(
            node,
            {"positionH", "positionV", "wrapNone", "wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom", "simplePos"},
        )
        stats["extent_repairs"] += _sync_picture_extents(node)
    return stats


def _sanitize_xml(data: bytes, is_document: bool) -> tuple[bytes, dict[str, int]]:
    root = ET.fromstring(data)
    stats = {
        "anchors_inlined": 0,
        "position_nodes_removed": 0,
        "picture_attrs_removed": 0,
        "extent_repairs": 0,
        "last_rendered_breaks_removed": 0,
        "duplicate_page_breaks_removed": 0,
        "section_breaks_collapsed": 0,
        "visual_row_heights_removed": 0,
    }
    for drawing in root.findall(".//" + _q(W_NS, "drawing")):
        for key, value in _sanitize_drawing(drawing).items():
            stats[key] += value

    # LibreOffice and Word both understand these, but they are renderer
    # bookkeeping rather than document content and can create phantom gaps.
    for node in list(root.iter()):
        if node.tag == _q(W_NS, "lastRenderedPageBreak"):
            parent = next((p for p in root.iter() if node in list(p)), None)
            if parent is not None:
                parent.remove(node)
                stats["last_rendered_breaks_removed"] += 1

    # Remove only consecutive explicit page breaks in the same run/paragraph.
    # A single page break is meaningful and is preserved.
    for parent in root.iter():
        children = list(parent)
        previous_page_break = False
        for child in children:
            is_page_break = child.tag == _q(W_NS, "br") and child.get(_q(W_NS, "type")) == "page"
            if is_page_break and previous_page_break:
                parent.remove(child)
                stats["duplicate_page_breaks_removed"] += 1
            else:
                previous_page_break = is_page_break

    # Fixed converter-authored row heights can clip inline pictures or
    # equations after anchor-to-inline normalization. Remove them only from
    # rows that contain visual content; ordinary text tables are preserved.
    if is_document:
        math_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        for row in root.findall(".//" + _q(W_NS, "tr")):
            if not row.findall(".//" + _q(W_NS, "drawing")) and not row.findall(".//" + _q(math_ns, "oMath")) and not row.findall(".//" + _q(math_ns, "oMathPara")):
                continue
            tr_pr = row.find(_q(W_NS, "trPr"))
            if tr_pr is None:
                continue
            for height in list(tr_pr.findall(_q(W_NS, "trHeight"))):
                tr_pr.remove(height)
                stats["visual_row_heights_removed"] += 1

    # Keep the final document section properties. Paragraph-level section
    # breaks are already normalized by the formatter; do not remove a real
    # landscape/page-size transition here.
    if is_document:
        sects = root.findall(".//" + _q(W_NS, "sectPr"))
        if len(sects) > 1:
            # Collapse adjacent identical section properties only.
            for first, second in zip(sects, sects[1:]):
                if ET.tostring(first, encoding="unicode") == ET.tostring(second, encoding="unicode"):
                    parent = next((p for p in root.iter() if second in list(p)), None)
                    if parent is not None:
                        parent.remove(second)
                        stats["section_breaks_collapsed"] += 1

    ET.register_namespace("w", W_NS)
    ET.register_namespace("wp", WP_NS)
    ET.register_namespace("a", A_NS)
    ET.register_namespace("pic", PIC_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), stats


def sanitize_docx(path: Path, output: Path | None = None) -> dict[str, Any]:
    """Sanitize a DOCX package in place (or write to ``output``).

    The original package is copied unchanged when XML parsing fails. The
    returned report records the package entries examined and every mutation.
    """
    path = Path(path)
    target = Path(output) if output else path
    target.parent.mkdir(parents=True, exist_ok=True)
    temp_dir = Path(tempfile.mkdtemp(prefix="docx-sanitize-", dir=str(target.parent)))
    temp_path = temp_dir / target.name
    totals: dict[str, int] = {
        "anchors_inlined": 0,
        "position_nodes_removed": 0,
        "picture_attrs_removed": 0,
        "extent_repairs": 0,
        "last_rendered_breaks_removed": 0,
        "duplicate_page_breaks_removed": 0,
        "section_breaks_collapsed": 0,
        "visual_row_heights_removed": 0,
    }
    xml_entries = 0
    try:
        with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as dest:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                    try:
                        data, stats = _sanitize_xml(data, info.filename == "word/document.xml")
                        xml_entries += 1
                        for key, value in stats.items():
                            totals[key] += value
                    except ET.ParseError:
                        pass
                dest.writestr(info, data)
        if target.resolve() == path.resolve():
            shutil.copy2(temp_path, path)
        else:
            shutil.copy2(temp_path, target)
        return {"checked": True, "output": str(target), "xml_entries": xml_entries, **totals}
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def validate_docx(path: Path) -> dict[str, Any]:
    """Run structural checks without opening Word or LibreOffice."""
    path = Path(path)
    checks = {"zip_readable": False, "document_xml": False, "anchors": 0, "inline_pictures": 0, "full_page_images": 0}
    try:
        with zipfile.ZipFile(path, "r") as package:
            checks["zip_readable"] = package.testzip() is None
            document = ET.fromstring(package.read("word/document.xml"))
            checks["document_xml"] = True
            checks["anchors"] = len(document.findall(".//" + _q(WP_NS, "anchor")))
            checks["inline_pictures"] = len(document.findall(".//" + _q(WP_NS, "inline") + "/" + _q(WP_NS, "extent")))
            checks["full_page_images"] = sum(1 for extent in document.findall(".//" + _q(WP_NS, "inline") + "/" + _q(WP_NS, "extent")) if int(extent.get("cy", "0")) >= 10000000)
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError, ValueError):
        pass
    checks["passed"] = bool(checks["zip_readable"] and checks["document_xml"] and checks["anchors"] == 0)
    return checks
