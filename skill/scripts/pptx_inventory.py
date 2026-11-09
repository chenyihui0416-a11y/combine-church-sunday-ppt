#!/usr/bin/env python3
"""Inspect PPTX slide order and content without modifying the package."""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import zipfile
from collections import Counter
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _xml(archive: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(archive.read(name))


def _rels(archive: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    pure = PurePosixPath(part)
    rel_name = str(pure.parent / "_rels" / f"{pure.name}.rels")
    if rel_name not in archive.namelist():
        return {}
    root = _xml(archive, rel_name)
    result: dict[str, tuple[str, str]] = {}
    for rel in root.findall("pr:Relationship", NS):
        result[rel.attrib["Id"]] = (rel.attrib.get("Type", ""), rel.attrib.get("Target", ""))
    return result


def _resolve(part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _text(root: ET.Element) -> str:
    paragraphs: list[str] = []
    for para in root.findall(".//a:p", NS):
        value = "".join(node.text or "" for node in para.findall(".//a:t", NS)).strip()
        if value:
            paragraphs.append(value)
    return "\n".join(paragraphs)


def _normalized(value: str) -> str:
    return re.sub(r"\s+", "", value).lower()


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def inventory(path: Path) -> dict:
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        presentation = _xml(archive, "ppt/presentation.xml")
        presentation_rels = _rels(archive, "ppt/presentation.xml")
        size = presentation.find("p:sldSz", NS)
        slide_size = {
            "cx": int(size.attrib["cx"]) if size is not None else None,
            "cy": int(size.attrib["cy"]) if size is not None else None,
        }
        slide_parts: list[str] = []
        for slide_id in presentation.findall(".//p:sldId", NS):
            rel_id = slide_id.attrib[f"{{{NS['r']}}}id"]
            slide_parts.append(_resolve("ppt/presentation.xml", presentation_rels[rel_id][1]))

        slides = []
        for index, part in enumerate(slide_parts, start=1):
            root = _xml(archive, part)
            relationships = _rels(archive, part)
            related = []
            media_hashes = []
            layout = None
            notes_text = ""
            for rel_id, (rel_type, target) in relationships.items():
                resolved = _resolve(part, target)
                kind = rel_type.rsplit("/", 1)[-1]
                related.append({"id": rel_id, "kind": kind, "target": resolved})
                if kind == "slideLayout":
                    layout = resolved
                elif kind in {"image", "audio", "video", "media"} and resolved in names:
                    media_hashes.append({
                        "kind": kind,
                        "target": resolved,
                        "sha256": _sha256(archive.read(resolved)),
                    })
                elif kind == "notesSlide" and resolved in names:
                    notes_text = _text(_xml(archive, resolved))

            text = _text(root)
            fonts = sorted({
                node.attrib["typeface"]
                for node in root.findall(".//*[@typeface]")
                if node.attrib.get("typeface")
            })
            slides.append({
                "number": index,
                "part": part,
                "text": text,
                "normalized_text": _normalized(text),
                "text_sha256": _sha256(_normalized(text).encode("utf-8")),
                "layout": layout,
                "shape_count": sum(len(root.findall(f".//p:{tag}", NS)) for tag in ("sp", "pic", "graphicFrame", "grpSp")),
                "picture_count": len(root.findall(".//p:pic", NS)),
                "hidden": root.attrib.get("show") == "0",
                "has_transition": root.find("p:transition", NS) is not None,
                "has_timing": root.find("p:timing", NS) is not None,
                "has_table": root.find(".//a:tbl", NS) is not None,
                "fonts": fonts,
                "media": sorted(media_hashes, key=lambda item: (item["kind"], item["sha256"])),
                "notes": notes_text,
                "relationships": sorted(related, key=lambda item: item["id"]),
            })

        return {
            "path": str(path.resolve()),
            "sha256": _sha256(path.read_bytes()),
            "slide_size": slide_size,
            "slide_count": len(slides),
            "slides": slides,
        }


def _tokens(value: str) -> Counter[str]:
    return Counter(re.findall(r"[\w\u3400-\u9fff]+", value.lower()))


def _overlap(left: Counter[str], right: Counter[str]) -> float:
    union = sum((left | right).values())
    return sum((left & right).values()) / union if union else 0.0


def best_matches(target: dict, sources: list[dict]) -> list[dict]:
    rows = []
    for slide in target["slides"]:
        candidates = []
        target_media = {item["sha256"] for item in slide["media"]}
        target_tokens = _tokens(slide["text"])
        for source in sources:
            for candidate in source["slides"]:
                candidate_media = {item["sha256"] for item in candidate["media"]}
                exact_text = bool(slide["normalized_text"]) and slide["normalized_text"] == candidate["normalized_text"]
                text_score = _overlap(target_tokens, _tokens(candidate["text"]))
                media_score = (
                    len(target_media & candidate_media) / len(target_media | candidate_media)
                    if target_media or candidate_media
                    else 0.0
                )
                shape_delta = abs(slide["shape_count"] - candidate["shape_count"])
                score = (4.0 if exact_text else 0.0) + (2.0 * text_score) + media_score - min(shape_delta, 10) * 0.01
                candidates.append({
                    "score": round(score, 5),
                    "exact_text": exact_text,
                    "text_score": round(text_score, 5),
                    "media_score": round(media_score, 5),
                    "source": source["path"],
                    "source_slide": candidate["number"],
                    "source_text": candidate["text"],
                })
        candidates.sort(key=lambda item: item["score"], reverse=True)
        rows.append({
            "target_slide": slide["number"],
            "target_text": slide["text"],
            "best": candidates[:5],
        })
    return rows


def main() -> None:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)
    inspect_parser = subparsers.add_parser("inspect")
    inspect_parser.add_argument("pptx", type=Path)
    summary_parser = subparsers.add_parser("summary")
    summary_parser.add_argument("pptx", type=Path)
    match_parser = subparsers.add_parser("match")
    match_parser.add_argument("target", type=Path)
    match_parser.add_argument("sources", nargs="+", type=Path)
    args = parser.parse_args()

    if args.command == "inspect":
        output = inventory(args.pptx)
    elif args.command == "summary":
        data = inventory(args.pptx)
        for slide in data["slides"]:
            one_line = slide["text"].replace("\t", " ").replace("\n", " / ")
            state = "hidden" if slide["hidden"] else "shown"
            print(f"{slide['number']:02d}\t{state}\t{one_line}")
        return
    else:
        target = inventory(args.target)
        sources = [inventory(path) for path in args.sources]
        output = {
            "target": target["path"],
            "sources": [source["path"] for source in sources],
            "matches": best_matches(target, sources),
        }
    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
