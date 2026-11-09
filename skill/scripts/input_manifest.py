#!/usr/bin/env python3
"""Discover and summarize the weekly Sunday-service source files."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from pptx_inventory import inventory


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def docx_paragraphs(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    paragraphs = []
    for para in root.findall(f".//{{{W}}}p"):
        chunks = []
        for node in para.iter():
            if node.tag == f"{{{W}}}t":
                chunks.append(node.text or "")
            elif node.tag == f"{{{W}}}tab":
                chunks.append("\t")
            elif node.tag == f"{{{W}}}br":
                chunks.append("\n")
        text = "".join(chunks).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def scripture_references(paragraphs: list[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    for paragraph in paragraphs:
        match = re.search(r"宣召[經经]文\s*[：:]\s*(.+)$", paragraph)
        if match:
            result["call_to_worship"] = match.group(1).strip()
        match = re.search(r"(?:講道|讲道|讀經|读经)[經经]文\s*[：:]\s*(.+)$", paragraph)
        if match:
            result["reading"] = match.group(1).strip()
    return result


def announcement_items(paragraphs: list[str]) -> list[str]:
    start = next((i for i, value in enumerate(paragraphs) if value.strip() == "報告事項"), None)
    if start is None:
        return []
    stop_labels = ("未來事工", "未来事工")
    items = []
    for value in paragraphs[start + 1 :]:
        if any(value.startswith(label) for label in stop_labels):
            break
        cleaned = re.sub(r"^[\s\u2022\u25cf\u25e6\-]+", "", value).strip()
        if cleaned:
            items.append(cleaned)
    return items


def _worship_title_slides(deck: dict) -> list[int]:
    result = []
    for slide in deck["slides"]:
        text = slide["text"]
        if re.search(r"(?:詞|词|曲)[：:]", text) and not re.search(r"\d+\s*/\s*\d+", text):
            result.append(slide["number"])
    return result


def _is_output_deck(path: Path) -> bool:
    return bool(re.search(r"^SundayPPT_\d{8}(?:_AM)?", path.name, re.I))


def _select_pptx(
    directory: Path,
    template_override: Path | None,
    worship_override: Path | None,
    sermon_override: Path | None,
) -> tuple[Path, Path, Path | None, dict, int]:
    decks = [path for path in directory.glob("*.pptx") if not path.name.startswith("~$")]
    template = template_override or next((path for path in decks if "template" in path.name.lower()), None)
    if template is None:
        raise FileNotFoundError(
            "No PPTX template found; put a template-named PPTX in the weekly folder "
            "or pass --template /private/path/template.pptx"
        )

    inventories = {}
    for path in decks:
        if path == template or _is_output_deck(path):
            continue
        try:
            inventories[path] = inventory(path)
        except Exception:
            continue

    worship = worship_override
    if worship is None:
        candidates = []
        for path, deck in inventories.items():
            titles = _worship_title_slides(deck)
            score = (10 if "worship" in path.name.lower() else 0) + len(titles)
            if len(titles) >= 4:
                candidates.append((score, path, deck, titles))
        if not candidates:
            raise FileNotFoundError("No worship PPTX with four song title slides found")
        _, worship, worship_deck, title_slides = max(candidates, key=lambda item: item[0])
    else:
        worship_deck = inventory(worship)
        title_slides = _worship_title_slides(worship_deck)
    if len(title_slides) < 4:
        raise ValueError(f"Expected four worship songs, found title slides {title_slides}")

    sermon = sermon_override
    if sermon is None:
        candidates = []
        for path, deck in inventories.items():
            if path == worship:
                continue
            score = (5 if re.fullmatch(r"\d{8}\.pptx", path.name) else 0) + (3 if deck["slide_count"] <= 20 else 0)
            candidates.append((score, path))
        sermon = max(candidates, key=lambda item: item[0])[1] if candidates else None

    sermon_slide_count = 0
    if sermon:
        sermon_deck = inventories.get(sermon) or inventory(sermon)
        sermon_slide_count = sermon_deck["slide_count"]

    return template, worship, sermon, {
        "worship_slide_count": worship_deck["slide_count"],
        "worship_title_slides": title_slides,
        "first_three_range": [title_slides[0], title_slides[3] - 1],
        "response_range": [title_slides[3], worship_deck["slide_count"]],
    }, sermon_slide_count


def _asset_text(path: Path) -> str:
    parts = [path.stem]
    if path.suffix.lower() == ".pdf":
        executable = shutil.which("pdftotext")
        if executable:
            result = subprocess.run(
                [executable, str(path), "-"],
                capture_output=True,
                text=True,
                check=False,
                timeout=30,
            )
            if result.returncode == 0:
                parts.append(result.stdout)
    return " ".join(parts)


def _tokens(value: str) -> set[str]:
    value = value.translate(str.maketrans({
        "会": "會",
        "员": "員",
        "礼": "禮",
        "调": "調",
        "问": "問",
        "间": "間",
        "报": "報",
        "时": "時",
    })).lower()
    tokens = set(re.findall(r"[a-z]{3,}|\d{2,}", value))
    for chunk in re.findall(r"[\u3400-\u9fff]{2,}", value):
        tokens.add(chunk)
        for size in (2, 3, 4):
            tokens.update(chunk[index : index + size] for index in range(len(chunk) - size + 1))
    return tokens


def asset_kind(path: Path) -> str:
    """Classify QR codes for same-slide placement; all other visuals get a new slide."""
    compact = re.sub(r"[\s_\-()（）]+", "", path.stem).lower()
    if path.suffix.lower() != ".pdf" and any(
        marker in compact for marker in ("qrcode", "qr", "二維碼", "二维码")
    ):
        return "qr"
    return "poster"


def _date_score(asset_name: str, announcement: str) -> float:
    digits = re.sub(r"\D", "", asset_name)
    score = 0.0
    for match in re.finditer(r"(20\d{2})(\d{2})(\d{2})", digits):
        month = str(int(match.group(2)))
        day = str(int(match.group(3)))
        if re.search(rf"(?:{month}\s*月\s*{day}\s*日|{month}\s*/\s*{day})", announcement):
            score += 6.0
    return score


def pair_assets(assets: list[Path], announcements: list[str]) -> list[dict]:
    assignments = []
    for path in assets:
        text = _asset_text(path)
        asset_tokens = _tokens(text)
        scored = []
        for index, announcement in enumerate(announcements, start=1):
            announcement_tokens = _tokens(announcement)
            overlap = len(asset_tokens & announcement_tokens)
            score = overlap + _date_score(path.name, announcement)
            for phrase in ("combined", "joint", "worship", "聯合", "联合"):
                if phrase in text.lower() and phrase in announcement.lower():
                    score += 2.0
            if "map" in path.stem.lower() and re.search(r"(?:地址|address|地點|地点)", announcement, re.I):
                score += 4.0
            scored.append((score, index))
        best_score, best_item = max(scored, key=lambda item: (item[0], -item[1]), default=(0.0, 0))
        assignments.append({
            "path": str(path.resolve()),
            "announcement": best_item if best_score > 0 else None,
            "score": best_score,
            "kind": asset_kind(path),
        })
    return assignments


def discover(
    directory: Path,
    template: Path | None = None,
    worship: Path | None = None,
    sermon: Path | None = None,
    scripture_docx: Path | None = None,
    bulletin_docx: Path | None = None,
) -> dict:
    directory = directory.resolve()
    docx_files = [path for path in directory.glob("*.docx") if not path.name.startswith("~$")]
    docx_contents = {path: docx_paragraphs(path) for path in docx_files}

    if scripture_docx is None:
        scripture_docx = next(
            (path for path, paragraphs in docx_contents.items() if scripture_references(paragraphs)),
            None,
        )
    if scripture_docx is None:
        raise FileNotFoundError("No scripture DOCX with call-to-worship and reading references found")
    scripture_paragraphs = docx_contents.get(scripture_docx) or docx_paragraphs(scripture_docx)
    references = scripture_references(scripture_paragraphs)
    if set(references) != {"call_to_worship", "reading"}:
        raise ValueError(f"Scripture DOCX is missing references: {references}")

    if bulletin_docx is None:
        bulletin_docx = next(
            (path for path in docx_files if "bulletin" in path.name.lower()),
            None,
        )
    if bulletin_docx is None:
        bulletin_docx = next(
            (path for path, paragraphs in docx_contents.items() if announcement_items(paragraphs)),
            None,
        )
    if bulletin_docx is None:
        raise FileNotFoundError("No bulletin DOCX with 報告事項 found")
    bulletin_paragraphs = docx_contents.get(bulletin_docx) or docx_paragraphs(bulletin_docx)
    announcements = announcement_items(bulletin_paragraphs)
    if not announcements:
        raise ValueError("No announcement items found in bulletin")

    template, worship, sermon, worship_info, sermon_slide_count = _select_pptx(
        directory,
        template.resolve() if template else None,
        worship.resolve() if worship else None,
        sermon.resolve() if sermon else None,
    )

    excluded_pdf_stems = {bulletin_docx.stem.lower()}
    assets = []
    for path in directory.iterdir():
        if path.suffix.lower() not in {".png", ".jpg", ".jpeg", ".pdf"}:
            continue
        if "bulletin" in path.name.lower() or path.stem.lower() in excluded_pdf_stems:
            continue
        assets.append(path)

    date_match = re.search(r"(20\d{2})[-_](\d{2})[-_](\d{2})", bulletin_docx.name)
    date_compact = "".join(date_match.groups()) if date_match else ""
    return {
        "directory": str(directory),
        "template": str(template.resolve()),
        "scripture_docx": str(scripture_docx.resolve()),
        "worship_pptx": str(worship.resolve()),
        "sermon_pptx": str(sermon.resolve()) if sermon else None,
        "sermon_slide_count": sermon_slide_count,
        "bulletin_docx": str(bulletin_docx.resolve()),
        "service_date": date_compact,
        "output_name": f"SundayPPT_{date_compact}_AM.pptx" if date_compact else "SundayPPT_AM.pptx",
        "scripture": references,
        "worship": worship_info,
        "announcements": announcements,
        "assets": pair_assets(sorted(assets), announcements),
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("directory", nargs="?", default=".", type=Path)
    parser.add_argument("--template", type=Path)
    parser.add_argument("--worship", type=Path)
    parser.add_argument("--sermon", type=Path)
    parser.add_argument("--scripture-docx", type=Path)
    parser.add_argument("--bulletin-docx", type=Path)
    args = parser.parse_args()
    print(json.dumps(discover(
        args.directory,
        template=args.template,
        worship=args.worship,
        sermon=args.sermon,
        scripture_docx=args.scripture_docx,
        bulletin_docx=args.bulletin_docx,
    ), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
