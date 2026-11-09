#!/usr/bin/env python3
"""Assemble a Sunday service deck while preserving source PowerPoint slides."""

from __future__ import annotations

import argparse
import base64
import copy
import hashlib
import json
import platform
import posixpath
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import time
import zipfile
from collections import defaultdict
from dataclasses import asdict
from pathlib import Path
from xml.etree import ElementTree as ET

from input_manifest import discover
from pptx_inventory import inventory
from spring_bible import Verse, fetch_reference, parse_reference


CALL_SHAPE = "Google Shape;261;p3"
SCRIPTURE_TITLE_SHAPE = "Google Shape;297;p9"
SCRIPTURE_BODY_SHAPE = "Google Shape;298;p9"
ANNOUNCEMENT_TITLE_SHAPE = "Google Shape;337;p15"
ANNOUNCEMENT_BODY_SHAPE = "Google Shape;338;p15"
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPE_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def as_quote(value: str | Path) -> str:
    """Return an AppleScript string literal."""
    value = str(value).replace("\\", "\\\\").replace('"', '\\"')
    return f'"{value}"'


def as_text_expression(value: str, paragraphs: bool = False) -> str:
    lines = value.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    separator = " & return & " if paragraphs else " & linefeed & "
    return separator.join(as_quote(line) for line in lines)


def ps_quote(value: str | Path) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def ps_text(value: str) -> str:
    payload = base64.b64encode(value.encode("utf-8")).decode("ascii")
    return f"(Decode {ps_quote(payload)})"


def chunk_reading(verses: list[Verse], maximum_verses: int = 4) -> list[list[Verse]]:
    """Keep chapters separate and use at most four numbered verses per slide."""
    by_chapter: dict[int, list[Verse]] = defaultdict(list)
    chapter_order: list[int] = []
    for verse in verses:
        if verse.chapter not in by_chapter:
            chapter_order.append(verse.chapter)
        by_chapter[verse.chapter].append(verse)
    pages: list[list[Verse]] = []
    for chapter in chapter_order:
        chapter_verses = by_chapter[chapter]
        pages.extend(
            chapter_verses[index : index + maximum_verses]
            for index in range(0, len(chapter_verses), maximum_verses)
        )
    return pages


def reference_for_page(book: str, verses: list[Verse]) -> str:
    return f"{book} {verses[0].chapter}:{verses[0].verse}-{verses[-1].verse}"


def split_announcement(text: str, maximum_characters: int = 220) -> list[str]:
    """Split only genuinely long announcement items at sentence boundaries."""
    if len(text) <= maximum_characters:
        return [text]
    sentences = [part for part in re.split(r"(?<=[。！？；])", text) if part]
    if len(sentences) == 1:
        sentences = [part for part in re.split(r"(?<=[，,、])", text) if part]
    pages: list[str] = []
    current = ""
    for sentence in sentences:
        if current and len(current) + len(sentence) > maximum_characters:
            pages.append(current)
            current = ""
        while len(sentence) > maximum_characters:
            available = maximum_characters - len(current)
            current += sentence[:available]
            pages.append(current)
            current = ""
            sentence = sentence[available:]
        current += sentence
    if current:
        pages.append(current)
    return pages


def render_pdf_first_page(path: Path, output_dir: Path) -> Path:
    executable = shutil.which("pdftoppm")
    if not executable:
        raise RuntimeError("pdftoppm is required to turn an announcement PDF into an image")
    stem = output_dir / f"{path.stem}-page-1"
    result = subprocess.run(
        [executable, "-png", "-f", "1", "-singlefile", "-r", "144", str(path), str(stem)],
        capture_output=True,
        text=True,
        check=False,
        timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(f"Could not render {path.name}: {result.stderr.strip()}")
    rendered = stem.with_suffix(".png")
    width, height = image_dimensions(rendered)
    desired_ratio = 763.37 / 575.28
    target_height = round(width / desired_ratio) if width else height
    if width and 0 < target_height < height:
        try:
            from PIL import Image

            with Image.open(rendered) as image:
                top = max(0, (height - target_height) // 2)
                image.crop((0, top, width, top + target_height)).save(rendered)
        except ImportError:
            # The uncropped page still remains usable. Pillow is only needed for
            # reproducing the established landscape treatment of rare PDF assets.
            pass
    return rendered


def image_dimensions(path: Path) -> tuple[int, int]:
    data = path.read_bytes()
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        return struct.unpack("<HH", data[6:10])
    if data.startswith(b"\xff\xd8"):
        index = 2
        start_of_frame = {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}
        while index + 9 < len(data):
            if data[index] != 0xFF:
                index += 1
                continue
            while index < len(data) and data[index] == 0xFF:
                index += 1
            if index >= len(data):
                break
            marker = data[index]
            index += 1
            if marker in (0xD8, 0xD9):
                continue
            if index + 2 > len(data):
                break
            length = int.from_bytes(data[index : index + 2], "big")
            if marker in start_of_frame and index + 7 < len(data):
                height = int.from_bytes(data[index + 3 : index + 5], "big")
                width = int.from_bytes(data[index + 5 : index + 7], "big")
                return width, height
            index += max(length, 2)
    return (0, 0)


def fitted_geometry(path: Path, from_pdf: bool = False) -> tuple[float, float, float, float]:
    """Return 16:9 slide coordinates in PowerPoint points."""
    if from_pdf:
        # Matches the established landscape screenshot treatment for a letter-size map.
        return (98.315, 0.0, 763.37, 575.28)
    width, height = image_dimensions(path)
    if not width or not height:
        return (0.0, 0.0, 960.0, 540.0)
    ratio = width / height
    if abs(ratio - (16 / 9)) < 0.025:
        return (0.0, 0.0, 960.0, 540.0)
    fitted_width = min(960.0, 540.0 * ratio)
    fitted_height = fitted_width / ratio
    if fitted_height > 540.0:
        fitted_height = 540.0
        fitted_width = fitted_height * ratio
    return ((960.0 - fitted_width) / 2, (540.0 - fitted_height) / 2, fitted_width, fitted_height)


def qr_geometry(path: Path, announcement: str) -> tuple[float, float, float, float]:
    """Center a QR code below its announcement while preserving its aspect ratio."""
    width, height = image_dimensions(path)
    ratio = width / height if width and height else 1.0
    compact_length = len(re.sub(r"\s+", "", announcement))
    if compact_length >= 120:
        fitted_height = 185.7
        top = 354.3
    else:
        fitted_height = 260.9
        top = 270.0
    fitted_width = fitted_height * ratio
    if fitted_width > 275.0:
        fitted_width = 275.0
        fitted_height = fitted_width / ratio
    return ((960.0 - fitted_width) / 2, top, fitted_width, fitted_height)


class AppleScriptDeck:
    def __init__(self, target: Path, template_slide_count: int):
        self.target = target
        self.target_name = target.name
        self.template_slide_count = template_slide_count
        self.lines = [
            'tell application "Microsoft PowerPoint"',
            "activate",
            "with timeout of 900 seconds",
            f"open POSIX file {as_quote(target)}",
            "delay 1",
        ]
        self.current_index = template_slide_count
        self.source_names: set[str] = set()

    def append_template(self, slide_number: int) -> int:
        self.lines.extend([
            f"tell presentation {as_quote(self.target_name)}",
            f"copy object slide {slide_number}",
            "paste object",
            f"set design of slide (count of slides) to design of slide {slide_number}",
            "end tell",
        ])
        self.current_index += 1
        return self.current_index

    def append_source(self, source: Path, slide_number: int) -> int:
        source_name = source.name
        if source_name not in self.source_names:
            self.source_names.add(source_name)
            self.lines.extend([
                f"open POSIX file {as_quote(source)}",
                "delay 0.2",
            ])
        self.lines.extend([
            f"tell presentation {as_quote(source_name)} to copy object slide {slide_number}",
            f"tell presentation {as_quote(self.target_name)}",
            "paste object",
            f"set design of slide (count of slides) to design of slide {slide_number} of presentation {as_quote(source_name)}",
            "end tell",
        ])
        self.current_index += 1
        return self.current_index

    def set_text(
        self,
        slide: int,
        shape: str,
        value: str,
        font_size: float | None = None,
        paragraphs: bool = False,
    ) -> None:
        self.lines.append(
            f"set content of text range of text frame of shape {as_quote(shape)} of slide {slide} "
            f"of presentation {as_quote(self.target_name)} to {as_text_expression(value, paragraphs=paragraphs)}"
        )
        if font_size:
            self.lines.append(
                f"set font size of font of text range of text frame of shape {as_quote(shape)} "
                f"of slide {slide} of presentation {as_quote(self.target_name)} to {font_size}"
            )

    def number_paragraphs(self, slide: int, shape: str, start: int | None) -> None:
        target = (
            f"paragraph format of text range of text frame of shape {as_quote(shape)} "
            f"of slide {slide} of presentation {as_quote(self.target_name)}"
        )
        commands = [
            f"set bullet type of bullet format of {target} to bullet type numbered",
            f"set bullet style of bullet format of {target} to numbered bullet style arabic period",
        ]
        if start is not None:
            commands.append(f"set bullet start value of bullet format of {target} to {start}")
        self.lines.extend(commands)

    def remove_bullets(self, slide: int, shape: str) -> None:
        target = (
            f"paragraph format of text range of text frame of shape {as_quote(shape)} "
            f"of slide {slide} of presentation {as_quote(self.target_name)}"
        )
        self.lines.append(f"set bullet type of bullet format of {target} to bullet type none")

    def center_shape_vertically(self, slide: int, shape: str) -> None:
        target = (
            f"shape {as_quote(shape)} of slide {slide} "
            f"of presentation {as_quote(self.target_name)}"
        )
        self.lines.append(f"set top of {target} to (540 - (height of {target})) / 2")

    def add_image_slide(self, image: Path, from_pdf: bool = False) -> int:
        slide = self.append_template(15)
        left, top, width, height = fitted_geometry(image, from_pdf=from_pdf)
        self.lines.extend([
            f"set imageSlide to slide {slide} of presentation {as_quote(self.target_name)}",
            "tell imageSlide",
            "repeat while (count of shapes) > 0",
            "delete shape 1",
            "end repeat",
            "end tell",
            f"set imageShape to make new picture at imageSlide with properties {{file name:{as_quote(image)}, left position:{left}, top:{top}, width:{width}, height:{height}}}",
        ])
        return slide

    def add_picture(self, slide: int, image: Path, geometry: tuple[float, float, float, float]) -> None:
        left, top, width, height = geometry
        self.lines.append(
            f"make new picture at slide {slide} of presentation {as_quote(self.target_name)} "
            f"with properties {{file name:{as_quote(image)}, left position:{left}, top:{top}, "
            f"width:{width}, height:{height}}}"
        )

    def append_blank(self, template_slide: int = 10) -> int:
        slide = self.append_template(template_slide)
        self.lines.extend([
            f"tell slide {slide} of presentation {as_quote(self.target_name)}",
            "repeat while (count of shapes) > 0",
            "delete shape 1",
            "end repeat",
            "end tell",
        ])
        return slide

    def hide_slide(self, slide: int) -> None:
        self.lines.append(
            f"set hidden of slide show transition of slide {slide} of presentation {as_quote(self.target_name)} to true"
        )

    def finish(self) -> str:
        self.lines.extend([
            f"tell presentation {as_quote(self.target_name)}",
            f"repeat {self.template_slide_count} times",
            "delete slide 1",
            "end repeat",
            "save",
            "set finalCount to count of slides",
            "end tell",
        ])
        for source_name in sorted(self.source_names):
            self.lines.extend([
                "try",
                f"close presentation {as_quote(source_name)} saving no",
                "end try",
            ])
        self.lines.extend([
            "try",
            f"close presentation {as_quote(self.target_name)} saving no",
            "end try",
            "return finalCount",
            "end timeout",
            "end tell",
        ])
        return "\n".join(self.lines)


class PowerShellDeck:
    """Generate Windows PowerPoint COM automation with the same deck API."""

    def __init__(self, target: Path, template_slide_count: int):
        self.target = target
        self.template_slide_count = template_slide_count
        self.current_index = template_slide_count
        self.lines = [
            "$ErrorActionPreference = 'Stop'",
            "function Decode([string]$value) { [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($value)) }",
            "$powerpoint = New-Object -ComObject PowerPoint.Application",
            "$powerpoint.Visible = -1",
            f"$target = $powerpoint.Presentations.Open({ps_quote(target)}, 0, 0, -1)",
        ]

    def append_template(self, slide_number: int) -> int:
        self.lines.extend([
            f"$target.Slides.Item({slide_number}).Copy()",
            "$null = $target.Slides.Paste($target.Slides.Count + 1)",
        ])
        self.current_index += 1
        return self.current_index

    def append_source(self, source: Path, slide_number: int) -> int:
        self.lines.append(
            f"$null = $target.Slides.InsertFromFile({ps_quote(source)}, $target.Slides.Count, {slide_number}, {slide_number})"
        )
        self.current_index += 1
        return self.current_index

    def set_text(
        self,
        slide: int,
        shape: str,
        value: str,
        font_size: float | None = None,
        paragraphs: bool = False,
    ) -> None:
        del paragraphs
        self.lines.extend([
            f"$shape = $target.Slides.Item({slide}).Shapes.Item({ps_quote(shape)})",
            f"$shape.TextFrame.TextRange.Text = {ps_text(value)}",
        ])
        if font_size:
            self.lines.append(f"$shape.TextFrame.TextRange.Font.Size = {font_size}")

    def number_paragraphs(self, slide: int, shape: str, start: int | None) -> None:
        self.lines.extend([
            f"$range = $target.Slides.Item({slide}).Shapes.Item({ps_quote(shape)}).TextFrame.TextRange",
            "$range.ParagraphFormat.Bullet.Type = 2",
        ])
        if start is not None:
            self.lines.append(f"$range.ParagraphFormat.Bullet.StartValue = {start}")

    def remove_bullets(self, slide: int, shape: str) -> None:
        self.lines.extend([
            f"$range = $target.Slides.Item({slide}).Shapes.Item({ps_quote(shape)}).TextFrame.TextRange",
            "$range.ParagraphFormat.Bullet.Type = 0",
        ])

    def center_shape_vertically(self, slide: int, shape: str) -> None:
        self.lines.extend([
            f"$shape = $target.Slides.Item({slide}).Shapes.Item({ps_quote(shape)})",
            "$shape.Top = (540 - $shape.Height) / 2",
        ])

    def add_image_slide(self, image: Path, from_pdf: bool = False) -> int:
        slide = self.append_template(15)
        left, top, width, height = fitted_geometry(image, from_pdf=from_pdf)
        self.lines.extend([
            f"$slide = $target.Slides.Item({slide})",
            "while ($slide.Shapes.Count -gt 0) { $slide.Shapes.Item(1).Delete() }",
            f"$null = $slide.Shapes.AddPicture({ps_quote(image)}, 0, -1, {left}, {top}, {width}, {height})",
        ])
        return slide

    def add_picture(self, slide: int, image: Path, geometry: tuple[float, float, float, float]) -> None:
        left, top, width, height = geometry
        self.lines.append(
            f"$null = $target.Slides.Item({slide}).Shapes.AddPicture({ps_quote(image)}, 0, -1, {left}, {top}, {width}, {height})"
        )

    def append_blank(self, template_slide: int = 10) -> int:
        slide = self.append_template(template_slide)
        self.lines.extend([
            f"$slide = $target.Slides.Item({slide})",
            "while ($slide.Shapes.Count -gt 0) { $slide.Shapes.Item(1).Delete() }",
        ])
        return slide

    def hide_slide(self, slide: int) -> None:
        self.lines.append(f"$target.Slides.Item({slide}).SlideShowTransition.Hidden = -1")

    def finish(self) -> str:
        self.lines.extend([
            f"1..{self.template_slide_count} | ForEach-Object {{ $target.Slides.Item(1).Delete() }}",
            "$target.Save()",
            "$finalCount = $target.Slides.Count",
            "$target.Close()",
            "$powerpoint.Quit()",
            "Write-Output $finalCount",
        ])
        return "\n".join(self.lines)


def select_engine(requested: str) -> str:
    if requested != "auto":
        return requested
    system = platform.system().lower()
    if system == "darwin":
        return "macos"
    if system == "windows":
        return "windows"
    raise RuntimeError("Native PowerPoint automation is supported on macOS and Windows")


def _relationship_part(part: str) -> str:
    pure = Path(part)
    return posixpath.join(posixpath.dirname(part), "_rels", f"{pure.name}.rels")


def _resolved_part(part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def repair_media_effect_slides(
    output: Path,
    copied_slides: list[tuple[Path, int, int]],
    output_inventory: dict | None = None,
    source_inventories: dict[Path, dict] | None = None,
) -> list[int]:
    """Restore source shape trees and picture backgrounds omitted by the clipboard."""
    output_inventory = output_inventory or inventory(output)
    source_inventories = source_inventories or {}
    with zipfile.ZipFile(output) as archive:
        infos = archive.infolist()
        parts = {info.filename: archive.read(info.filename) for info in infos}

    source_cache: dict[Path, tuple[dict, dict[str, bytes]]] = {}
    for source in dict.fromkeys(item[0] for item in copied_slides):
        with zipfile.ZipFile(source) as source_archive:
            source_parts = {
                name: source_archive.read(name)
                for name in source_archive.namelist()
            }
        source_cache[source] = (
            source_inventories.get(source) or inventory(source),
            source_parts,
        )

    repaired: list[int] = []
    extra_parts: dict[str, bytes] = {}
    content_types = ET.fromstring(parts["[Content_Types].xml"])
    content_defaults = {
        node.attrib.get("Extension", "").lower()
        for node in content_types.findall(f"{{{CONTENT_TYPE_NS}}}Default")
    }

    for source, source_number, output_number in copied_slides:
        source_inventory, source_parts = source_cache[source]
        source_part = source_inventory["slides"][source_number - 1]["part"]
        output_part = output_inventory["slides"][output_number - 1]["part"]
        source_rels_part = _relationship_part(source_part)
        output_rels_part = _relationship_part(output_part)

        if source_rels_part not in source_parts:
            continue
        source_rels = ET.fromstring(source_parts[source_rels_part])
        source_relationships = list(source_rels)
        has_picture_relationship = any(
            rel.attrib.get("Type", "").endswith(("/image", "/hdphoto"))
            for rel in source_relationships
        )
        source_slide = source_inventory["slides"][source_number - 1]
        output_slide = output_inventory["slides"][output_number - 1]
        source_media = sorted(item["sha256"] for item in source_slide["media"])
        output_media = sorted(item["sha256"] for item in output_slide["media"])
        source_root = ET.fromstring(source_parts[source_part])
        output_root = ET.fromstring(parts[output_part])
        source_color_map = source_root.find(f"{{{PRESENTATION_NS}}}clrMapOvr")
        output_color_map = output_root.find(f"{{{PRESENTATION_NS}}}clrMapOvr")
        source_background = source_root.find(
            f"{{{PRESENTATION_NS}}}cSld/{{{PRESENTATION_NS}}}bg"
        )
        output_background = output_root.find(
            f"{{{PRESENTATION_NS}}}cSld/{{{PRESENTATION_NS}}}bg"
        )
        color_map_changed = ET.tostring(source_color_map) != ET.tostring(output_color_map)
        background_changed = (
            (source_background is None) != (output_background is None)
            or (
                source_background is not None
                and output_background is not None
                and ET.tostring(source_background) != ET.tostring(output_background)
            )
        )
        if not has_picture_relationship or (
            source_media == output_media
            and source_slide["shape_count"] == output_slide["shape_count"]
            and not color_map_changed
            and not background_changed
        ):
            continue

        output_tree = output_root.find(f".//{{{PRESENTATION_NS}}}spTree")
        source_tree = source_root.find(f".//{{{PRESENTATION_NS}}}spTree")
        if output_tree is None or source_tree is None:
            raise RuntimeError(f"Could not repair shape tree for {source.name} slide {source_number}")
        parent = output_root.find(f".//{{{PRESENTATION_NS}}}cSld")
        if parent is None:
            raise RuntimeError(f"Could not find cSld for output slide {output_number}")
        if output_background is not None:
            parent.remove(output_background)
        if source_background is not None:
            parent.insert(0, copy.deepcopy(source_background))
        position = list(parent).index(output_tree)
        parent.remove(output_tree)
        parent.insert(position, copy.deepcopy(source_tree))
        if output_color_map is not None:
            color_position = list(output_root).index(output_color_map)
            output_root.remove(output_color_map)
            if source_color_map is not None:
                output_root.insert(color_position, copy.deepcopy(source_color_map))
        elif source_color_map is not None:
            output_root.append(copy.deepcopy(source_color_map))
        parts[output_part] = ET.tostring(output_root, encoding="utf-8", xml_declaration=True)

        output_rels = ET.fromstring(parts[output_rels_part])
        output_layout = next(
            (rel for rel in output_rels if rel.attrib.get("Type", "").endswith("/slideLayout")),
            None,
        )
        if output_layout is None:
            raise RuntimeError(f"No slide layout relationship for output slide {output_number}")

        rebuilt_rels = ET.Element(f"{{{PACKAGE_REL_NS}}}Relationships")
        used_ids: set[str] = set()
        for relationship in source_relationships:
            attributes = dict(relationship.attrib)
            relation_type = attributes.get("Type", "")
            relation_id = attributes["Id"]
            if relation_type.endswith(("/notesSlide", "/comments")):
                continue
            if relation_type.endswith("/slideLayout"):
                attributes["Target"] = output_layout.attrib["Target"]
            elif attributes.get("TargetMode") != "External":
                source_target = _resolved_part(source_part, attributes["Target"])
                if relation_type.endswith(("/image", "/hdphoto")):
                    payload = source_parts[source_target]
                    digest = hashlib.sha256(payload).hexdigest()[:12]
                    media_name = f"church-{digest}-{Path(source_target).name}"
                    destination_target = f"ppt/media/{media_name}"
                    extra_parts[destination_target] = payload
                    attributes["Target"] = f"../media/{media_name}"
                    extension = Path(media_name).suffix.lstrip(".").lower()
                    if extension not in content_defaults:
                        source_types = ET.fromstring(source_parts["[Content_Types].xml"])
                        source_default = next(
                            (
                                node
                                for node in source_types.findall(f"{{{CONTENT_TYPE_NS}}}Default")
                                if node.attrib.get("Extension", "").lower() == extension
                            ),
                            None,
                        )
                        if source_default is not None:
                            content_types.append(copy.deepcopy(source_default))
                            content_defaults.add(extension)
                else:
                    raise RuntimeError(
                        f"Unsupported relationship {relation_type} on WDP slide "
                        f"{source.name}#{source_number}"
                    )
            rebuilt_rels.append(ET.Element(f"{{{PACKAGE_REL_NS}}}Relationship", attributes))
            used_ids.add(relation_id)

        for relationship in output_rels:
            relation_type = relationship.attrib.get("Type", "")
            if not relation_type.endswith(("/notesSlide", "/comments")):
                continue
            attributes = dict(relationship.attrib)
            suffix = 90
            while f"rId{suffix}" in used_ids:
                suffix += 1
            attributes["Id"] = f"rId{suffix}"
            used_ids.add(attributes["Id"])
            rebuilt_rels.append(ET.Element(f"{{{PACKAGE_REL_NS}}}Relationship", attributes))

        ET.register_namespace("", PACKAGE_REL_NS)
        parts[output_rels_part] = ET.tostring(
            rebuilt_rels,
            encoding="utf-8",
            xml_declaration=True,
        )
        repaired.append(output_number)

    if not repaired:
        return []

    ET.register_namespace("", CONTENT_TYPE_NS)
    parts["[Content_Types].xml"] = ET.tostring(
        content_types,
        encoding="utf-8",
        xml_declaration=True,
    )
    temporary = output.with_name(f".{output.stem}.repairing.pptx")
    existing_names = {info.filename for info in infos}
    with zipfile.ZipFile(temporary, "w") as archive:
        for info in infos:
            archive.writestr(info, parts[info.filename])
        for name, payload in extra_parts.items():
            if name not in existing_names:
                archive.writestr(name, payload, compress_type=zipfile.ZIP_DEFLATED)
    temporary.replace(output)
    return repaired


def validate_native_copies(
    output: Path,
    copied_slides: list[tuple[Path, int, int]],
    output_inventory: dict | None = None,
    source_inventories: dict[Path, dict] | None = None,
) -> list[dict]:
    """Fail when an immutable worship, sermon, or template slide changed during copying."""
    output_inventory = output_inventory or inventory(output)
    source_cache: dict[Path, dict] = source_inventories or {}
    failures: list[dict] = []
    fields = (
        "normalized_text",
        "shape_count",
        "picture_count",
        "has_transition",
        "has_timing",
    )
    for source, source_number, output_number in copied_slides:
        source_inventory = source_cache.setdefault(source, inventory(source))
        expected = source_inventory["slides"][source_number - 1]
        actual = output_inventory["slides"][output_number - 1]
        differences = {
            field: {"expected": expected[field], "actual": actual[field]}
            for field in fields
            if expected[field] != actual[field]
        }
        expected_media = sorted(item["sha256"] for item in expected["media"])
        actual_media = sorted(item["sha256"] for item in actual["media"])
        if expected_media != actual_media:
            differences["media_sha256"] = {
                "expected": expected_media,
                "actual": actual_media,
            }
        if differences:
            failures.append({
                "source": str(source),
                "source_slide": source_number,
                "output_slide": output_number,
                "differences": differences,
            })
    if failures:
        raise RuntimeError(
            "Native slide-copy validation failed:\n"
            + json.dumps(failures, ensure_ascii=False, indent=2)
        )
    return [
        {
            "source": str(source),
            "source_slide": source_number,
            "output_slide": output_number,
        }
        for source, source_number, output_number in copied_slides
    ]


def prepare_assets(manifest: dict, build_dir: Path) -> dict[int, list[dict]]:
    result: dict[int, list[dict]] = defaultdict(list)
    assigned = [item for item in manifest["assets"] if item["announcement"]]
    assigned.sort(key=lambda item: (item["announcement"], Path(item["path"]).suffix.lower() == ".pdf", item["path"]))
    for item in assigned:
        source = Path(item["path"])
        is_pdf = source.suffix.lower() == ".pdf"
        rendered = render_pdf_first_page(source, build_dir) if is_pdf else source
        result[int(item["announcement"])].append({
            "path": rendered.resolve(),
            "from_pdf": is_pdf,
            "kind": item.get("kind", "poster"),
        })
    return result


def make_plan(manifest: dict, call_reference, call_verses, reading_reference, reading_pages, include_guide: bool) -> list[dict]:
    plan: list[dict] = [{"kind": "template", "slide": 1, "role": "cover"}]
    plan.append({"kind": "generated", "template_slide": 3, "role": "call_to_worship", "reference": call_reference.display})
    first, last = manifest["worship"]["first_three_range"]
    plan.extend({"kind": "worship", "slide": number, "role": "worship"} for number in range(first, last + 1))
    plan.append({"kind": "template", "slide": 7, "role": "welcome"})
    if include_guide:
        plan.append({"kind": "template", "slide": 8, "role": "scripture_guide"})
    for page in reading_pages:
        plan.append({"kind": "generated", "template_slide": 9, "role": "scripture", "reference": reference_for_page(reading_reference.book.traditional, page)})
    if manifest.get("sermon_pptx"):
        sermon_count = int(manifest["sermon_slide_count"])
        plan.extend({"kind": "sermon", "slide": number, "role": "sermon"} for number in range(1, sermon_count + 1))
    else:
        plan.append({"kind": "generated", "template_slide": 10, "role": "sermon_placeholder"})
    first, last = manifest["worship"]["response_range"]
    plan.extend({"kind": "worship", "slide": number, "role": "response"} for number in range(first, last + 1))
    plan.extend([
        {"kind": "template", "slide": 12, "role": "offering"},
        {"kind": "template", "slide": 13, "role": "welcome"},
    ])
    for number, text in enumerate(manifest["announcements"], start=1):
        pages = split_announcement(text)
        qr_assets = [
            item for item in manifest["assets"]
            if item["announcement"] == number and item.get("kind") == "qr"
        ]
        for page_number, page_text in enumerate(pages, start=1):
            plan.append({
                "kind": "generated",
                "template_slide": 14 if number == 1 and page_number == 1 else 15,
                "role": "announcement",
                "number": number,
                "page": page_number,
                "page_count": len(pages),
                "text": page_text,
                "qr_assets": [item["path"] for item in qr_assets] if page_number == len(pages) else [],
            })
        for asset in [
            item for item in manifest["assets"]
            if item["announcement"] == number and item.get("kind") != "qr"
        ]:
            plan.append({"kind": "asset", "role": "announcement_asset", "path": asset["path"]})
    plan.extend([
        {"kind": "template", "slide": 16, "role": "doxology"},
        {"kind": "template", "slide": 17, "role": "parking", "hidden": True},
    ])
    return plan


def assemble(
    manifest: dict,
    output: Path,
    include_guide: bool,
    dry_run: bool = False,
    engine: str = "auto",
) -> dict:
    started = time.perf_counter()
    timings: dict[str, float] = {}
    template = Path(manifest["template"])
    worship = Path(manifest["worship_pptx"])
    sermon = Path(manifest["sermon_pptx"]) if manifest.get("sermon_pptx") else None
    template_count = inventory(template)["slide_count"]

    phase = time.perf_counter()
    call_reference = parse_reference(manifest["scripture"]["call_to_worship"])
    reading_reference = parse_reference(manifest["scripture"]["reading"])
    call_verses = fetch_reference(call_reference)
    reading_verses = fetch_reference(reading_reference)
    reading_pages = chunk_reading(reading_verses)
    timings["scripture"] = time.perf_counter() - phase
    plan = make_plan(manifest, call_reference, call_verses, reading_reference, reading_pages, include_guide)
    selected_engine = select_engine(engine)

    result = {
        "manifest": manifest,
        "scripture": {
            "call_to_worship": {"reference": call_reference.display, "verses": [asdict(verse) for verse in call_verses]},
            "reading": {"reference": reading_reference.display, "verses": [asdict(verse) for verse in reading_verses]},
        },
        "plan": plan,
        "expected_slide_count": len(plan),
        "output": str(output.resolve()),
        "engine": selected_engine,
        "timings_seconds": timings,
    }
    if dry_run:
        timings["assemble_total"] = time.perf_counter() - started
        return result

    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        raise FileExistsError(f"Refusing to overwrite existing output: {output}")
    shutil.copy2(template, output)

    with tempfile.TemporaryDirectory(prefix=".sunday-ppt-build-", dir=output.parent) as temporary:
        phase = time.perf_counter()
        assets = prepare_assets(manifest, Path(temporary))
        timings["assets"] = time.perf_counter() - phase
        deck_class = AppleScriptDeck if selected_engine == "macos" else PowerShellDeck
        deck = deck_class(output.resolve(), template_count)
        repair_slides: list[tuple[Path, int, int]] = []
        native_slides: list[tuple[Path, int, int]] = []

        def append_template(slide_number: int, immutable: bool = False) -> int:
            output_number = deck.append_template(slide_number)
            if immutable:
                copied = (template, slide_number, output_number - template_count)
                repair_slides.append(copied)
                native_slides.append(copied)
            return output_number

        append_template(1, immutable=True)
        call_slide = append_template(3)
        call_text = "\n".join(verse.text for verse in call_verses) + f"\n({call_reference.display})"
        deck.set_text(call_slide, CALL_SHAPE, call_text, font_size=32)
        deck.center_shape_vertically(call_slide, CALL_SHAPE)

        first, last = manifest["worship"]["first_three_range"]
        for slide in range(first, last + 1):
            copied = (
                worship,
                slide,
                deck.append_source(worship, slide) - template_count,
            )
            repair_slides.append(copied)
            native_slides.append(copied)

        append_template(7, immutable=True)
        if include_guide:
            append_template(8)

        for page in reading_pages:
            slide = append_template(9)
            deck.set_text(
                slide,
                SCRIPTURE_TITLE_SHAPE,
                reading_reference.display,
                font_size=32,
            )
            deck.set_text(
                slide,
                SCRIPTURE_BODY_SHAPE,
                "\n".join(f"{verse.verse}. {verse.text}" for verse in page),
                font_size=40,
            )
            deck.remove_bullets(slide, SCRIPTURE_BODY_SHAPE)

        if sermon:
            sermon_count = int(manifest["sermon_slide_count"])
            for slide in range(1, sermon_count + 1):
                copied = (
                    sermon,
                    slide,
                    deck.append_source(sermon, slide) - template_count,
                )
                repair_slides.append(copied)
                native_slides.append(copied)
        else:
            deck.append_blank()

        first, last = manifest["worship"]["response_range"]
        for slide in range(first, last + 1):
            copied = (
                worship,
                slide,
                deck.append_source(worship, slide) - template_count,
            )
            repair_slides.append(copied)
            native_slides.append(copied)

        append_template(12, immutable=True)
        append_template(13, immutable=True)

        for number, announcement in enumerate(manifest["announcements"], start=1):
            pages = split_announcement(announcement)
            item_assets = assets.get(number, [])
            qr_assets = [item for item in item_assets if item["kind"] == "qr"]
            for page_number, page_text in enumerate(pages, start=1):
                use_first_layout = number == 1 and page_number == 1
                slide = append_template(14 if use_first_layout else 15)
                if not use_first_layout:
                    deck.set_text(slide, ANNOUNCEMENT_TITLE_SHAPE, "歡迎與報告")
                body_shape = "Google Shape;332;p14" if use_first_layout else ANNOUNCEMENT_BODY_SHAPE
                deck.set_text(slide, body_shape, page_text, font_size=36)
                deck.number_paragraphs(slide, body_shape, number)
                if page_number == len(pages):
                    for asset in qr_assets:
                        deck.add_picture(
                            slide,
                            asset["path"],
                            qr_geometry(asset["path"], page_text),
                        )
            for asset in item_assets:
                if asset["kind"] != "qr":
                    deck.add_image_slide(asset["path"], from_pdf=asset["from_pdf"])

        append_template(16, immutable=True)
        parking = append_template(17, immutable=True)
        deck.hide_slide(parking)
        script = deck.finish()
        suffix = ".applescript" if selected_engine == "macos" else ".ps1"
        script_path = Path(temporary) / f"build{suffix}"
        script_path.write_text(script, encoding="utf-8")
        if selected_engine == "macos":
            command = ["osascript", "-"]
            script_input = script
        else:
            executable = shutil.which("powershell.exe") or shutil.which("pwsh") or shutil.which("powershell")
            if not executable:
                raise RuntimeError("PowerShell is required for Windows PowerPoint automation")
            command = [executable, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script_path)]
            script_input = None
        phase = time.perf_counter()
        completed = subprocess.run(
            command,
            input=script_input,
            capture_output=True,
            text=True,
            check=False,
            timeout=900,
        )
        timings["powerpoint"] = time.perf_counter() - phase
        if completed.returncode != 0:
            failed_script = output.with_suffix(f".failed{suffix}")
            failed_script.write_text(script, encoding="utf-8")
            raise RuntimeError(
                f"PowerPoint automation failed: {completed.stderr.strip()}\n"
                f"Debug script: {failed_script}"
            )
        counts = re.findall(r"(?m)^\s*(\d+)\s*$", completed.stdout)
        if not counts:
            raise RuntimeError(f"PowerPoint did not report a slide count: {completed.stdout.strip()}")
        result["actual_slide_count"] = int(counts[-1])

    if result["actual_slide_count"] != result["expected_slide_count"]:
        raise RuntimeError(
            f"Expected {result['expected_slide_count']} slides, got {result['actual_slide_count']}"
        )
    source_inventories = {
        source: inventory(source)
        for source in dict.fromkeys(item[0] for item in native_slides)
    }
    phase = time.perf_counter()
    clipboard_inventory = inventory(output)
    result["repaired_media_effect_slides"] = repair_media_effect_slides(
        output,
        repair_slides,
        output_inventory=clipboard_inventory,
        source_inventories=source_inventories,
    )
    timings["media_repair"] = time.perf_counter() - phase
    phase = time.perf_counter()
    output_inventory = inventory(output)
    result["native_copy_validation"] = validate_native_copies(
        output,
        native_slides,
        output_inventory=output_inventory,
        source_inventories=source_inventories,
    )
    result["output_inventory"] = output_inventory
    timings["validation"] = time.perf_counter() - phase
    timings["assemble_total"] = time.perf_counter() - started
    return result


def main() -> None:
    main_started = time.perf_counter()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", nargs="?", default=".", type=Path)
    parser.add_argument("--template", type=Path)
    parser.add_argument("--worship", type=Path)
    parser.add_argument("--sermon", type=Path)
    parser.add_argument("--scripture-docx", type=Path)
    parser.add_argument("--bulletin-docx", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--include-scripture-guide", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--engine", choices=("auto", "macos", "windows"), default="auto")
    parser.add_argument("--json", type=Path, help="Write the complete plan and result as JSON")
    args = parser.parse_args()

    discovery_started = time.perf_counter()
    manifest = discover(
        args.directory,
        template=args.template,
        worship=args.worship,
        sermon=args.sermon,
        scripture_docx=args.scripture_docx,
        bulletin_docx=args.bulletin_docx,
    )
    discovery_elapsed = time.perf_counter() - discovery_started
    output = args.output or args.directory / manifest["output_name"]
    result = assemble(
        manifest,
        output.resolve(),
        args.include_scripture_guide,
        args.dry_run,
        engine=args.engine,
    )
    result["timings_seconds"]["discovery"] = discovery_elapsed
    result["timings_seconds"]["total"] = time.perf_counter() - main_started
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "output": result["output"],
        "expected_slide_count": result["expected_slide_count"],
        "actual_slide_count": result.get("actual_slide_count"),
        "engine": result["engine"],
        "timings_seconds": result["timings_seconds"],
        "plan_json": str(args.json.resolve()) if args.json else None,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise
