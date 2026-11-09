#!/usr/bin/env python3
"""Validate a generated Sunday deck against its sources and a finished reference deck."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from input_manifest import discover
from pptx_inventory import NS, inventory
from spring_bible import parse_reference


NATIVE_FIELDS = (
    "normalized_text",
    "shape_count",
    "picture_count",
    "has_transition",
    "has_timing",
)


def normalized(value: str) -> str:
    return re.sub(r"\s+", "", value).lower()


def native_difference(expected: dict, actual: dict) -> dict:
    differences = {
        field: {"expected": expected[field], "actual": actual[field]}
        for field in NATIVE_FIELDS
        if expected[field] != actual[field]
    }
    expected_media = sorted(item["sha256"] for item in expected["media"])
    actual_media = sorted(item["sha256"] for item in actual["media"])
    if expected_media != actual_media:
        differences["media_sha256"] = {
            "expected": expected_media,
            "actual": actual_media,
        }
    return differences


def find_native_sequence(target: dict, source_slides: list[dict]) -> int | None:
    count = len(source_slides)
    for index in range(len(target["slides"]) - count + 1):
        if all(
            not native_difference(source_slides[offset], target["slides"][index + offset])
            for offset in range(count)
        ):
            return index + 1
    return None


def reading_slides(deck: dict, title: str) -> list[dict]:
    marker = normalized(title)
    return [slide for slide in deck["slides"] if marker in slide["normalized_text"]]


def announcement_slides(deck: dict) -> list[dict]:
    marker = normalized("歡迎與報告")
    return [slide for slide in deck["slides"] if marker in slide["normalized_text"]]


def font_sizes(path: Path, slide: dict, shape_name: str) -> list[int]:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read(slide["part"]))
    for shape in root.findall(".//p:sp", NS):
        properties = shape.find("./p:nvSpPr/p:cNvPr", NS)
        if properties is None or properties.attrib.get("name") != shape_name:
            continue
        return sorted({
            int(node.attrib["sz"])
            for node in shape.findall(".//*[@sz]")
            if node.tag in {f"{{{NS['a']}}}rPr", f"{{{NS['a']}}}endParaRPr"}
        })
    return []


def shape_properties(path: Path, slide: dict, shape_name: str) -> dict:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read(slide["part"]))
    for shape in root.findall(".//p:sp", NS):
        properties = shape.find("./p:nvSpPr/p:cNvPr", NS)
        if properties is None or properties.attrib.get("name") != shape_name:
            continue
        transform = shape.find("./p:spPr/a:xfrm", NS)
        offset = transform.find("a:off", NS) if transform is not None else None
        extent = transform.find("a:ext", NS) if transform is not None else None
        return {
            "geometry": (
                int(offset.attrib["x"]),
                int(offset.attrib["y"]),
                int(extent.attrib["cx"]),
                int(extent.attrib["cy"]),
            ) if offset is not None and extent is not None else None,
            "has_auto_numbering": shape.find(".//a:buAutoNum", NS) is not None,
        }
    return {"geometry": None, "has_auto_numbering": False}


def asset_slide(deck: dict, digest: str) -> dict | None:
    matches = [
        slide for slide in deck["slides"]
        if digest in {item["sha256"] for item in slide["media"]}
    ]
    return matches[0] if len(matches) == 1 else None


def validate(directory: Path, generated_path: Path, reference_path: Path) -> dict:
    manifest = discover(directory)
    generated = inventory(generated_path)
    reference = inventory(reference_path)
    worship = inventory(Path(manifest["worship_pptx"]))
    template = inventory(Path(manifest["template"]))
    failures: list[dict] = []

    def check(name: str, condition: bool, details: object = None) -> None:
        if not condition:
            failures.append({"check": name, "details": details})

    first, last = manifest["worship"]["first_three_range"]
    worship_source = worship["slides"][first - 1 : last]
    response_first, response_last = manifest["worship"]["response_range"]
    response_source = worship["slides"][response_first - 1 : response_last]
    positions = {}
    for label, deck in (("generated", generated), ("reference", reference)):
        worship_start = find_native_sequence(deck, worship_source)
        response_start = find_native_sequence(deck, response_source)
        positions[label] = {
            "worship_start": worship_start,
            "response_start": response_start,
        }
        check(f"{label}_native_worship", worship_start is not None)
        check(f"{label}_native_response", response_start is not None)

    reading_title = parse_reference(manifest["scripture"]["reading"]).display
    generated_reading = reading_slides(generated, reading_title)
    reference_reading = reading_slides(reference, reading_title)
    check(
        "reading_pages_match_reference",
        [slide["normalized_text"] for slide in generated_reading]
        == [slide["normalized_text"] for slide in reference_reading],
        {
            "generated": [slide["number"] for slide in generated_reading],
            "reference": [slide["number"] for slide in reference_reading],
        },
    )
    for slide, reference_slide in zip(generated_reading, reference_reading):
        check(
            f"scripture_title_32pt_slide_{slide['number']}",
            font_sizes(generated_path, slide, "Google Shape;297;p9") == [3200],
        )
        check(
            f"scripture_body_font_matches_reference_slide_{slide['number']}",
            font_sizes(generated_path, slide, "Google Shape;298;p9")
            in (
                [4000],
                font_sizes(reference_path, reference_slide, "Google Shape;298;p9"),
            ),
        )
        check(
            f"scripture_explicit_verse_numbers_slide_{slide['number']}",
            not shape_properties(
                generated_path,
                slide,
                "Google Shape;298;p9",
            )["has_auto_numbering"],
        )

    generated_call = generated["slides"][1]
    reference_call = reference["slides"][1]
    generated_call_properties = shape_properties(
        generated_path,
        generated_call,
        "Google Shape;261;p3",
    )
    reference_call_properties = shape_properties(
        reference_path,
        reference_call,
        "Google Shape;261;p3",
    )
    check(
        "call_to_worship_matches_reference",
        generated_call["normalized_text"] == reference_call["normalized_text"],
    )
    check(
        "call_to_worship_32pt",
        font_sizes(generated_path, generated_call, "Google Shape;261;p3") == [3200],
    )
    generated_geometry = generated_call_properties["geometry"]
    reference_geometry = reference_call_properties["geometry"]
    check(
        "call_to_worship_geometry_matches_reference",
        generated_geometry is not None
        and reference_geometry is not None
        and all(abs(left - right) <= 2000 for left, right in zip(generated_geometry, reference_geometry)),
        {"generated": generated_geometry, "reference": reference_geometry},
    )

    generated_announcements = announcement_slides(generated)
    reference_announcements = announcement_slides(reference)
    check(
        "announcement_text_matches_reference",
        [slide["normalized_text"] for slide in generated_announcements]
        == [slide["normalized_text"] for slide in reference_announcements],
        {
            "generated_count": len(generated_announcements),
            "reference_count": len(reference_announcements),
        },
    )

    qr_results = []
    for asset in manifest["assets"]:
        digest = hashlib.sha256(Path(asset["path"]).read_bytes()).hexdigest()
        expected_text = normalized(manifest["announcements"][asset["announcement"] - 1])
        generated_asset_slide = asset_slide(generated, digest)
        reference_asset_slide = asset_slide(reference, digest)
        qr_results.append({
            "asset": Path(asset["path"]).name,
            "kind": asset["kind"],
            "generated_slide": generated_asset_slide["number"] if generated_asset_slide else None,
            "reference_slide": reference_asset_slide["number"] if reference_asset_slide else None,
        })
        if asset["kind"] == "qr":
            check(
                f"qr_with_announcement_{Path(asset['path']).name}",
                generated_asset_slide is not None
                and expected_text in generated_asset_slide["normalized_text"],
            )
        else:
            check(
                f"poster_on_separate_slide_{Path(asset['path']).name}",
                generated_asset_slide is not None and not generated_asset_slide["normalized_text"],
            )

    template_doxology = template["slides"][15]
    for label, deck in (("generated", generated), ("reference", reference)):
        doxology = deck["slides"][-2]
        check(
            f"{label}_doxology_native",
            not native_difference(template_doxology, doxology),
            native_difference(template_doxology, doxology),
        )
        check(f"{label}_parking_hidden", deck["slides"][-1]["hidden"])

    if not manifest.get("sermon_pptx"):
        generated_response_start = positions["generated"]["response_start"]
        if generated_reading and generated_response_start:
            sermon_span = generated["slides"][generated_reading[-1]["number"] : generated_response_start - 1]
            check(
                "single_blank_sermon_placeholder",
                len(sermon_span) == 1
                and sermon_span[0]["shape_count"] == 0
                and not sermon_span[0]["normalized_text"],
                [slide["number"] for slide in sermon_span],
            )

    return {
        "generated": str(generated_path.resolve()),
        "reference": str(reference_path.resolve()),
        "positions": positions,
        "reading_slides": {
            "generated": [slide["number"] for slide in generated_reading],
            "reference": [slide["number"] for slide in reference_reading],
        },
        "qr_assets": qr_results,
        "passed": not failures,
        "failures": failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    parser.add_argument("generated", type=Path)
    parser.add_argument("reference", type=Path)
    args = parser.parse_args()
    result = validate(args.directory.resolve(), args.generated.resolve(), args.reference.resolve())
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if not result["passed"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
