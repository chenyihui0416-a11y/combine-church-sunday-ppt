#!/usr/bin/env python3
"""Install the Sunday PPT skill for Codex, Claude Code, or both."""

from __future__ import annotations

import argparse
import os
import shutil
from pathlib import Path


SKILL_NAME = "combine-church-sunday-ppt"
SOURCE = Path(__file__).resolve().parent / "skill"


def codex_home() -> Path:
    return Path(os.environ.get("CODEX_HOME", Path.home() / ".codex"))


def destinations(target: str, project: Path | None) -> list[tuple[str, Path]]:
    result: list[tuple[str, Path]] = []
    if target in {"codex", "both"}:
        result.append(("Codex", codex_home() / "skills" / SKILL_NAME))
    if target in {"claude", "both"}:
        root = project.resolve() / ".claude" if project else Path.home() / ".claude"
        result.append(("Claude Code", root / "skills" / SKILL_NAME))
    return result


def install(destination: Path, force: bool) -> None:
    bundled_decks = list(SOURCE.rglob("*.pptx"))
    if bundled_decks:
        names = ", ".join(str(path.relative_to(SOURCE)) for path in bundled_decks)
        raise RuntimeError(f"Refusing to install bundled PPTX files: {names}")
    if destination.exists():
        if not force:
            raise FileExistsError(f"{destination} already exists; rerun with --force to update it")
        shutil.rmtree(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(
        SOURCE,
        destination,
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store"),
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--target", choices=("codex", "claude", "both"), default="both")
    parser.add_argument(
        "--project",
        type=Path,
        help="Install Claude Code skill into PROJECT/.claude/skills instead of ~/.claude/skills",
    )
    parser.add_argument("--force", action="store_true", help="Replace an existing installation")
    args = parser.parse_args()

    if not SOURCE.joinpath("SKILL.md").exists():
        raise FileNotFoundError(f"Missing bundled skill at {SOURCE}")
    for label, destination in destinations(args.target, args.project):
        install(destination, args.force)
        print(f"Installed for {label}: {destination}")


if __name__ == "__main__":
    main()
