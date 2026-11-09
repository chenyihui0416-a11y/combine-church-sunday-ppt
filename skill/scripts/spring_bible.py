#!/usr/bin/env python3
"""Fetch Traditional Chinese Union Version verses from springbible.fhl.net."""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import os
import re
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass
from pathlib import Path


@dataclass(frozen=True)
class Book:
    traditional: str
    api_name: str
    aliases: tuple[str, ...]


@dataclass(frozen=True)
class Reference:
    book: Book
    start_chapter: int
    start_verse: int
    end_chapter: int
    end_verse: int

    @property
    def display(self) -> str:
        if self.start_chapter == self.end_chapter:
            span = f"{self.start_chapter}:{self.start_verse}-{self.end_verse}"
        else:
            span = (
                f"{self.start_chapter}:{self.start_verse}-"
                f"{self.end_chapter}:{self.end_verse}"
            )
        return f"{self.book.traditional} {span}"


@dataclass(frozen=True)
class Verse:
    chapter: int
    verse: int
    text: str


def _book(traditional: str, api_name: str, *aliases: str) -> Book:
    base = traditional.replace(" ", "")
    return Book(traditional, api_name, tuple(dict.fromkeys((base, api_name, *aliases))))


BOOKS = (
    _book("創世記", "創", "创世记", "創世", "创世"),
    _book("出埃及記", "出", "出埃及记"),
    _book("利未記", "利", "利未记"),
    _book("民數記", "民", "民数记"),
    _book("申命記", "申", "申命记"),
    _book("約書亞記", "書", "约书亚记", "書", "书"),
    _book("士師記", "士", "士师记"),
    _book("路得記", "得", "路得记"),
    _book("撒母耳記上", "撒上", "撒母耳记上"),
    _book("撒母耳記下", "撒下", "撒母耳记下"),
    _book("列王紀上", "王上", "列王纪上"),
    _book("列王紀下", "王下", "列王纪下"),
    _book("歷代志上", "代上", "历代志上"),
    _book("歷代志下", "代下", "历代志下"),
    _book("以斯拉記", "拉", "以斯拉记"),
    _book("尼希米記", "尼", "尼希米记"),
    _book("以斯帖記", "斯", "以斯帖记"),
    _book("約伯記", "伯", "约伯记"),
    _book("詩篇", "詩", "诗篇", "詩", "诗"),
    _book("箴言", "箴"),
    _book("傳道書", "傳", "传道书", "传"),
    _book("雅歌", "歌"),
    _book("以賽亞書", "賽", "以赛亚书", "赛"),
    _book("耶利米書", "耶", "耶利米书"),
    _book("耶利米哀歌", "哀"),
    _book("以西結書", "結", "以西结书", "结"),
    _book("但以理書", "但", "但以理书"),
    _book("何西阿書", "何", "何西阿书"),
    _book("約珥書", "珥", "约珥书"),
    _book("阿摩司書", "摩", "阿摩司书"),
    _book("俄巴底亞書", "俄", "俄巴底亚书"),
    _book("約拿書", "拿", "约拿书"),
    _book("彌迦書", "彌", "弥迦书", "弥"),
    _book("那鴻書", "鴻", "那鸿书", "鸿"),
    _book("哈巴谷書", "哈", "哈巴谷书"),
    _book("西番雅書", "番", "西番雅书"),
    _book("哈該書", "該", "哈该书", "该"),
    _book("撒迦利亞書", "亞", "撒迦利亚书", "亚"),
    _book("瑪拉基書", "瑪", "玛拉基书", "玛"),
    _book("馬太福音", "太", "马太福音", "馬太", "马太"),
    _book("馬可福音", "可", "马可福音", "馬可", "马可"),
    _book("路加福音", "路", "路加"),
    _book("約翰福音", "約", "约翰福音", "約翰", "约翰", "约"),
    _book("使徒行傳", "徒", "使徒行传", "使徒"),
    _book("羅馬書", "羅", "罗马书", "羅馬", "罗马", "罗"),
    _book("哥林多前書", "林前", "哥林多前书"),
    _book("哥林多後書", "林後", "哥林多后书", "林后"),
    _book("加拉太書", "加", "加拉太书"),
    _book("以弗所書", "弗", "以弗所书"),
    _book("腓立比書", "腓", "腓立比书"),
    _book("歌羅西書", "西", "歌罗西书"),
    _book("帖撒羅尼迦前書", "帖前", "帖撒罗尼迦前书"),
    _book("帖撒羅尼迦後書", "帖後", "帖撒罗尼迦后书", "帖后"),
    _book("提摩太前書", "提前", "提摩太前书"),
    _book("提摩太後書", "提後", "提摩太后书", "提后"),
    _book("提多書", "多", "提多书"),
    _book("腓利門書", "門", "腓利门书", "门"),
    _book("希伯來書", "來", "希伯来书", "来"),
    _book("雅各書", "雅", "雅各书"),
    _book("彼得前書", "彼前", "彼得前书"),
    _book("彼得後書", "彼後", "彼得后书", "彼后"),
    _book("約翰一書", "約一", "约翰一书", "约一"),
    _book("約翰二書", "約二", "约翰二书", "约二"),
    _book("約翰三書", "約三", "约翰三书", "约三"),
    _book("猶大書", "猶", "犹大书", "犹"),
    _book("啟示錄", "啟", "启示录", "啓示錄", "启"),
)


def parse_reference(raw: str) -> Reference:
    compact = re.sub(r"\s+", "", raw).replace("：", ":").replace("－", "-").replace("–", "-")
    aliases = sorted(
        ((alias, book) for book in BOOKS for alias in book.aliases),
        key=lambda item: len(item[0]),
        reverse=True,
    )
    book = None
    span = compact
    for alias, candidate in aliases:
        if compact.startswith(alias):
            book = candidate
            span = compact[len(alias) :]
            break
    if book is None:
        raise ValueError(f"Unsupported Bible book in reference: {raw}")
    match = re.fullmatch(r"(\d+):(\d+)(?:-(?:(\d+):)?(\d+))?", span)
    if not match:
        raise ValueError(f"Unsupported scripture range: {raw}")
    start_chapter = int(match.group(1))
    start_verse = int(match.group(2))
    end_chapter = int(match.group(3) or start_chapter)
    end_verse = int(match.group(4) or start_verse)
    if (end_chapter, end_verse) < (start_chapter, start_verse):
        raise ValueError(f"Range ends before it starts: {raw}")
    return Reference(book, start_chapter, start_verse, end_chapter, end_verse)


def _strip_tags(value: str) -> str:
    value = re.sub(r"<p\s*/?>", " ", value, flags=re.I)
    value = re.sub(r"<[^>]+>", "", value)
    return re.sub(r"\s+", " ", html.unescape(value)).strip()


def scripture_cache_dir() -> Path:
    override = os.environ.get("SUNDAY_PPT_CACHE_DIR")
    if override:
        return Path(override).expanduser() / "scripture"
    if os.name == "nt":
        root = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        root = Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache"))
    return root / "sunday-ppt-skill" / "scripture"


def fetch_chapter(book: Book, chapter: int, timeout: int = 20, refresh: bool = False) -> dict[int, str]:
    cache_key = hashlib.sha256(f"{book.api_name}:{chapter}".encode("utf-8")).hexdigest()[:16]
    cache_file = scripture_cache_dir() / f"{cache_key}.json"
    if not refresh and cache_file.exists():
        try:
            cached = json.loads(cache_file.read_text(encoding="utf-8"))
            verses = {int(number): text for number, text in cached["verses"].items()}
            if verses:
                return verses
        except (OSError, ValueError, KeyError, TypeError):
            pass

    encoded_book = urllib.parse.quote_from_bytes(book.api_name.encode("big5"))
    url = (
        "https://springbible.fhl.net/Bible2/cgic201/read201.cgi"
        f"?na={encoded_book}&chap={chapter:03d}&ver=big5&ft=0"
    )
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        source = response.read().decode("big5", errors="replace")
    verses: dict[int, str] = {}
    pattern = re.compile(r'<li\s+value="(\d+)"[^>]*>(.*?)</li>', re.I | re.S)
    for match in pattern.finditer(source):
        verses[int(match.group(1))] = _strip_tags(match.group(2))
    if not verses:
        raise RuntimeError(f"No verses found at {url}")
    try:
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        temporary = cache_file.with_suffix(".tmp")
        temporary.write_text(
            json.dumps({"book": book.traditional, "chapter": chapter, "verses": verses}, ensure_ascii=False),
            encoding="utf-8",
        )
        temporary.replace(cache_file)
    except OSError:
        pass
    return verses


def fetch_reference(reference: Reference) -> list[Verse]:
    result: list[Verse] = []
    for chapter in range(reference.start_chapter, reference.end_chapter + 1):
        chapter_verses = fetch_chapter(reference.book, chapter)
        start = reference.start_verse if chapter == reference.start_chapter else 1
        end = reference.end_verse if chapter == reference.end_chapter else max(chapter_verses)
        missing = [verse for verse in range(start, end + 1) if verse not in chapter_verses]
        if missing:
            raise RuntimeError(
                f"Missing {reference.book.traditional} {chapter}:{missing} from springbible"
            )
        result.extend(Verse(chapter, verse, chapter_verses[verse]) for verse in range(start, end + 1))
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("reference")
    args = parser.parse_args()
    reference = parse_reference(args.reference)
    verses = fetch_reference(reference)
    print(json.dumps({
        "reference": reference.display,
        "book": asdict(reference.book),
        "verses": [asdict(verse) for verse in verses],
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
