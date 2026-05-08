
#!/usr/bin/env python3
"""
把「新聖詩歌詞」.docx 轉成 ProPresenter 可逐首匯入的 .txt 檔案。

使用方式：
  1. 把此檔案放在含有「新聖詩歌詞1-300.docx」與「新聖詩歌詞301-650.docx」的資料夾。
  2. 在終端機執行：
       python3 convert_hymns_to_propresenter_txt.py
  3. 轉出的檔案會在「ProPresenter_單首TXT」資料夾。

ProPresenter 匯入文字檔時，通常會用「空白行」分隔投影片；
此程式預設會把每一節歌詞做成一張投影片。
"""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from zipfile import ZipFile
from xml.etree import ElementTree as ET


# WordprocessingML namespace
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

DEFAULT_INPUTS = ["新聖詩歌詞1-300.docx", "新聖詩歌詞301-650.docx"]
DEFAULT_OUTPUT = "ProPresenter_單首TXT"


@dataclass
class Song:
    number: int
    title: str
    lines: list[str]

    @property
    def display_title(self) -> str:
        return f"新聖詩{self.number}首"


def normalize_text(text: str) -> str:
    """Normalize spaces and common punctuation without changing Chinese/Taiwanese text."""
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def paragraph_text(p: ET.Element) -> str:
    """Extract visible text from one Word paragraph, including manual line breaks."""
    parts: list[str] = []
    for node in p.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "t":
            parts.append(node.text or "")
        elif tag == "tab":
            parts.append(" ")
        elif tag in {"br", "cr"}:
            parts.append("\n")
    return normalize_text("".join(parts))


def extract_docx_lines(docx_path: Path) -> list[str]:
    """Read paragraphs from a .docx using only Python standard library."""
    with ZipFile(docx_path) as zf:
        xml = zf.read("word/document.xml")

    root = ET.fromstring(xml)
    lines: list[str] = []

    # Reading all paragraphs also captures paragraphs inside Word tables.
    for p in root.findall(".//w:p", NS):
        text = paragraph_text(p)
        if not text:
            continue
        for line in text.splitlines():
            line = normalize_text(line)
            if line:
                lines.append(line)

    return lines


def looks_like_metadata(line: str) -> bool:
    """Skip document-level headers/footers that are not song text."""
    bad_patterns = [
        r"^新聖詩歌詞",
        r"^目錄$",
        r"^頁\s*\d+$",
        r"^\d+\s*/\s*\d+$",
    ]
    return any(re.search(pattern, line) for pattern in bad_patterns)


def parse_song_heading(line: str) -> tuple[int, str] | None:
    """
    Try to detect a hymn heading.

    Accepted examples:
      001 主上帝創造天地
            001 131 1.聖哉,聖哉,聖哉3,

        The source Word files use three-digit hymn numbers at the start of each song.
        Do not treat stanza lyrics like "2.聖哉..." as a new song.
    """
    line = normalize_text(line)

    patterns = [
        r"^(\d{3})#?(?:\s+(.+))?$",
    ]
    for pattern in patterns:
        match = re.match(pattern, line)
        if not match:
            continue
        number = int(match.group(1))
        title = normalize_text(match.group(2) or "")
        # Remove the old hymn number column when present, e.g.
        # "131 1.聖哉..." -> "1.聖哉..." and "62A 1.上帝..." -> "1.上帝...".
        old_number = re.match(r"^(\d{1,3}[A-Z]?)(?:\s+(.+))?$", title, re.IGNORECASE)
        if old_number:
            title = normalize_text(old_number.group(2) or "")
        if 1 <= number <= 999:
            return number, title
    return None


def split_songs(lines: Iterable[str]) -> list[Song]:
    songs: list[Song] = []
    current: Song | None = None
    previous_number = 0

    for raw_line in lines:
        line = normalize_text(raw_line)
        if not line or looks_like_metadata(line):
            continue

        heading = parse_song_heading(line)
        if heading:
            number, title = heading
            if number > previous_number:
                current = Song(number=number, title=title, lines=[title] if title else [])
                songs.append(current)
                previous_number = number
                continue
            if previous_number == 0:
                current = Song(number=number, title=title, lines=[title] if title else [])
                songs.append(current)
                previous_number = number
                continue

        if current is not None:
            current.lines.append(line)

    return songs


def is_stanza_marker(line: str) -> bool:
    """Detect stanza labels that should start a new slide."""
    return bool(
        re.match(r"^(?:\(?\d{1,2}\)?|[一二三四五六七八九十]+)\s*[.．、:：]?\s*$", line)
        or re.match(r"^(?:第\s*)?[一二三四五六七八九十\d]{1,3}\s*(?:節|段)\s*[:：]?$", line)
        or re.match(r"^(?:副歌|和|阿們|Amen)\s*[:：]?$", line, re.IGNORECASE)
    )


def starts_with_stanza_number(line: str) -> bool:
    """Detect lyrics like '1. 主上帝...' where stanza marker and lyric are on same line."""
    return bool(re.match(r"^(?:\(?\d{1,2}\)?|[一二三四五六七八九十]+)\s*[.．、:：]\s*\S+", line))


def make_slides(song: Song, max_lines_per_slide: int = 6) -> list[list[str]]:
    """Split one song into slides. Blank line between slides is ProPresenter-friendly."""
    slides: list[list[str]] = []
    current: list[str] = []

    def flush() -> None:
        nonlocal current
        if current:
            slides.append(current)
            current = []

    for line in song.lines:
        line = normalize_text(line)
        if not line:
            flush()
            continue

        if is_stanza_marker(line):
            flush()
            current.append(line)
            continue

        if starts_with_stanza_number(line):
            flush()
            current.append(line)
            continue

        current.append(line)
        if max_lines_per_slide > 0 and len(current) >= max_lines_per_slide:
            flush()

    flush()

    if not slides and song.lines:
        slides = [song.lines]
    return slides


def safe_filename(name: str) -> str:
    name = normalize_text(name)
    name = re.sub(r"[\\/:*?\"<>|]", "-", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:120]


def write_song_txt(song: Song, output_dir: Path, max_lines_per_slide: int, overwrite: bool) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = safe_filename(f"新聖詩{song.number}首.txt")
    out_path = output_dir / filename

    if out_path.exists() and not overwrite:
        raise FileExistsError(f"檔案已存在，請加 --overwrite 覆蓋：{out_path}")

    slides = make_slides(song, max_lines_per_slide=max_lines_per_slide)

    # 第一行是歌名；之後用空白行分隔每張投影片。
    blocks = [song.display_title]
    blocks.extend("\n".join(slide) for slide in slides if slide)
    text = "\n\n".join(blocks).strip() + "\n"
    out_path.write_text(text, encoding="utf-8")
    return out_path


def convert(inputs: list[Path], output_dir: Path, max_lines_per_slide: int, overwrite: bool) -> list[Path]:
    all_songs: list[Song] = []

    for docx_path in inputs:
        if not docx_path.exists():
            print(f"找不到檔案，略過：{docx_path}", file=sys.stderr)
            continue
        lines = extract_docx_lines(docx_path)
        songs = split_songs(lines)
        print(f"{docx_path.name}: 找到 {len(songs)} 首")
        all_songs.extend(songs)

    # 若兩個 Word 檔有重複編號，保留後面檔案最後讀到的版本。
    by_number: dict[int, Song] = {song.number: song for song in all_songs}
    songs_sorted = [by_number[number] for number in sorted(by_number)]

    written: list[Path] = []
    for song in songs_sorted:
        written.append(write_song_txt(song, output_dir, max_lines_per_slide, overwrite))

    return written


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="把新聖詩 Word 歌詞檔轉成 ProPresenter 可逐首匯入的 UTF-8 txt。"
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        type=Path,
        default=[Path(name) for name in DEFAULT_INPUTS],
        help="要轉換的 .docx 檔；預設為新聖詩歌詞1-300.docx 與 新聖詩歌詞301-650.docx。",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path(DEFAULT_OUTPUT),
        help=f"輸出資料夾；預設：{DEFAULT_OUTPUT}",
    )
    parser.add_argument(
        "--max-lines-per-slide",
        type=int,
        default=6,
        help="每張投影片最多幾行；0 表示只照節數分頁，不再按行數切。預設：6",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="覆蓋已存在的 txt 檔。",
    )
    return parser


def main() -> int:
    parser = build_arg_parser()
    args = parser.parse_args()

    written = convert(
        inputs=args.inputs,
        output_dir=args.output,
        max_lines_per_slide=args.max_lines_per_slide,
        overwrite=args.overwrite,
    )

    print(f"完成：輸出 {len(written)} 個 txt 檔到 {args.output.resolve()}")
    if written:
        print(f"範例：{written[0].name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
