#!/usr/bin/env python3
"""Convert 新聖詩 DOCX/PPTX files into ProPresenter-importable TXT files.

Each output file is named `新聖詩N首.txt`, and the first line inside the
file is the same title without any leading zero before the hymn number.
Slides are separated by blank lines for ProPresenter text import.
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

HYMNV_RE = re.compile(r"新聖詩\s*0*(\d+)\s*首")
FOOTER_RE = re.compile(r"新聖詩\s*\d+\s*首")
SLIDE_RE = re.compile(r"slide(\d+)\.xml$")
DOCX_HEADING_RE = re.compile(r"^(\d{3})(.*)$")
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


@dataclass(frozen=True)
class TextBox:
    index: int
    y: int
    paragraphs: list[str]


@dataclass
class Song:
    number: int
    lines: list[str]

    @property
    def title(self) -> str:
        return f"新聖詩{self.number}首"


def normalize_line(text: str) -> str:
    """Normalize whitespace while preserving hymn text."""
    return re.sub(r"[ \t\u3000]+", " ", text).strip()


def is_metadata_line(text: str) -> bool:
    return not text or text in {"新舊歌詞", "新舊歌 詞", "新舊歌  詞"}


def hymn_title(path: Path) -> str:
    match = HYMNV_RE.search(path.stem)
    if not match:
        raise ValueError(f"Cannot find hymn number in file name: {path.name}")
    return f"新聖詩{int(match.group(1))}首"


def read_docx_lines(docx_path: Path) -> list[str]:
    """Extract visible paragraph text from a DOCX file using the standard library."""
    with zipfile.ZipFile(docx_path) as docx:
        root = ET.fromstring(docx.read("word/document.xml"))

    lines: list[str] = []
    for para in root.findall(".//w:p", W_NS):
        text = "".join(node.text or "" for node in para.findall(".//w:t", W_NS))
        text = normalize_line(text)
        if text and not is_metadata_line(text):
            lines.append(text)
    return lines


def looks_like_lyric_start(text: str) -> bool:
    return bool(
        re.match(r"^1\s*(?:[.．、:：-]|$)", text)
        or re.match(r"^(?:台語|華語|客語)\b", text)
        or re.match(r"^[^0-9]", text)
    )


def strip_old_hymn_number(text: str) -> str:
    """Remove the old-hymnal number column from a DOCX heading remainder."""
    text = text.lstrip("# ")
    candidates: list[tuple[int, str]] = []
    for digit_count in range(1, min(3, len(text)) + 1):
        token = text[:digit_count]
        if not token.isdigit():
            break
        end = digit_count
        if end < len(text) and text[end].isalpha() and text[end].isascii():
            end += 1
        remainder = text[end:].strip()
        if remainder and looks_like_lyric_start(remainder):
            candidates.append((end, remainder))

    if candidates:
        return max(candidates, key=lambda item: item[0])[1]
    return text.strip()


def parse_docx_heading(line: str) -> tuple[int, str] | None:
    match = DOCX_HEADING_RE.match(line)
    if not match:
        return None
    number = int(match.group(1))
    if not 1 <= number <= 650:
        return None
    first_line = strip_old_hymn_number(match.group(2))
    return number, first_line


def load_docx_songs(docs_dir: Path) -> list[Song]:
    docx_files = sorted(docs_dir.glob("*.docx"))
    if not docx_files:
        raise FileNotFoundError(f"No .docx files found in {docs_dir}")

    songs: dict[int, Song] = {}
    current: Song | None = None
    for docx_path in docx_files:
        for line in read_docx_lines(docx_path):
            heading = parse_docx_heading(line)
            if heading:
                number, first_line = heading
                current = songs.setdefault(number, Song(number=number, lines=[]))
                if first_line:
                    current.lines.append(first_line)
                continue
            if current is not None:
                current.lines.append(line)

    return [songs[number] for number in sorted(songs)]


def starts_new_docx_block(line: str) -> bool:
    return bool(
        re.match(r"^(?:\d{1,2}(?:-\d+)?\s*[.．、:：]|【(?:複歌|副歌|尾聲)|(?:複歌|副歌))", line)
        or re.match(r"^(?:台語|華語|客語)\s*\d{0,2}\s*[.．、:：]", line)
    )


def docx_blocks(lines: list[str]) -> list[str]:
    blocks: list[str] = []
    current: list[str] = []
    for line in lines:
        line = normalize_line(line)
        if not line:
            continue
        if current and starts_new_docx_block(line):
            blocks.append("\n".join(current))
            current = []
        current.append(line)
    if current:
        blocks.append("\n".join(current))
    return blocks


def write_song_txt(song: Song, output_dir: Path, overwrite: bool) -> Path:
    output_path = output_dir / f"{song.title}.txt"
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"Output exists; use --overwrite: {output_path}")

    content = song.title + "\n\n"
    blocks = docx_blocks(song.lines)
    if blocks:
        content += "\n\n".join(blocks).rstrip() + "\n"
    output_path.write_text(content, encoding="utf-8")
    return output_path


def slide_number(path_in_zip: str) -> int:
    match = SLIDE_RE.search(Path(path_in_zip).name)
    if not match:
        return 0
    return int(match.group(1))


def textbox_y(shape: ET.Element) -> int:
    off = shape.find(".//a:off", NS)
    if off is None:
        return 0
    try:
        return int(off.attrib.get("y", "0"))
    except ValueError:
        return 0


def shape_paragraphs(shape: ET.Element) -> list[str]:
    paragraphs: list[str] = []
    text_body = shape.find("p:txBody", NS)
    if text_body is None:
        return paragraphs

    for para in text_body.findall("a:p", NS):
        text = "".join(node.text or "" for node in para.findall(".//a:t", NS))
        text = normalize_line(text)
        if not text:
            continue
        # PPT footers are usually appended as their own paragraph, for example:
        # 新聖詩101首（1/3）舊131首
        if FOOTER_RE.search(text):
            continue
        paragraphs.append(text)
    return paragraphs


def slide_texts(pptx_path: Path) -> list[str]:
    blocks: list[str] = []
    with zipfile.ZipFile(pptx_path) as deck:
        slide_paths = sorted(
            (
                name
                for name in deck.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            ),
            key=slide_number,
        )

        for slide_path in slide_paths:
            root = ET.fromstring(deck.read(slide_path))
            boxes: list[TextBox] = []
            for index, shape in enumerate(root.findall(".//p:sp", NS)):
                paragraphs = shape_paragraphs(shape)
                if paragraphs:
                    boxes.append(TextBox(index=index, y=textbox_y(shape), paragraphs=paragraphs))

            # Sort primarily by vertical position so verse labels placed above the
            # main lyrics are emitted before lyrics; keep original shape order for
            # text boxes on the same horizontal row.
            boxes.sort(key=lambda box: (box.y, box.index))
            lines = [line for box in boxes for line in box.paragraphs]
            if lines:
                blocks.append("\n".join(lines))
    return blocks


def convert_one_pptx(pptx_path: Path, output_dir: Path, overwrite: bool) -> Path:
    title = hymn_title(pptx_path)
    output_path = output_dir / f"{title}.txt"
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"Output exists; use --overwrite: {output_path}")

    blocks = slide_texts(pptx_path)
    content = title + "\n\n"
    if blocks:
        content += "\n\n".join(blocks).rstrip() + "\n"

    output_path.write_text(content, encoding="utf-8")
    return output_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--source",
        choices=("docs", "ppt"),
        default="docs",
        help="Use complete DOCX lyrics from raw_data/新聖詩/Docs, or PPTX files from raw_data/新聖詩/PPT.",
    )
    parser.add_argument(
        "--docs-dir",
        type=Path,
        default=Path("raw_data/新聖詩/Docs"),
        help="Directory containing complete 新聖詩 .docx lyric files.",
    )
    parser.add_argument(
        "--ppt-dir",
        type=Path,
        default=Path("raw_data/新聖詩/PPT"),
        help="Directory containing 新聖詩 .pptx files.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("transform_data/新聖詩"),
        help="Directory where .txt files will be written.",
    )
    parser.add_argument(
        "--clean-output",
        action="store_true",
        help="Delete existing 新聖詩*.txt files in the output directory before converting.",
    )
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing .txt files.")
    args = parser.parse_args(argv)

    output_dir = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)
    if args.clean_output:
        for existing_path in output_dir.glob("新聖詩*首.txt"):
            existing_path.unlink()

    converted: list[Path] = []
    if args.source == "docs":
        songs = load_docx_songs(args.docs_dir)
        for song in songs:
            converted.append(write_song_txt(song, output_dir, args.overwrite))
        print(f"Converted {len(converted)} DOCX songs to {output_dir}")
        return 0

    pptx_files = sorted(args.ppt_dir.glob("*.pptx"), key=lambda path: hymn_title(path))
    if not pptx_files:
        print(f"No .pptx files found in {args.ppt_dir}", file=sys.stderr)
        return 1

    for pptx_path in pptx_files:
        converted.append(convert_one_pptx(pptx_path, output_dir, args.overwrite))

    print(f"Converted {len(converted)} PPTX files to {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
