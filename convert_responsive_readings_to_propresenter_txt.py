#!/usr/bin/env python3
"""Convert 啟應文 PPTX files into ProPresenter-importable TXT files.

Each output file is named `啟應文N.txt`, and the first line inside the file is
`啟應文N` without any leading zero. Slides are separated by blank lines for
ProPresenter text import.
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

READING_RE = re.compile(r"新\s*0*(\d+)")
HEADER_RE = re.compile(r"啟應文\s*\d+\s*(?:[（(]\s*\d+\s*/\s*\d+\s*[）)])?")
SLIDE_RE = re.compile(r"slide(\d+)\.xml$")


@dataclass(frozen=True)
class TextBox:
    index: int
    y: int
    paragraphs: list[str]


def normalize_line(text: str) -> str:
    """Normalize whitespace while preserving Taiwanese text."""
    return re.sub(r"[ \t\u3000]+", " ", text).strip()


def reading_title(path: Path) -> str:
    match = READING_RE.search(path.stem)
    if not match:
        raise ValueError(f"Cannot find reading number in file name: {path.name}")
    return f"啟應文{int(match.group(1))}"


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
        # Remove slide headers like: 啟應文1(1/4)
        if HEADER_RE.fullmatch(text):
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

            boxes.sort(key=lambda box: (box.y, box.index))
            lines = [line for box in boxes for line in box.paragraphs]
            if lines:
                blocks.append("\n".join(lines))
    return blocks


def convert_one(pptx_path: Path, output_dir: Path, overwrite: bool) -> Path:
    title = reading_title(pptx_path)
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
        "--input-dir",
        type=Path,
        default=Path("raw_data/啟應文"),
        help="Directory containing 啟應文 .pptx files.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("transform_data/啟應文"),
        help="Directory where .txt files will be written.",
    )
    parser.add_argument(
        "--clean-output",
        action="store_true",
        help="Delete existing 啟應文*.txt files in the output directory before converting.",
    )
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing .txt files.")
    args = parser.parse_args(argv)

    output_dir = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)
    if args.clean_output:
        for existing_path in output_dir.glob("啟應文*.txt"):
            existing_path.unlink()

    pptx_files = sorted(args.input_dir.glob("*.pptx"), key=lambda path: int(READING_RE.search(path.stem).group(1)))
    if not pptx_files:
        print(f"No .pptx files found in {args.input_dir}", file=sys.stderr)
        return 1

    converted: list[Path] = []
    for pptx_path in pptx_files:
        converted.append(convert_one(pptx_path, output_dir, args.overwrite))

    print(f"Converted {len(converted)} files to {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
