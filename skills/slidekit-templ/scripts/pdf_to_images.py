#!/usr/bin/env python3
"""Convert a PDF file to individual slide images.

Pipeline: PDF → JPEG images (pdftoppm)

Usage:
    python pptx_to_images.py <input.pdf> [output_dir]

Output:
    output_dir/slide-01.jpg, slide-02.jpg, ...

Dependencies:
    - Poppler (pdftoppm) — brew install poppler
"""

import argparse
import subprocess
import sys
from pathlib import Path


def pdf_to_images(pdf_path: Path, output_dir: Path, dpi: int = 150) -> list[Path]:
    """Convert PDF to JPEG images using pdftoppm."""
    prefix = output_dir / "slide"

    result = subprocess.run(
        [
            "pdftoppm",
            "-jpeg",
            "-r", str(dpi),
            str(pdf_path),
            str(prefix),
        ],
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        print(f"Error: Image conversion failed", file=sys.stderr)
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        sys.exit(1)

    images = sorted(output_dir.glob("slide-*.jpg"))
    if not images:
        print("Error: No images generated", file=sys.stderr)
        sys.exit(1)

    return images


def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF to slide images"
    )
    parser.add_argument("input", help="Input PDF file (.pdf)")
    parser.add_argument(
        "output_dir",
        nargs="?",
        default="slides",
        help="Output directory for images (default: slides)",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=150,
        help="Image resolution in DPI (default: 150)",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)

    if not input_path.exists():
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if input_path.suffix.lower() != ".pdf":
        print(f"Error: Expected .pdf file, got: {input_path.suffix}", file=sys.stderr)
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Generating slide images from: {input_path.name} (DPI={args.dpi})...")
    images = pdf_to_images(input_path, output_dir, args.dpi)

    print(f"\nDone! {len(images)} slides:")
    for img in images:
        print(f"  {img}")


if __name__ == "__main__":
    main()
