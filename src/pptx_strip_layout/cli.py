from __future__ import annotations

import argparse
import sys

from .processor import process_presentation


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx-strip-layout",
        description="Strip template layout objects and normalize PPTX slide styles.",
    )
    parser.add_argument("filename", help="Path to the input .pptx file")
    parser.add_argument(
        "--grayscale-images",
        action="store_true",
        help="Convert all images on slides to grayscale and replace originals.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        output_path = process_presentation(
            args.filename,
            grayscale_images=args.grayscale_images,
        )
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(str(output_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
