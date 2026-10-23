"""Render Word PDF references using the scorer's exact page pixel geometry."""

import argparse
import json
from fnmatch import fnmatch
from pathlib import Path

import numpy
import pymupdf


SIZES = ((400, 520), (1200, 1560))
RECIPE_VERSION = 1


def render_page(page, width, height, target):
    scale = min(width / page.rect.width, height / page.rect.height)
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
    canvas = numpy.full((height, width, 3), 255, dtype=numpy.uint8)
    pixels = numpy.frombuffer(pixmap.samples, dtype=numpy.uint8).reshape(
        pixmap.height, pixmap.width, pixmap.n
    )[:, :, :3]
    canvas[: pixmap.height, : pixmap.width] = pixels
    output = pymupdf.Pixmap(pymupdf.csRGB, width, height, canvas.tobytes(), False)
    target.parent.mkdir(parents=True, exist_ok=True)
    output.save(target)


def render_pdf(source, pdf_root, image_root):
    relative = source.relative_to(pdf_root).with_suffix("")
    destination = image_root / relative
    manifest_path = destination / "render.json"
    source_stat = source.stat()
    recipe = {
        "source_mtime_ns": source_stat.st_mtime_ns,
        "source_size": source_stat.st_size,
        "pymupdf": pymupdf.VersionBind,
        "recipe_version": RECIPE_VERSION,
        "sizes": [list(size) for size in SIZES],
    }
    try:
        prior = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        prior = None

    with pymupdf.open(source) as document:
        page_count = len(document)
        targets = [
            destination / f"{width}x{height}" / f"page-{number:04d}.png"
            for width, height in SIZES
            for number in range(1, page_count + 1)
        ]
        if prior == recipe and all(target.is_file() for target in targets):
            return 0, len(targets)

        for number, page in enumerate(document, start=1):
            for width, height in SIZES:
                target = destination / f"{width}x{height}" / f"page-{number:04d}.png"
                render_page(page, width, height, target)

    for width, height in SIZES:
        directory = destination / f"{width}x{height}"
        expected = {f"page-{number:04d}.png" for number in range(1, page_count + 1)}
        for stale in directory.glob("page-*.png"):
            if stale.name not in expected:
                stale.unlink()
    manifest_path.write_text(json.dumps(recipe, indent=2) + "\n", encoding="utf-8")
    return len(targets), 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pdf-dir", type=Path, default=Path("truth/pdf-word"))
    parser.add_argument("--image-dir", type=Path, default=Path("output/pdf/word-truth-images"))
    parser.add_argument("--name-filter", default="*")
    args = parser.parse_args()

    pdf_root = args.pdf_dir.resolve()
    image_root = args.image_dir.resolve()
    sources = sorted(
        source for source in pdf_root.rglob("*.pdf")
        if fnmatch(source.stem.casefold(), args.name_filter.casefold())
    )
    if not sources:
        parser.error(f"No Word PDFs matched {args.name_filter!r} under {pdf_root}")

    rendered = skipped = 0
    for index, source in enumerate(sources, start=1):
        made, kept = render_pdf(source, pdf_root, image_root)
        rendered += made
        skipped += kept
        if index % 100 == 0 or index == len(sources):
            print(f"{index}/{len(sources)} PDFs: {rendered} images rendered, {skipped} current", flush=True)


if __name__ == "__main__":
    main()
