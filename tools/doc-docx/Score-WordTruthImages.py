"""Compare existing Flow PDFs with the cached Word scoring images."""

import argparse
import csv
import importlib.util
import json
from fnmatch import fnmatch
from pathlib import Path

import numpy
import pymupdf

_spec = importlib.util.spec_from_file_location(
    "word_truth_renderer", Path(__file__).with_name("Render-WordTruthImages.py")
)
RENDERER = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(RENDERER)


def source_relative_path(source):
    parts = source.parts
    for index in range(len(parts) - 1):
        if parts[index].casefold() == "fixtures" and parts[index + 1].casefold() == "docx":
            return Path(*parts[index + 2:])
    raise ValueError(f"Source is not under fixtures/docx: {source}")


def image_pixels(path, width, height):
    pixmap = pymupdf.Pixmap(str(path))
    if (pixmap.width, pixmap.height, pixmap.n) != (width, height, 3):
        raise ValueError(f"Unexpected image geometry or colorspace: {path}")
    return numpy.frombuffer(pixmap.samples, dtype=numpy.uint8).reshape(height, width, 3)


def flow_pixels(page, width, height):
    scale = min(width / page.rect.width, height / page.rect.height)
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
    canvas = numpy.full((height, width, 3), 255, dtype=numpy.uint8)
    pixels = numpy.frombuffer(pixmap.samples, dtype=numpy.uint8).reshape(
        pixmap.height, pixmap.width, pixmap.n
    )[:, :, :3]
    canvas[: pixmap.height, : pixmap.width] = pixels
    return canvas


def dilate_ink(mask):
    padded = numpy.pad(mask, ((1, 1), (1, 1)), mode="constant")
    return numpy.logical_or.reduce([
        padded[1 + dy:1 + dy + mask.shape[0], 1 + dx:1 + dx + mask.shape[1]]
        for dy in (-1, 0, 1) for dx in (-1, 0, 1)
    ])


def smooth_pixels(pixels):
    padded = numpy.pad(pixels, ((1, 1), (1, 1), (0, 0)), mode="edge")
    return sum(
        padded[1 + dy:1 + dy + pixels.shape[0],
               1 + dx:1 + dx + pixels.shape[1]] *
        ((1, 2, 1)[dy + 1] * (1, 2, 1)[dx + 1] / 16)
        for dy in (-1, 0, 1) for dx in (-1, 0, 1)
    )


def page_scores(word_small, flow_small, word_large, flow_large):
    reference = word_small.astype(numpy.float32)
    actual = flow_small.astype(numpy.float32)
    difference = reference - actual
    smoothed = smooth_pixels(reference) - smooth_pixels(actual)
    color_reference = reference.max(axis=2) - reference.min(axis=2) > 35
    color_actual = actual.max(axis=2) - actual.min(axis=2) > 35
    soft_reference = reference.max(axis=2) - reference.min(axis=2) > 20
    soft_actual = actual.max(axis=2) - actual.min(axis=2) > 20
    ink_reference = reference.min(axis=2) < 160
    ink_actual = actual.min(axis=2) < 160
    high_reference = word_large.min(axis=2) < 160
    high_actual = flow_large.min(axis=2) < 160

    def mismatch(left, right):
        count = left.sum() + right.sum()
        return float((numpy.count_nonzero(left & ~dilate_ink(right)) +
                      numpy.count_nonzero(right & ~dilate_ink(left))) / count) if count else 0.0

    return {
        "VisualRmse": float(numpy.sqrt(numpy.mean(difference * difference)) / 255),
        "SmoothedVisualRmse": float(numpy.sqrt(numpy.mean(smoothed * smoothed)) / 255),
        "ColorMaskDifference": float(numpy.count_nonzero(color_reference != color_actual) / 208000),
        "ColorMaskDifference1px": float((numpy.count_nonzero(soft_reference & ~dilate_ink(soft_actual)) +
                                         numpy.count_nonzero(soft_actual & ~dilate_ink(soft_reference))) / 208000),
        "InkMismatch1px": mismatch(ink_reference, ink_actual),
        "InkMismatchHiRes1px": mismatch(high_reference, high_actual),
    }


def verified_images(pdf, pdf_root, image_root):
    destination = image_root / pdf.relative_to(pdf_root).with_suffix("")
    manifest_path = destination / "render.json"
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise ValueError(f"Missing Word images for {pdf}; run make truth-image") from error
    stat = pdf.stat()
    expected = {
        "source_mtime_ns": stat.st_mtime_ns,
        "source_size": stat.st_size,
        "pymupdf": pymupdf.VersionBind,
        "recipe_version": RENDERER.RECIPE_VERSION,
        "sizes": [list(size) for size in RENDERER.SIZES],
    }
    if manifest != expected:
        raise ValueError(f"Stale Word images for {pdf}; run make truth-image")
    return destination


def write_csv(path, rows, fieldnames):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--flow-dir", type=Path, required=True)
    parser.add_argument("--truth-pdf-dir", type=Path, required=True)
    parser.add_argument("--truth-image-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--name-filter", default="*")
    args = parser.parse_args()

    flow_root = args.flow_dir.resolve()
    pdf_root = args.truth_pdf_dir.resolve()
    image_root = args.truth_image_dir.resolve()
    with (flow_root / "flow-results.csv").open(encoding="utf-8-sig", newline="") as stream:
        exports = [row for row in csv.DictReader(stream)
                   if fnmatch(Path(row["Source"]).stem.casefold(), args.name_filter.casefold())]
    if not exports:
        parser.error(f"No Flow exports matched {args.name_filter!r}")

    documents = []
    pages = []
    for position, row in enumerate(exports, start=1):
        if row["Error"]:
            raise ValueError(f"Flow export failed for {row['Source']}: {row['Error']}")
        source = Path(row["Source"])
        relative = source_relative_path(source)
        word_pdf = pdf_root / relative.with_suffix(".pdf")
        flow_pdf = flow_root / relative.with_suffix(".flow.pdf")
        if not flow_pdf.is_file():
            flow_pdf = flow_root / source.parent.name / (source.stem + ".flow.pdf")
        image_dir = verified_images(word_pdf, pdf_root, image_root)
        with pymupdf.open(word_pdf) as reference, pymupdf.open(flow_pdf) as actual:
            values = []
            for index in range(min(len(reference), len(actual))):
                number = index + 1
                def word(width, height):
                    path = image_dir / f"{width}x{height}" / f"page-{number:04d}.png"
                    return image_pixels(path, width, height)

                score = page_scores(
                    word(400, 520), flow_pixels(actual[index], 400, 520),
                    word(1200, 1560), flow_pixels(actual[index], 1200, 1560)
                )
                values.append(score)
                pages.append({"Source": str(relative), "Page": number,
                              **{key: round(value, 6) for key, value in score.items()}})
            if not values:
                raise ValueError(f"No pages to compare: {relative}")
            document = {"Source": str(relative), "ReferencePages": len(reference),
                        "ExportPages": len(actual)}
            for key in values[0]:
                value = sum(item[key] for item in values) / len(values)
                if key.startswith("InkMismatch"):
                    value = (sum(item[key] for item in values) + abs(len(reference) - len(actual))) / max(
                        len(reference), len(actual)
                    )
                document[key] = round(value, 6)
            documents.append(document)
        if position % 100 == 0 or position == len(exports):
            print(f"Scored {position}/{len(exports)} documents", flush=True)

    output = args.output_dir.resolve()
    write_csv(output / "flow-word-visual-results.csv", documents, documents[0].keys())
    write_csv(output / "flow-word-page-visual-results.csv", pages, pages[0].keys())
    print(f"Scored {len(documents)} documents, {len(pages)} paired pages.")
    print(f"Page-count mismatches: {sum(row['ReferencePages'] != row['ExportPages'] for row in documents)}")
    print(f"Report: {output / 'flow-word-visual-results.csv'}")


if __name__ == "__main__":
    main()
