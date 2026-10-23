"""Check that selected fixtures have readable Word or Flow PDFs."""

import argparse
import csv
from fnmatch import fnmatch
from pathlib import Path

import pymupdf


def relative_source(source):
    parts = source.parts
    for index in range(len(parts) - 1):
        if parts[index].casefold() == "fixtures" and parts[index + 1].casefold() == "docx":
            return Path(*parts[index + 2:])
    raise ValueError(f"Source is not under fixtures/docx: {source}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("kind", choices=("truth", "flow"))
    parser.add_argument("--fixture-dir", type=Path, required=True)
    parser.add_argument("--pdf-dir", type=Path, required=True)
    parser.add_argument("--name-filter", default="*")
    args = parser.parse_args()

    fixtures = {
        path.relative_to(args.fixture_dir)
        for path in args.fixture_dir.rglob("*.docx")
        if not path.name.startswith("~$") and
        fnmatch(path.stem.casefold(), args.name_filter.casefold())
    }
    if not fixtures:
        parser.error(f"No DOCX fixtures matched {args.name_filter!r}")

    rows = None
    if args.kind == "flow":
        with (args.pdf_dir / "flow-results.csv").open(encoding="utf-8-sig", newline="") as stream:
            rows = {
                relative_source(Path(row["Source"])): row
                for row in csv.DictReader(stream)
                if fnmatch(Path(row["Source"]).stem.casefold(), args.name_filter.casefold())
            }

    failures = []
    pages = 0
    for relative in sorted(fixtures):
        if rows is not None and relative not in rows:
            failures.append(f"Missing Flow export record: {relative}")
            continue
        if rows is not None and rows[relative]["Error"]:
            failures.append(f"Flow export failed: {relative}: {rows[relative]['Error']}")
            continue
        if args.kind == "truth":
            pdf = args.pdf_dir / relative.with_suffix(".pdf")
        else:
            pdf = args.pdf_dir / relative.with_suffix(".flow.pdf")
            if not pdf.is_file():
                pdf = args.pdf_dir / relative.parent.name / (relative.stem + ".flow.pdf")
        if not pdf.is_file():
            failures.append(f"Missing {args.kind} PDF: {relative}")
            continue
        try:
            with pymupdf.open(pdf) as document:
                if len(document) == 0:
                    raise ValueError("PDF has no pages")
                pages += len(document)
                if rows is not None and rows[relative]["ExportPages"] != str(len(document)):
                    raise ValueError("PDF page count differs from flow-results.csv")
        except Exception as error:
            failures.append(f"Unreadable {args.kind} PDF: {relative}: {error}")

    if rows is not None:
        failures.extend(f"Unexpected Flow record: {relative}" for relative in sorted(rows.keys() - fixtures))
    if failures:
        for failure in failures[:20]:
            print(failure)
        raise SystemExit(f"{len(failures)} {args.kind} PDF checks failed")
    print(f"Checked {len(fixtures)} {args.kind} PDFs ({pages} pages).")


if __name__ == "__main__":
    main()
