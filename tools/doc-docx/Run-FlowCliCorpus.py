"""Export the DOCX fixture corpus through the Docxport CLI for visual scoring."""

import argparse
import csv
from fnmatch import fnmatch
from pathlib import Path
import subprocess

import pymupdf


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fixture-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--name-filter", default="*")
    parser.add_argument("--dotnet", default="dotnet")
    parser.add_argument("--cli-dll", type=Path, required=True)
    args = parser.parse_args()

    fixture_root = args.fixture_dir.resolve()
    output_root = args.output_dir.resolve()
    cli_dll = args.cli_dll.resolve()
    if not cli_dll.is_file():
        parser.error(f"CLI DLL does not exist: {cli_dll}")

    sources = sorted(
        (path for path in fixture_root.rglob("*.docx")
         if not path.name.startswith("~$") and
         fnmatch(path.stem.casefold(), args.name_filter.casefold())),
        key=lambda path: str(path).casefold(),
    )
    if not sources:
        parser.error(f"No DOCX fixtures matched {args.name_filter!r}")

    output_root.mkdir(parents=True, exist_ok=True)
    report = output_root / "flow-results.csv"
    failures = 0
    with report.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=("Source", "ExportPages", "ExportBytes", "Error"))
        writer.writeheader()
        for index, source in enumerate(sources, start=1):
            relative = source.relative_to(fixture_root)
            target = output_root / relative.with_suffix(".flow.pdf")
            target.parent.mkdir(parents=True, exist_ok=True)
            target.unlink(missing_ok=True)
            row = {"Source": str(source), "ExportPages": "", "ExportBytes": "", "Error": ""}
            try:
                result = subprocess.run(
                    (args.dotnet, str(cli_dll), str(source), "--format=pdf",
                     "--fields=cache", "--no-progress", "--output", str(target)),
                    capture_output=True, text=True, errors="replace", check=False,
                )
                if result.returncode != 0:
                    detail = result.stderr.strip() or result.stdout.strip()
                    raise RuntimeError(f"CLI exited {result.returncode}: {detail}")
                if not target.is_file():
                    raise RuntimeError("CLI succeeded without creating a PDF")
                with pymupdf.open(target) as document:
                    if len(document) == 0:
                        raise ValueError("PDF has no pages")
                    row["ExportPages"] = len(document)
                row["ExportBytes"] = target.stat().st_size
            except Exception as error:
                failures += 1
                row["Error"] = f"{type(error).__name__}: {error}"
                print(f"{relative}: {row['Error']}", flush=True)
            writer.writerow(row)
            stream.flush()
            if index % 100 == 0 or index == len(sources):
                print(f"Exported {index}/{len(sources)} fixtures", flush=True)

    print(f"Flow report: {report}")
    if failures:
        raise SystemExit(f"{failures} Flow exports failed")


if __name__ == "__main__":
    main()
