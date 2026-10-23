"""Report visual score improvements and regressions against a reviewed baseline."""

import argparse
import csv
from pathlib import Path


METRICS = ("VisualRmse", "SmoothedVisualRmse", "ColorMaskDifference",
           "ColorMaskDifference1px", "InkMismatch1px", "InkMismatchHiRes1px")


def read_scores(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream))
    if not rows:
        raise ValueError(f"Empty scorecard: {path}")
    scores = {}
    for row in rows:
        source = row["Source"].replace("\\", "/").casefold()
        if source in scores:
            raise ValueError(f"Duplicate source in {path}: {row['Source']}")
        scores[source] = row
    return scores


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--current", type=Path, required=True)
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--tolerance", type=float, default=0.001)
    args = parser.parse_args()
    if args.tolerance < 0:
        parser.error("Tolerance cannot be negative")

    current = read_scores(args.current)
    baseline = read_scores(args.baseline)
    missing = current.keys() - baseline.keys()
    if missing:
        raise ValueError(f"{len(missing)} fixtures have no baseline: " + ", ".join(sorted(missing)[:10]))

    changes = []
    for source, now in current.items():
        before = baseline[source]
        delta = {metric + "Delta": round(float(now[metric]) - float(before[metric]), 6)
                 for metric in METRICS}
        page_change = (int(now["ReferencePages"]) != int(before["ReferencePages"]) or
                       int(now["ExportPages"]) != int(before["ExportPages"]))
        primary = delta["SmoothedVisualRmseDelta"]
        status = "pages changed" if page_change else "regression" if primary > args.tolerance else \
            "improvement" if primary < -args.tolerance else "within tolerance"
        changes.append({"Source": now["Source"], "Status": status,
                        "BaselinePages": before["ExportPages"], "CurrentPages": now["ExportPages"],
                        "BaselineSmoothedVisualRmse": before["SmoothedVisualRmse"],
                        "CurrentSmoothedVisualRmse": now["SmoothedVisualRmse"], **delta})
    changes.sort(key=lambda row: abs(row["SmoothedVisualRmseDelta"]), reverse=True)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=changes[0].keys())
        writer.writeheader()
        writer.writerows(changes)

    counts = {status: sum(row["Status"] == status for row in changes)
              for status in ("regression", "improvement", "pages changed", "within tolerance")}
    print(f"Compared {len(changes)} documents: {counts['improvement']} improvements, "
          f"{counts['regression']} regressions, {counts['pages changed']} page changes, "
          f"{counts['within tolerance']} within tolerance.")
    for row in changes[:10]:
        print(f"  {row['SmoothedVisualRmseDelta']:+.6f} {row['Source']} ({row['Status']})")
    print(f"Diff report: {args.output}")
    if counts["regression"] or counts["pages changed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
