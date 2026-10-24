"""Replace the reviewed visual baseline with a complete current scorecard."""

import argparse
import csv
import shutil
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--current", type=Path, required=True)
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--fixture-dir", type=Path, default=Path("fixtures/docx"))
    args = parser.parse_args()
    with args.current.open(encoding="utf-8-sig", newline="") as stream:
        scores = {row["Source"].replace("\\", "/") for row in csv.DictReader(stream)}
    fixtures = {path.relative_to(args.fixture_dir).as_posix()
                for path in args.fixture_dir.rglob("*.docx") if not path.name.startswith("~$")}
    if scores != fixtures:
        raise ValueError(f"Baseline requires the full corpus: {len(scores)} scores, "
                         f"{len(fixtures)} fixtures, {len(fixtures - scores)} missing")
    args.baseline.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(args.current, args.baseline)
    print(f"Updated {args.baseline} from {args.current} ({len(scores)} fixtures).")


if __name__ == "__main__":
    main()
