import argparse
import csv
import re
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile


WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
FEATURE_TAGS = {
    "tables": (WORD, "tbl"),
    "drawings": (WORD, "drawing"),
    "legacy-images": (WORD, "pict"),
    "text-boxes": (WORD, "txbxContent"),
    "math": (MATH, "oMath"),
    "ruby": (WORD, "ruby"),
    "footnotes": (WORD, "footnoteReference"),
    "endnotes": (WORD, "endnoteReference"),
    "fields": (WORD, "fldChar"),
    "simple-fields": (WORD, "fldSimple"),
    "hyperlinks": (WORD, "hyperlink"),
    "comments": (WORD, "commentReference"),
    "insertions": (WORD, "ins"),
    "deletions": (WORD, "del"),
    "moved-text": (WORD, "moveFrom"),
    "formatting-changes": (WORD, "rPrChange"),
    "alt-chunks": (WORD, "altChunk"),
    "floating-anchors": ("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "anchor"),
    "fit-text": (WORD, "fitText"),
    "tabs": (WORD, "tabs"),
    "numbering": (WORD, "numPr"),
    "bidi": (WORD, "bidi"),
}
STORIES = ("document.xml", "header", "footer", "footnotes", "endnotes", "comments")
FIELDS = ("source", "kind", "family", "cluster", "features")


def cluster_for(family: str, features: set[str]) -> str:
    name = family.lower()
    if name.startswith(("punctuation", "closing-mark", "repeated-space", "nobreak-wrap")):
        return "text-breaking"
    if name.startswith("russian-"):
        return "lists"
    if name.startswith(("omml", "nested-omml", "mixed-omml", "display-mixed-omml", "mixed-radical")):
        return "math"
    if name.startswith(("ruby", "highansi", "theme-shade", "cjk")):
        return "fonts-cjk"
    if name.startswith(("endnote", "footnote")):
        return "notes"
    if name.startswith(("floating-table", "table")):
        return "tables"
    if name.startswith(("vml", "square-wrap", "floating-image", "inline-image", "exact-inline-image", "frame", "text-box", "wordtightdiamond")):
        return "images-shapes"
    if name.startswith("altchunk"):
        return "alt-chunks"
    if name.startswith(("styled-number", "line-spacing", "fit-", "exact-line-grid")):
        return "text-layout"
    for feature, cluster in (
        ("comments", "annotations"), ("insertions", "revisions"),
        ("deletions", "revisions"), ("moved-text", "revisions"),
        ("formatting-changes", "revisions"),
        ("math", "math"), ("ruby", "fonts-cjk"),
        ("footnotes", "notes"), ("endnotes", "notes"),
        ("text-boxes", "images-shapes"), ("floating-anchors", "images-shapes"),
        ("tables", "tables"), ("fields", "fields-links"),
        ("hyperlinks", "fields-links"), ("drawings", "images-shapes"),
        ("fit-text", "text-layout"), ("numbering", "lists"),
    ):
        if feature in features:
            return cluster
    return "text-layout"


def features_for(path: Path) -> set[str]:
    tags = set()
    with ZipFile(path) as package:
        for name in package.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            if not any(story in name for story in STORIES):
                continue
            tags.update(element.tag for element in ElementTree.fromstring(package.read(name)).iter())
    return {
        feature for feature, (namespace, local) in FEATURE_TAGS.items()
        if f"{{{namespace}}}{local}" in tags
    }


def sources(repository: Path):
    fixtures = repository / "DocxportNet.Tests" / "Fixtures"
    yield from sorted(fixtures.rglob("*.docx"))
    yield from sorted((repository / "samples").glob("*.docx"))


def family_for(path: Path) -> str:
    if path.parent.name != "PdfProbeCorpus":
        return path.parent.name
    if "__" in path.stem:
        return path.stem.split("__", 1)[0]
    if path.stem.startswith("WordTightDiamond"):
        return "WordTightDiamond"
    return re.sub(r"\d+$", "", path.stem).rstrip("-")


def build_rows(repository: Path) -> list[dict[str, str]]:
    rows = []
    for path in sources(repository):
        if path.name.startswith("~$") or path.parent.name == "samples" and path.stem.endswith(".projected"):
            continue
        try:
            features = features_for(path)
        except (BadZipFile, ElementTree.ParseError) as error:
            raise ValueError(f"Cannot inventory {path}: {error}") from error
        kind = "sample" if path.parent.name == "samples" else "fixture"
        family = family_for(path)
        rows.append({
            "source": path.relative_to(repository).as_posix(),
            "kind": kind,
            "family": family,
            "cluster": "sample" if kind == "sample" else cluster_for(family, features),
            "features": ";".join(sorted(features)),
        })
    return sorted(rows, key=lambda row: row["source"].lower())


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    repository = Path(__file__).resolve().parents[2]
    output = args.output or Path(__file__).with_name("pdf-corpus-catalog.csv")
    rows = build_rows(repository)
    if args.check:
        with output.open(newline="", encoding="utf-8-sig") as stream:
            existing = list(csv.DictReader(stream))
        if existing != rows:
            raise SystemExit(f"Catalog is stale: regenerate {output}")
    else:
        with output.open("w", newline="", encoding="utf-8") as stream:
            writer = csv.DictWriter(stream, fieldnames=FIELDS)
            writer.writeheader()
            writer.writerows(rows)
    counts = Counter(row["cluster"] for row in rows)
    print(f"{len(rows)} documents: {dict(sorted(counts.items()))}")


if __name__ == "__main__":
    main()
