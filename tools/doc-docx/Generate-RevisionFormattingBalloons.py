from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED


ROOT = Path(__file__).resolve().parents[2]
FIXTURES = ROOT / "DocxportNet.Tests/Fixtures/PdfFeatures"
SOURCE = FIXTURES / "RevisionFormattingBalloonA4.docx"
TARGETS = (
    (FIXTURES / "RevisionFormattingBalloonThirdItem.docx", True),
    (FIXTURES / "RevisionFormattingBalloonTwoItems.docx", False),
)


for target, move_change in TARGETS:
    with ZipFile(SOURCE) as original, ZipFile(target, "w", ZIP_DEFLATED) as output:
        for entry in original.infolist():
            data = original.read(entry.filename)
            if entry.filename == "word/document.xml":
                paragraphs = re.split(r"(?=<w:p(?:\s|>))", data.decode("utf-8"))
                first_index = next(index for index, paragraph in enumerate(paragraphs)
                                   if "<w:t>Original 1</w:t>" in paragraph)
                third_index = next(index for index, paragraph in enumerate(paragraphs)
                                   if "Original 3</w:t>" in paragraph)
                change = re.search(r"<w:rPrChange\b.*?</w:rPrChange>",
                                   paragraphs[first_index], re.S)
                if change is None:
                    raise ValueError("Formatting change not found in the first item")
                if move_change:
                    paragraphs[first_index] = paragraphs[first_index].replace(change.group(0), "", 1)
                    added_change = change.group(0)
                else:
                    added_change = change.group(0).replace('w:id="2"', 'w:id="3"', 1)
                original_properties = (
                    '<w:r><w:rPr><w:lang w:val="fr-CA" '
                    'w:eastAsia="ja-JP" w:bidi="th-TH"/>'
                )
                updated_properties = (
                    '<w:r><w:rPr><w:lang w:val="en-US" '
                    'w:eastAsia="ja-JP" w:bidi="th-TH"/>' + added_change
                )
                if original_properties not in paragraphs[third_index]:
                    raise ValueError("Third item run properties not found")
                paragraphs[third_index] = paragraphs[third_index].replace(
                    original_properties, updated_properties, 1)
                data = "".join(paragraphs).encode("utf-8")
            output.writestr(entry, data)
    print(target)
