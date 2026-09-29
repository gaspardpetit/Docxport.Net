using System.Buffers.Binary;
using DocxportNet.Doc;
using OpenMcdf;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.Visitors.PlainText;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using DocxportNet.Wasm;

namespace DocxportNet.Tests;

public class DocStructureWalkerTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WalksContainerFibAndEnvelopeWithScopes(bool tableOne)
    {
        using var input = CreateDoc(tableOne);
        using var output = new StringWriter();
        var walker = new DocStructureWalker();
        var visitor = new DocStructurePrintVisitor(output);
        using var structure = walker.Accept(input, visitor);

        Assert.Equal(tableOne ? "1Table" : "0Table", structure.FibBase.TableStreamName);
        Assert.Equal(600u, structure.FindLocation("EmailEnvelope")!.Offset);
        Assert.Equal(20u, structure.FindLocation("EmailEnvelope")!.Length);
        var lines = output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.StartsWith("Document DOC", lines[0]);
        Assert.Contains(lines, line => line.Contains("Stream WordDocument") && line.Contains("path=WordDocument"));
        Assert.Contains(lines, line => line.Contains("Stream \\u0005SummaryInformation"));
        Assert.All(lines, line => Assert.All(line.TrimEnd('\r'), character => Assert.InRange((int)character, 0x20, 0x7E)));
        Assert.Contains(lines, line => line.Contains("FibBase FibBase") && line.Contains("tableStream="));
        Assert.Contains(lines, line => line.Contains("FibLocation EmailEnvelope") && line.Contains("offset=600"));
        Assert.Contains(lines, line => line.Contains("Dop2000 EnvelopeVisibility") && line.Contains("length=1"));
        Assert.Contains(lines, line => line.Contains("MsoEnvelopeCLSID EmailEnvelopeHeader") && line.Contains("length=20"));
        var headerNodes = structure.Root.Children.SelectMany(x => x.Children)
            .SelectMany(x => x.Children).SelectMany(x => x.Children).SelectMany(x => x.Children).ToList();
        var visibility = headerNodes.Single(x => x.Name == "EnvelopeVisibility");
        var envelopeHeader = headerNodes.Single(x => x.Name == "EmailEnvelopeHeader");
        Assert.True(((DocDopVisibility)visibility.Payload!).Visible);
        Assert.Equal(8u, ((DocEnvelopeHeader)envelopeHeader.Payload!).Version);
        Assert.Same(envelopeHeader.Payload, envelopeHeader.Payload);
        Assert.True(Array.FindIndex(lines, x => x.Contains("FibDirectory")) <
                    Array.FindIndex(lines, x => x.Contains("FibLocation EmailEnvelope")));
    }

    [Fact]
    public void RejectsEnvelopeLocationOutsideSelectedTable()
    {
        using var input = CreateDoc(true, badEnvelopeOffset: true);
        Assert.Throws<InvalidDataException>(() => new DocStructureWalker().Read(input));
    }

    [Fact]
    public void VisitorCanSkipEnvelopeWithoutReadingItsHeader()
    {
        using var input = CreateDoc(true);
        var visitor = new SkippingVisitor();
        using var structure = new DocStructureWalker().Accept(input, visitor);
        Assert.Contains(visitor.Names, name => name == "EmailEnvelope");
        Assert.DoesNotContain(visitor.Names, name => name == "EmailEnvelopeHeader");
        var envelopeNode = structure.Root.Children.SelectMany(x => x.Children)
            .SelectMany(x => x.Children).SelectMany(x => x.Children)
            .First(x => x.Name == "EmailEnvelope");
        Assert.Single(envelopeNode.Children);
        Assert.False(envelopeNode.Children[0].IsPayloadLoaded);
    }

    private sealed class SkippingVisitor : IDocStructureVisitor
    {
        public List<string> Names { get; } = new();

        public IDisposable? Enter(DocStructureNode node, int depth)
        {
            Names.Add(node.Name);
            return node.Name == "EmailEnvelope" ? null : DocxportNet.Core.DxpDisposable.Empty;
        }
    }

    [Fact]
    public void ClosesParentAfterChildrenEvenIfAChildThrows()
    {
        using var input = CreateDoc(true);
        var visitor = new ScopeVisitor();
        Assert.Throws<InvalidOperationException>(() => new DocStructureWalker().Accept(input, visitor));
        Assert.Equal(new[] { "enter:DOC", "enter:WordDocument", "close:WordDocument", "close:DOC" }, visitor.Events);
    }

    [Fact]
    public void XmlVisitorProducesNestedWellFormedStructure()
    {
        using var input = CreateDoc(true);
        using var output = new StringWriter();
        using (var xml = new DocStructureXmlVisitor(output))
        using (new DocStructureWalker().Accept(input, xml)) { }
        var document = XDocument.Parse(output.ToString());
        Assert.Equal("Document", document.Root?.Name.LocalName);
        Assert.Contains(document.Descendants("Stream"), x => (string?)x.Attribute("name") == "\\u0005SummaryInformation");
        Assert.Contains(document.Descendants("MsoEnvelopeCLSID"), x => (string?)x.Attribute("name") == "EmailEnvelopeHeader");
        Assert.DoesNotContain(output.ToString(), c => c > 0x7E);
    }

    [Fact]
    public void ClxWalkExposesTextLocationsWithoutReadingText()
    {
        using var input = CreateDocWithPieces();
        using var output = new StringWriter();
        using var structure = new DocStructureWalker().Accept(input, new DocStructurePrintVisitor(output));
        var clx = structure.Root.Children.SelectMany(x => x.Children)
            .SelectMany(x => x.Children).SelectMany(x => x.Children)
            .Single(x => x.Name == "TextPieceTable");
        Assert.Equal("Prc", clx.Children[0].Kind);
        var plc = clx.Children[1].Children.Single();
        Assert.Equal("PlcPcd", plc.Kind);
        Assert.Equal(2, plc.Children.Count);
        Assert.Equal("compressed", plc.Children[0].Attributes["encoding"]);
        Assert.Equal("0", plc.Children[0].Attributes["cpStart"]);
        Assert.Equal("5", plc.Children[0].Attributes["cpEnd"]);
        Assert.Equal("1000", plc.Children[0].Attributes["textOffset"]);
        Assert.Equal("utf16", plc.Children[1].Attributes["encoding"]);
        Assert.Equal("5", plc.Children[1].Attributes["cpStart"]);
        Assert.Equal("7", plc.Children[1].Attributes["cpEnd"]);
        Assert.Equal("1010", plc.Children[1].Attributes["textOffset"]);
        Assert.Equal(new[] { "Main", "Footnotes" }, structure.Parts.Select(x => x.Name));
        Assert.Equal("Main", plc.Children[0].Attributes["part"]);
        Assert.Equal(new[] { "Footnotes", "TerminalParagraphMark" }, plc.Children[1].Children.Select(x => x.Name));
        Assert.True(plc.Children[0].HasPayload);
        Assert.False(plc.Children[0].IsPayloadLoaded);
        var compressedText = (DocTextPieceContent)plc.Children[0].Payload!;
        Assert.True(plc.Children[0].IsPayloadLoaded);
        Assert.Equal("Hel—\r", compressedText.Text);
        Assert.Equal("Ω\r", ((DocTextPieceContent)plc.Children[1].Payload!).Text);
        Assert.Same(compressedText, plc.Children[0].Payload);
    }

    [Fact]
    public void TextIndexKeepsPiecesLazyAndSlicesMainStoryAcrossPieceBoundary()
    {
        using var input = CreateDocWithPieces(allMain: true, reverseStorage: true);
        using var index = new DocTextIndexWalker().Index(input);
        Assert.Equal(2, index.Pieces.Count);
        Assert.Equal(new long[] { 1000, 900 }, index.Pieces.Select(x => x.TextOffset));
        Assert.All(index.Pieces, piece => Assert.False(piece.IsTextLoaded));

        var spans = index.GetPartSpans("Main");
        Assert.Equal(new uint[] { 0, 5 }, spans.Select(x => x.CpStart));
        Assert.Equal(new uint[] { 5, 7 }, spans.Select(x => x.CpEnd));
        Assert.All(index.Pieces, piece => Assert.False(piece.IsTextLoaded));
        Assert.Equal("Hel—\rΩ\r", string.Concat(spans.Select(x => x.Text)));
        Assert.All(index.Pieces, piece => Assert.True(piece.IsTextLoaded));
        Assert.Same(index.Pieces[0].Node.Payload, index.Pieces[0].Node.Payload);
    }

    [Fact]
    public void TextIndexExcludesOtherPartsAndTerminalParagraphMark()
    {
        using var input = CreateDocWithPieces();
        using var index = new DocTextIndexWalker().Index(input);
        Assert.Equal("Hel—\r", string.Concat(index.GetPartSpans("Main").Select(x => x.Text)));
        Assert.Equal("Ω", string.Concat(index.GetPartSpans("Footnotes").Select(x => x.Text)));
        Assert.Empty(index.GetPartSpans("Headers"));
    }

    [Fact]
    public void ProjectsMainTextAcrossPhysicalPiecesIntoAReadableDocx()
    {
        using var input = CreateDocWithPieces(allMain: true, reverseStorage: true, splitParagraph: true);
        using var index = new DocTextIndexWalker().Index(input);
        var projection = new DocToDocxProjector().Project(index);

        using (var stream = new MemoryStream(projection.DocxBytes))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var paragraphs = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToArray();
            Assert.Single(paragraphs);
            Assert.Equal("Hel— Ω", paragraphs[0].InnerText);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        Assert.Equal(7u, projection.Coverage.ProjectedMainPart!.CpEnd);
        Assert.Empty(projection.Coverage.DeferredParts);
        Assert.Empty(projection.Coverage.OmittedCharacters);
        var text = DxpExport.ExportToString(projection.DocxBytes,
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
        Assert.Contains("Hel— Ω", text);
    }

    [Fact]
    public void ProjectionReportsOtherPartsAndOmitsTheirText()
    {
        using var input = CreateDocWithPieces();
        using var index = new DocTextIndexWalker().Index(input);
        var projection = new DocToDocxProjector().Project(index);
        using var stream = new MemoryStream(projection.DocxBytes);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Equal("Hel—", document.MainDocumentPart!.Document!.Body!.InnerText);
        Assert.Equal("Footnotes", Assert.Single(projection.Coverage.DeferredParts).Name);
        Assert.DoesNotContain(document.MainDocumentPart.Document.Body.InnerText, c => c == 'Ω');
    }

    [Fact]
    public void ProjectionTranslatesCommonControlsAndReportsLossyOnes()
    {
        using var input = CreateDocWithPieces(allMain: true, specialControls: true);
        using var index = new DocTextIndexWalker().Index(input);
        var projection = new DocToDocxProjector().Project(index);
        using var stream = new MemoryStream(projection.DocxBytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var paragraphs = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToArray();
        Assert.Single(paragraphs);
        Assert.Equal(2, paragraphs[0].Descendants<Break>().Count());
        Assert.Equal(2, paragraphs[0].Descendants<TabChar>().Count());
        Assert.Equal(1, projection.Coverage.ApproximateCharacters[0x0007]);
        Assert.Contains(paragraphs[0].Descendants<Break>(), x => x.Type?.Value == BreakValues.Page);
        Assert.False(projection.Coverage.ApproximateCharacters.ContainsKey(0x000C));
        Assert.Equal(1, projection.Coverage.OmittedCharacters[0x0013]);
    }

    [Fact]
    public void ExistingExportSurfaceAcceptsBinaryDocBytesAndPath()
    {
        using var input = CreateDocWithPieces(allMain: true, reverseStorage: true, splitParagraph: true);
        var bytes = input.ToArray();
        var fromBytes = DxpExport.ExportToString(bytes,
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
        Assert.Contains("Hel— Ω", fromBytes);

        var path = Path.Combine(Path.GetTempPath(), $"doc-export-{Guid.NewGuid():N}.doc");
        try
        {
            File.WriteAllBytes(path, bytes);
            var fromPath = DxpExport.ExportToString(path,
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
            Assert.Equal(fromBytes, fromPath);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }

        var docxBytes = DxpDocxExport.Export(bytes);
        using var stream = new MemoryStream(docxBytes);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Equal("Hel— Ω", document.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void DirectDocProjectionExposesCoverageAndMatchesExportInput()
    {
        using var input = CreateDocWithPieces();
        var bytes = input.ToArray();
        var projection = DxpDocToDocx.Project(bytes);

        using var projectedStream = new MemoryStream(projection.DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Equal("Hel—", projected.MainDocumentPart!.Document!.Body!.InnerText);
        Assert.Equal("Footnotes", Assert.Single(projection.Coverage.DeferredParts).Name);

        var directText = DxpExport.ExportToString(projection.DocxBytes,
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
        var docText = DxpExport.ExportToString(bytes,
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
        Assert.Equal(directText, docText);
    }

    [Fact]
    public void BrowserDirectProjectionReturnsIntermediateDocx()
    {
        using var input = CreateDocWithPieces(allMain: true, splitParagraph: true);
        var intermediate = BrowserExports.ProjectDocxForTests(input.ToArray());
        using var stream = new MemoryStream(intermediate);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Equal("Hel— Ω", document.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void DocVisitorWritesPlainTextDocumentFromDocx()
    {
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Hello"), new TabChar(), new Text("Ω"),
                    new Break(), new Text("World"))),
                new Paragraph(new Run(new Text("Next")))));
            main.Document.Save();
        }

        var docBytes = DxpDocExport.Export(source.ToArray());
        using (var compoundInput = new MemoryStream(docBytes, false))
        using (var compound = RootStorage.Open(compoundInput))
        using (var word = compound.OpenStream("WordDocument"))
        {
            var fib = new byte[68];
            Assert.Equal(fib.Length, word.Read(fib, 0, fib.Length));
            Assert.Equal(word.Length, BinaryPrimitives.ReadUInt32LittleEndian(fib.AsSpan(64)));
        }
        using var docStream = new MemoryStream(docBytes);
        using var index = new DocTextIndexWalker().Index(docStream);
        Assert.Equal("Hello\tΩ\vWorld\rNext\r",
            string.Concat(index.GetPartSpans("Main").Select(x => x.Text)));
    }

    [Fact]
    public void DocProjectionPreservesManualPageBreak()
    {
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(
                new Text("Before"), new Break { Type = BreakValues.Page }, new Text("After")))));
            main.Document.Save();
        }
        var docBytes = DxpDocExport.Export(source.ToArray());
        var projection = DxpDocToDocx.Project(docBytes);
        using var projectedStream = new MemoryStream(projection.DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        var paragraph = Assert.Single(projected.MainDocumentPart!.Document!.Body!.Elements<Paragraph>());
        Assert.Equal("BeforeAfter", paragraph.InnerText);
        Assert.Equal(BreakValues.Page, Assert.Single(paragraph.Descendants<Break>()).Type!.Value);
    }

    [Fact]
    public void ExportToFilesAcceptsBinaryDocPath()
    {
        using var input = CreateDocWithPieces(allMain: true);
        var inputPath = Path.Combine(Path.GetTempPath(), $"doc-merge-{Guid.NewGuid():N}.doc");
        var outputPath = Path.ChangeExtension(inputPath, ".txt");
        try
        {
            File.WriteAllBytes(inputPath, input.ToArray());
            var outputs = DxpExport.ExportToFiles(inputPath, new SingleRecordCursor(),
                _ => new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()),
                _ => outputPath);
            Assert.Single(outputs);
            Assert.Equal(outputs[0], File.ReadAllText(outputPath));
            Assert.Contains("Hel—", outputs[0]);
            Assert.Contains("Ω", outputs[0]);
        }
        finally
        {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    private sealed class SingleRecordCursor : IDxpMergeRecordCursor
    {
        public bool HasCurrent => true;
        public int RecordIndex => 1;
        public bool MoveNext() => false;
        public DxpFieldValue? GetValue(string fieldName) => null;
    }

    [Fact]
    public void DocWriterPreservesParagraphsAcrossFormattingPages()
    {
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(Enumerable.Range(0, 35)
                .Select(i => new Paragraph(new Run(new Text($"Paragraph {i} Ω"))))));
            main.Document.Save();
        }

        var docBytes = DxpDocExport.Export(source.ToArray());
        using var docStream = new MemoryStream(docBytes);
        using var index = new DocTextIndexWalker().Index(docStream);
        Assert.False(index.IsFormattingIndexLoaded);
        Assert.False(index.IsStyleIndexLoaded);
        Assert.False(index.IsSectionIndexLoaded);
        Assert.Equal(string.Concat(Enumerable.Range(0, 35).Select(i => $"Paragraph {i} Ω\r")),
            string.Concat(index.GetPartSpans("Main").Select(x => x.Text)));
        Assert.False(index.IsFormattingIndexLoaded);
        Assert.Equal(3, index.FormattingPages.Count);
        Assert.Equal(2, index.FormattingPages.Count(x => !x.IsCharacterFormatting));
        Assert.NotEmpty(index.Styles);
        Assert.Single(index.Sections);
        Assert.All(index.FormattingPages, page => Assert.False(page.AreRunsLoaded));
        var paragraphs = index.FormattingPages.Where(x => !x.IsCharacterFormatting)
            .SelectMany(x => x.Runs).ToArray();
        Assert.Equal(35, paragraphs.Length);
        Assert.All(paragraphs, node => Assert.Equal("PapxRange", node.Kind));
        Assert.True(index.FormattingPages.Last().AreRunsLoaded);

        using var walkerInput = new MemoryStream(docBytes);
        using var xmlOutput = new StringWriter();
        using (var xml = new DocStructureXmlVisitor(xmlOutput))
        using (new DocStructureWalker().Accept(walkerInput, xml)) { }
        var dump = XDocument.Parse(xmlOutput.ToString());
        Assert.Equal(2, dump.Descendants("PapxFkp").Count());
        Assert.Equal(35, dump.Descendants("PapxRange").Count());
        Assert.NotEmpty(dump.Descendants("STSH"));
    }

    [Fact]
    public void DocVisitorAcceptsBinaryDocThroughProjection()
    {
        using var input = CreateDocWithPieces(allMain: true, splitParagraph: true);
        var rewritten = DxpDocExport.Export(input.ToArray());
        using var output = new MemoryStream(rewritten);
        using var index = new DocTextIndexWalker().Index(output);
        Assert.Equal("Hel— Ω\r", string.Concat(index.GetPartSpans("Main").Select(x => x.Text)));
    }

    [Fact]
    public void BrowserDocExportReturnsBinaryDoc()
    {
        using var input = CreateDocWithPieces(allMain: true, splitParagraph: true);
        var bytes = BrowserExports.ExportDocForTests(input.ToArray(), new BrowserResolveRequest());
        using var output = new MemoryStream(bytes);
        using var index = new DocTextIndexWalker().Index(output);
        Assert.Equal("Hel— Ω\r", string.Concat(index.GetPartSpans("Main").Select(x => x.Text)));
    }

    [Theory]
    [InlineData(BrowserExportFormat.Html)]
    [InlineData(BrowserExportFormat.Markdown)]
    [InlineData(BrowserExportFormat.Text)]
    public void BrowserExportSurfaceAcceptsBinaryDocBytes(BrowserExportFormat format)
    {
        using var input = CreateDocWithPieces(allMain: true, splitParagraph: true);
        var text = BrowserExports.ExportForTests(input.ToArray(), new BrowserExportRequest
        {
            Format = format
        });
        Assert.Contains("Hel", text);
        Assert.Contains("Ω", text);
        var docxBytes = BrowserExports.ResolveDocxForTests(input.ToArray(), new BrowserResolveRequest());
        using var stream = new MemoryStream(docxBytes);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Equal("Hel— Ω", document.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void SkippingClxAvoidsParsingMalformedPieceTable()
    {
        using var input = CreateDocWithPieces(badClxTag: true);
        var visitor = new NamedSkippingVisitor("TextPieceTable");
        using var structure = new DocStructureWalker().Accept(input, visitor);
        Assert.Contains("TextPieceTable", visitor.Names);
        Assert.DoesNotContain("PieceTable", visitor.Names);
    }

    [Fact]
    public void XmlVisitorDecodesPiecesAndEscapesControlCharacters()
    {
        using var input = CreateDocWithPieces();
        using var output = new StringWriter();
        using (var xml = new DocStructureXmlVisitor(output))
        using (new DocStructureWalker().Accept(input, xml)) { }
        var document = XDocument.Parse(output.ToString());
        var texts = document.Descendants("Pcd").Select(x => (string?)x.Element("Text")).ToArray();
        Assert.Equal(new[] { "Hel\\u2014\\u000D", "\\u03A9\\u000D" }, texts);
    }

    [Fact]
    public void SectionTableExposesCpRangesAndLazyPropertyLocations()
    {
        using var input = CreateDocWithSections();
        using var output = new StringWriter();
        using (var xml = new DocStructureXmlVisitor(output))
        using (new DocStructureWalker().Accept(input, xml)) { }
        var document = XDocument.Parse(output.ToString());
        var sections = document.Descendants("Sed").ToArray();
        Assert.Equal(2, sections.Length);
        Assert.Equal("0", sections[0].Elements("attribute").Single(x => (string?)x.Attribute("name") == "cpStart").Attribute("value")?.Value);
        Assert.Equal("5", sections[0].Elements("attribute").Single(x => (string?)x.Attribute("name") == "cpEnd").Attribute("value")?.Value);
        Assert.Equal("5", sections[1].Elements("attribute").Single(x => (string?)x.Attribute("name") == "cpStart").Attribute("value")?.Value);
        Assert.Equal("10", sections[1].Elements("attribute").Single(x => (string?)x.Attribute("name") == "cpEnd").Attribute("value")?.Value);
        Assert.Equal("900", sections[0].Element("Sepx")?.Attribute("offset")?.Value);
        Assert.Equal("6", sections[0].Element("Sepx")?.Attribute("length")?.Value);
        Assert.Null(sections[1].Element("Sepx"));
    }

    [Fact]
    public void SkippingSectionsLeavesMalformedSectionTableUnread()
    {
        using var input = CreateDocWithSections(badSectionTable: true);
        var visitor = new NamedSkippingVisitor("Sections");
        using var structure = new DocStructureWalker().Accept(input, visitor);
        Assert.Contains("Sections", visitor.Names);
        Assert.DoesNotContain("SectionTable", visitor.Names);
    }

    private sealed class NamedSkippingVisitor(string skipName) : IDocStructureVisitor
    {
        public List<string> Names { get; } = new();
        public IDisposable? Enter(DocStructureNode node, int depth)
        {
            Names.Add(node.Name);
            return node.Name == skipName ? null : DocxportNet.Core.DxpDisposable.Empty;
        }
    }

    private sealed class ScopeVisitor : IDocStructureVisitor
    {
        public List<string> Events { get; } = new();

        public IDisposable Enter(DocStructureNode node, int depth)
        {
            if (node.Name == "WordDocument")
            {
                Events.Add("enter:WordDocument");
                return DocxportNet.Core.DxpDisposable.Create(() => Events.Add("close:WordDocument"));
            }
            if (node.Kind == "FIB") throw new InvalidOperationException("Stop inside WordDocument");
            if (node.Name == "DOC")
            {
                Events.Add("enter:DOC");
                return DocxportNet.Core.DxpDisposable.Create(() => Events.Add("close:DOC"));
            }
            return DocxportNet.Core.DxpDisposable.Empty;
        }
    }

    private static MemoryStream CreateDoc(bool tableOne, bool badEnvelopeOffset = false)
    {
        var output = new MemoryStream();
        using (var root = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen))
        {
            var word = new byte[1200];
            BinaryPrimitives.WriteUInt16LittleEndian(word, 0xA5EC);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(2), 0x00C1);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(10), tableOne ? (ushort)0x0200 : (ushort)0);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(32), 14);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(62), 22);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(152), 108);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 31 * 8), 0);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 31 * 8 + 4), 600);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 97 * 8), badEnvelopeOffset ? 10000u : 600u);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 97 * 8 + 4), 20);
            using (var stream = root.CreateStream("WordDocument")) stream.Write(word);
            var table = new byte[620];
            table[504] = 2;
            new Guid("0006F01A-0000-0000-C000-000000000046").ToByteArray().CopyTo(table, 600);
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(616), 8);
            using (var stream = root.CreateStream(tableOne ? "1Table" : "0Table")) stream.Write(table);
            using (var stream = root.CreateStream("Unrelated")) stream.WriteByte(42);
            using (var stream = root.CreateStream("\u0005SummaryInformation")) stream.WriteByte(1);
        }
        output.Position = 0;
        return output;
    }

    private static MemoryStream CreateDocWithPieces(bool badClxTag = false, bool allMain = false,
        bool reverseStorage = false, bool splitParagraph = false, bool specialControls = false)
    {
        var output = new MemoryStream();
        using (var root = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen))
        {
            var word = new byte[1200];
            BinaryPrimitives.WriteUInt16LittleEndian(word, 0xA5EC);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(2), 0x00C1);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(10), 0x0200);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(32), 14);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(62), 22);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(76), allMain ? 7u : 5u);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(80), allMain ? 0u : 1u);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(152), 108);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 31 * 8 + 4), 600);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 33 * 8), 600);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 33 * 8 + 4), 38);
            (specialControls
                ? new byte[] { (byte)'A', 0x09, 0x0B, 0x07, 0x0C }
                : new byte[] { (byte)'H', (byte)'e', (byte)'l', 0x97,
                    splitParagraph ? (byte)' ' : (byte)0x0D }).CopyTo(word, 1000);
            var secondTextOffset = reverseStorage ? 900 : 1010;
            System.Text.Encoding.Unicode.GetBytes(specialControls ? "\u0013\r" : "Ω\r")
                .CopyTo(word, secondTextOffset);
            using (var stream = root.CreateStream("WordDocument")) stream.Write(word);
            var table = new byte[638];
            table[600] = badClxTag ? (byte)0x7F : (byte)1;
            BinaryPrimitives.WriteUInt16LittleEndian(table.AsSpan(601), 2);
            table[605] = 2;
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(606), 28);
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(614), 5);
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(618), 7);
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(624), 0x40000000u | 2000u);
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(632), (uint)secondTextOffset);
            using (var stream = root.CreateStream("1Table")) stream.Write(table);
        }
        output.Position = 0;
        return output;
    }

    private static MemoryStream CreateDocWithSections(bool badSectionTable = false)
    {
        var output = new MemoryStream();
        using (var root = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen))
        {
            var word = new byte[1200];
            BinaryPrimitives.WriteUInt16LittleEndian(word, 0xA5EC);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(2), 0x00C1);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(10), 0x0200);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(32), 14);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(62), 22);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(76), 10);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(152), 108);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 6 * 8), 500);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 6 * 8 + 4), badSectionTable ? 35u : 36u);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(900), 4);
            using (var stream = root.CreateStream("WordDocument")) stream.Write(word);
            var table = new byte[600];
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(504), 5);
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(508), 10);
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(514), 900);
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(526), uint.MaxValue);
            using (var stream = root.CreateStream("1Table")) stream.Write(table);
        }
        output.Position = 0;
        return output;
    }
}
