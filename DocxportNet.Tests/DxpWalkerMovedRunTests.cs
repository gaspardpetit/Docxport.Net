using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Middleware;
using DocxportNet.Visitors;
using DocxportNet.Walker;

namespace DocxportNet.Tests;

public sealed class DxpWalkerMovedRunTests
{
    [Fact]
    public void DefaultVisitorDoesNotExpandMovedRuns()
    {
        using var document = CreateDocument();
        var visitor = new TextRecorder();
        new DxpWalker().Accept(document, new DxpContextMiddleware(visitor));

        Assert.Equal("before after", visitor.Text);
    }

    [Fact]
    public void VisitorCanExpandMovedRuns()
    {
        using var document = CreateDocument();
        var visitor = new TextRecorder(includeMovedFrom: true, includeMovedTo: true);
        new DxpWalker().Accept(document, new DxpContextMiddleware(visitor));

        Assert.Equal("before old new after", visitor.Text);
        Assert.True(visitor.WasMoveFromScopeDisposed);
        Assert.True(visitor.WasMoveToScopeDisposed);
    }

    [Fact]
    public void VisitorCanExpandOnlyMovedToContent()
    {
        using var document = CreateDocument();
        var visitor = new TextRecorder(includeMovedTo: true);
        new DxpWalker().Accept(document, new DxpContextMiddleware(visitor));

        Assert.Equal("before new after", visitor.Text);
        Assert.False(visitor.WasMoveFromScopeDisposed);
        Assert.True(visitor.WasMoveToScopeDisposed);
    }

    [Fact]
    public void VisitorCanExpandOnlyMovedFromContent()
    {
        using var document = CreateDocument();
        var visitor = new TextRecorder(includeMovedFrom: true);
        new DxpWalker().Accept(document, new DxpContextMiddleware(visitor));

        Assert.Equal("before old after", visitor.Text);
        Assert.True(visitor.WasMoveFromScopeDisposed);
        Assert.False(visitor.WasMoveToScopeDisposed);
    }

    private static WordprocessingDocument CreateDocument()
    {
        var document = WordprocessingDocument.Create(new MemoryStream(),
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(new Paragraph(
            new Run(new Text("before ")),
            new MoveFromRun(new Run(new Text("old "))),
            new MoveToRun(new Run(new Text("new "))),
            new Run(new Text("after")))));
        return document;
    }

    private sealed class TextRecorder : DxpVisitor
    {
        private readonly List<string> _parts = new();
        private readonly bool _includeMovedFrom;
        private readonly bool _includeMovedTo;

        public TextRecorder(bool includeMovedFrom = false, bool includeMovedTo = false) : base(null)
        {
            _includeMovedFrom = includeMovedFrom;
            _includeMovedTo = includeMovedTo;
        }

        public string Text => string.Concat(_parts);
        public bool WasMoveFromScopeDisposed { get; private set; }
        public bool WasMoveToScopeDisposed { get; private set; }

        public override IDisposable? VisitMoveFromRunBegin(MoveFromRun moved,
            DxpIDocumentContext context) => _includeMovedFrom
            ? DxpDisposable.Create(() => WasMoveFromScopeDisposed = true) : null;

        public override IDisposable? VisitMoveToRunBegin(MoveToRun moved,
            DxpIDocumentContext context) => _includeMovedTo
            ? DxpDisposable.Create(() => WasMoveToScopeDisposed = true) : null;

        public override void VisitText(Text text, DxpIDocumentContext context)
            => _parts.Add(text.Text);
    }
}
