using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Middleware;
using DocxportNet.Visitors;
using DocxportNet.Walker;

namespace DocxportNet.Tests;

public sealed class RubyWalkerTests
{
    [Fact]
    public void WalksParagraphLevelRuby()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "RubyBasic.docx");
        using var document = WordprocessingDocument.Open(source, false);
        var visitor = new RubyTextRecorder();

        new DxpWalker().Accept(document, new DxpContextMiddleware(visitor));

        Assert.Equal(1, visitor.RubyCount);
        Assert.Contains("にほん", visitor.Text);
        Assert.Contains("日本", visitor.Text);
    }

    private sealed class RubyTextRecorder : DxpVisitor
    {
        public RubyTextRecorder() : base(null) { }

        public int RubyCount { get; private set; }
        public List<string> Text { get; } = new();

        public override IDisposable VisitRubyBegin(Ruby ruby, DxpIDocumentContext context)
        {
            RubyCount++;
            return base.VisitRubyBegin(ruby, context);
        }

        public override void VisitText(Text text, DxpIDocumentContext context)
            => Text.Add(text.Text);
    }
}
