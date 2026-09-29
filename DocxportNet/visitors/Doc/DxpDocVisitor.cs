using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Doc;
using DocxportNet.Fields;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Visitors.Doc;

/// <summary>Collects main-story plain text from a DOCX walk and writes a binary DOC.</summary>
public sealed class DxpDocVisitor : DxpVisitor, DxpIFieldEvalProvider
{
    private readonly StringBuilder _text = new();
    private Stream? _output;
    private MainDocumentPart? _mainPart;
    private bool _inMainParagraph;

    public DxpFieldEval FieldEval { get; }

    public DxpDocVisitor(ILogger? logger = null, DxpFieldEval? fieldEval = null) : base(logger)
        => FieldEval = fieldEval ?? new DxpFieldEval(logger: logger);

    public override void SetOutput(Stream stream) =>
        _output = stream ?? throw new ArgumentNullException(nameof(stream));

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext context)
    {
        if (_output == null)
            throw new InvalidOperationException("An output stream must be assigned before walking the document.");
        _text.Clear();
        _mainPart = doc.MainDocumentPart;
        return DxpDisposable.Create(() => DocPlainTextWriter.Write(_output, _text.ToString()));
    }

    public override IDisposable VisitParagraphBegin(Paragraph paragraph, DxpIDocumentContext context,
        DxpIParagraphContext paragraphContext)
    {
        if (context.CurrentPart != _mainPart) return DxpDisposable.Empty;
        var previous = _inMainParagraph;
        _inMainParagraph = true;
        return DxpDisposable.Create(() =>
        {
            _text.Append('\r');
            _inMainParagraph = previous;
        });
    }

    public override void VisitText(Text text, DxpIDocumentContext context)
    {
        if (_inMainParagraph) _text.Append(text.Text);
    }

    public override void VisitTab(TabChar tab, DxpIDocumentContext context)
    {
        if (_inMainParagraph) _text.Append('\t');
    }

    public override void VisitBreak(Break lineBreak, DxpIDocumentContext context)
    {
        if (_inMainParagraph) _text.Append(lineBreak.Type?.Value == BreakValues.Page ? '\f' : '\v');
    }

    public override void VisitCarriageReturn(CarriageReturn carriageReturn, DxpIDocumentContext context)
    {
        if (_inMainParagraph) _text.Append('\v');
    }

    public override void VisitNoBreakHyphen(NoBreakHyphen hyphen, DxpIDocumentContext context)
    {
        if (_inMainParagraph) _text.Append('\u001E');
    }

    public override void VisitSoftHyphen(SoftHyphen hyphen, DxpIDocumentContext context)
    {
        if (_inMainParagraph) _text.Append('\u001F');
    }
}
