using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Middleware;

public abstract class DxpLoggingMiddleware : DxpMiddleware
{
    private readonly ILogger? _logger;
    private readonly string _name;
    private readonly LogLevel _level;

    public DxpLoggingMiddleware(ILogger? logger, string name = "DxpLogging", LogLevel level = LogLevel.Debug)
        : base()
    {
        _logger = logger;
        _name = name;
        _level = level;
    }

    protected virtual void LogEvent(string evt, DxpIDocumentContext d, string? details = null)
    {
        if (_logger == null || !_logger.IsEnabled(_level))
            return;
        if (string.IsNullOrEmpty(details))
            _logger.Log(_level, "{Name}: {Event}", _name, evt);
        else
            _logger.Log(_level, "{Name}: {Event} {Details}", _name, evt, details);
    }

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext d)
    {
        LogEvent("VisitDocumentBegin", d);
        var inner = base.VisitDocumentBegin(doc, d);
        return DxpDisposable.Create(() => {
            inner.Dispose();
            LogEvent("VisitDocumentEnd", d);
        });
    }

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        LogEvent("VisitParagraphBegin", d);
        var inner = base.VisitParagraphBegin(p, d, paragraph);
        return DxpDisposable.Create(() => {
            inner.Dispose();
            LogEvent("VisitParagraphEnd", d);
        });
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        LogEvent("VisitRunBegin", d);
        var inner = base.VisitRunBegin(r, d);
        return DxpDisposable.Create(() => {
            inner.Dispose();
            LogEvent("VisitRunEnd", d);
        });
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        LogEvent("VisitText", d, $"text='{t.Text}'");
        base.VisitText(t, d);
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        LogEvent("VisitSimpleFieldBegin", d, $"instr='{fld.Instruction?.Value ?? string.Empty}'");
        var inner = base.VisitSimpleFieldBegin(fld, d);
        return DxpDisposable.Create(() => {
            inner.Dispose();
            LogEvent("VisitSimpleFieldEnd", d);
        });
    }

    public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
    {
        LogEvent("VisitComplexFieldBegin", d);
        base.VisitComplexFieldBegin(begin, d);
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        LogEvent("VisitComplexFieldInstruction", d, $"text='{text}'");
        base.VisitComplexFieldInstruction(instr, text, d);
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        LogEvent("VisitComplexFieldSeparate", d);
        base.VisitComplexFieldSeparate(separate, d);
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        LogEvent("VisitComplexFieldCachedResultText", d, $"text='{text}'");
        base.VisitComplexFieldCachedResultText(text, d);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        LogEvent("VisitComplexFieldEnd", d);
        base.VisitComplexFieldEnd(end, d);
    }
}
