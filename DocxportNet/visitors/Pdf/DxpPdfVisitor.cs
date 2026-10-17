using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using Microsoft.Extensions.Logging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.Text;
using System.Text.RegularExpressions;

namespace DocxportNet.Visitors.Pdf;

/// <summary>
/// A basic, text-focused PDF exporter. PDFsharp requires the host application to
/// configure a font resolver before exporting on platforms without built-in fonts.
/// </summary>
public sealed class DxpPdfVisitor : DxpVisitor, IDisposable
{
    private readonly string _fontFamily;
    private readonly double _fontSize;
    private readonly StringBuilder _paragraph = new();
    private Stream? _output;
    private PdfDocument? _document;
    private XGraphics? _graphics;
    private XFont? _font;
    private PdfPage? _page;
    private double _y;
    private bool _saved;

    public DxpPdfVisitor(string fontFamily = "Arial", double fontSize = 11, ILogger? logger = null)
        : base(logger)
    {
        if (string.IsNullOrWhiteSpace(fontFamily)) throw new ArgumentException("A font family is required.", nameof(fontFamily));
        if (fontSize <= 0) throw new ArgumentOutOfRangeException(nameof(fontSize));
        _fontFamily = fontFamily;
        _fontSize = fontSize;
    }

    public override void SetOutput(Stream stream)
    {
        if (_document != null) throw new InvalidOperationException("The visitor has already been used.");
        _output = stream ?? throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The PDF output stream must be writable.", nameof(stream));
    }

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext context)
    {
        if (_output == null) throw new InvalidOperationException("SetOutput must be called before walking the document.");
        _document = new PdfDocument();
        _font = new XFont(_fontFamily, _fontSize, XFontStyleEx.Regular);
        NewPage();
        return DxpDisposable.Create(Save);
    }

    public override IDisposable VisitParagraphBegin(Paragraph paragraph, DxpIDocumentContext context, DxpIParagraphContext paragraphContext)
    {
        _paragraph.Clear();
        if (paragraphContext.MarkerAccept?.marker is string marker)
            _paragraph.Append(NormalizeMarker(marker)).Append(' ');
        return DxpDisposable.Create(FlushParagraph);
    }

    public override void VisitText(Text text, DxpIDocumentContext context) => _paragraph.Append(text.Text);

    public override void VisitTab(TabChar tab, DxpIDocumentContext context) => _paragraph.Append("    ");

    public override void VisitBreak(Break br, DxpIDocumentContext context)
    {
        if (br.Type?.Value == BreakValues.Page)
        {
            FlushParagraph();
            NewPage();
        }
        else
        {
            _paragraph.Append('\n');
        }
    }

    internal static string NormalizeMarker(string marker)
    {
        var font = Regex.Match(marker, "font-family\\s*:\\s*([^;\\\">]+)", RegexOptions.IgnoreCase);
        var family = font.Success ? font.Groups[1].Value.Trim().Trim('"', '\'') : null;
        var text = System.Net.WebUtility.HtmlDecode(Regex.Replace(marker, "<[^>]*>", "")).Trim();
        if (text.Length == 1)
            text = DxpFontSymbols.Substitute(family, text);
        return text == "" || text == "·" ? "•" : text;
    }

    private void FlushParagraph()
    {
        if (_graphics == null || _font == null || _page == null) return;
        var availableWidth = _page.Width.Point - 100;
        foreach (var sourceLine in _paragraph.ToString().Split('\n'))
        {
            var line = new StringBuilder();
            foreach (var word in sourceLine.Split(' '))
            {
                var candidate = line.Length == 0 ? word : line + " " + word;
                if (line.Length > 0 && _graphics.MeasureString(candidate, _font).Width > availableWidth)
                {
                    DrawLine(line.ToString());
                    line.Clear();
                    line.Append(word);
                }
                else
                {
                    line.Clear();
                    line.Append(candidate);
                }
            }
            DrawLine(line.ToString());
        }
        _y += _fontSize * 0.5;
        _paragraph.Clear();
    }

    private void DrawLine(string text)
    {
        if (_page == null || _font == null) return;
        if (_y + _fontSize * 1.4 > _page.Height.Point - 50) NewPage();
        _graphics!.DrawString(text, _font, XBrushes.Black, new XPoint(50, _y + _fontSize));
        _y += _fontSize * 1.4;
    }

    private void NewPage()
    {
        _graphics?.Dispose();
        _page = _document!.AddPage();
        _page.Size = PdfSharp.PageSize.A4;
        _graphics = XGraphics.FromPdfPage(_page);
        _y = 50;
    }

    private void Save()
    {
        if (_saved || _document == null || _output == null) return;
        _graphics?.Dispose();
        _graphics = null;
        _document.Save(_output, false);
        _saved = true;
    }

    public void Dispose()
    {
        _graphics?.Dispose();
        _document?.Dispose();
    }
}
