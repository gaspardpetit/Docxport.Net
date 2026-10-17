using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using Microsoft.Extensions.Logging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
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
    private readonly List<StyledFragment> _paragraph = new();
    private readonly Dictionary<(string Family, double Size, XFontStyleEx Style), XFont> _fonts = new();
    private Stream? _output;
    private PdfDocument? _document;
    private XGraphics? _graphics;
    private XFont? _font;
    private PdfPage? _page;
    private double _y;
    private bool _saved;

    private sealed record StyledFragment(string Text, XFont Font);

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
        _fonts.Add((_fontFamily, _fontSize, XFontStyleEx.Regular), _font);
        NewPage();
        return DxpDisposable.Create(Save);
    }

    public override IDisposable VisitParagraphBegin(Paragraph paragraph, DxpIDocumentContext context, DxpIParagraphContext paragraphContext)
    {
        _paragraph.Clear();
        if (paragraphContext.MarkerAccept?.marker is string marker)
            _paragraph.Add(new StyledFragment(NormalizeMarker(marker) + " ", _font!));
        return DxpDisposable.Create(FlushParagraph);
    }

    public override void VisitText(Text text, DxpIDocumentContext context)
    {
        if (string.IsNullOrEmpty(text.Text)) return;
        var style = context.CurrentRun?.Style;
        string family = style?.FontName ?? _fontFamily;
        double size = style?.FontSizeHalfPoints is int halfPoints && halfPoints > 0
            ? halfPoints / 2.0 : _fontSize;
        var fontStyle = (style?.Bold == true ? XFontStyleEx.Bold : XFontStyleEx.Regular) |
            (style?.Italic == true ? XFontStyleEx.Italic : XFontStyleEx.Regular);
        var key = (family, size, fontStyle);
        if (!_fonts.TryGetValue(key, out var font))
            _fonts.Add(key, font = new XFont(family, size, fontStyle));
        _paragraph.Add(new StyledFragment(text.Text, font));
    }

    public override void VisitTab(TabChar tab, DxpIDocumentContext context) =>
        _paragraph.Add(new StyledFragment("    ", _font!));

    public override void VisitBreak(Break br, DxpIDocumentContext context)
    {
        if (br.Type?.Value == BreakValues.Page)
        {
            FlushParagraph();
            NewPage();
        }
        else
        {
            _paragraph.Add(new StyledFragment("\n", _font!));
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
        double availableWidth = _page.Width.Point - 100;
        var line = new List<StyledFragment>();
        var word = new List<StyledFragment>();
        var spaces = new List<StyledFragment>();
        double lineWidth = 0;

        void AddWord()
        {
            if (word.Count == 0) return;
            double wordWidth = Measure(word);
            double spaceWidth = line.Count == 0 ? 0 : Measure(spaces);
            if (line.Count > 0 && lineWidth + spaceWidth + wordWidth > availableWidth)
            {
                DrawLine(line);
                line.Clear();
                lineWidth = 0;
                spaceWidth = 0;
            }
            if (spaceWidth > 0)
                line.AddRange(spaces);
            line.AddRange(word);
            lineWidth += spaceWidth + wordWidth;
            word.Clear();
            spaces.Clear();
        }

        foreach (var fragment in _paragraph)
        {
            foreach (Match token in Regex.Matches(fragment.Text, "\\n| +|[^ \\n]+"))
            {
                if (token.Value == "\n")
                {
                    AddWord();
                    DrawLine(line);
                    line.Clear();
                    lineWidth = 0;
                    spaces.Clear();
                }
                else if (token.Value[0] == ' ')
                {
                    AddWord();
                    spaces.Add(new StyledFragment(token.Value, fragment.Font));
                }
                else
                {
                    word.Add(new StyledFragment(token.Value, fragment.Font));
                }
            }
        }
        AddWord();
        DrawLine(line);
        _y += _fontSize * 0.5;
        _paragraph.Clear();
    }

    private double Measure(IEnumerable<StyledFragment> fragments) =>
        fragments.Sum(fragment => _graphics!.MeasureString(fragment.Text, fragment.Font).Width);

    private void DrawLine(IReadOnlyList<StyledFragment> line)
    {
        if (_page == null || _font == null) return;
        double fontSize = line.Count == 0 ? _fontSize : line.Max(fragment => fragment.Font.Size);
        if (_y + fontSize * 1.4 > _page.Height.Point - 50) NewPage();
        double x = 50;
        foreach (var fragment in line)
        {
            _graphics!.DrawString(fragment.Text, fragment.Font, XBrushes.Black, new XPoint(x, _y + fontSize));
            x += _graphics.MeasureString(fragment.Text, fragment.Font).Width;
        }
        _y += fontSize * 1.4;
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
