using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Word;
using Microsoft.Extensions.Logging;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocxportNet.Fields;

namespace DocxportNet.Visitors.PlainText;

public enum DxpPlainTextTrackedChangeMode
{
    AcceptChanges,
    RejectChanges
}

public sealed record DxpPlainTextVisitorConfig
{
    public DxpPlainTextTrackedChangeMode TrackedChangeMode = DxpPlainTextTrackedChangeMode.AcceptChanges;
    public string ImagePlaceholder = "[IMAGE]";
    public bool EmitDocumentProperties = true;
    public bool EmitCustomProperties = true;
    public static DxpPlainTextVisitorConfig CreateAcceptConfig() => new();
    public static DxpPlainTextVisitorConfig CreateRejectConfig() => new() { TrackedChangeMode = DxpPlainTextTrackedChangeMode.RejectChanges };
}

public sealed class DxpPlainTextVisitor : DxpVisitor, DxpITextVisitor, IDisposable, DxpIFieldEvalProvider
{
    private TextWriter _sinkWriter;
    private StreamWriter? _ownedStreamWriter;
    private readonly DxpPlainTextVisitorConfig _config;
    private DxpPlainTextVisitorState _state = new();
    private readonly DxpFieldEval _fieldEval;

    public DxpFieldEval FieldEval => _fieldEval;

    public DxpPlainTextVisitor(TextWriter writer, DxpPlainTextVisitorConfig config, ILogger? logger, DxpFieldEval? fieldEval = null)
        : base(logger)
    {
        _sinkWriter = writer;
        _config = config;
        _state.WriterStack.Push(writer);
        _fieldEval = fieldEval ?? new DxpFieldEval(logger: logger);

        if (_config.TrackedChangeMode != DxpPlainTextTrackedChangeMode.AcceptChanges &&
            _config.TrackedChangeMode != DxpPlainTextTrackedChangeMode.RejectChanges)
        {
            throw new ArgumentOutOfRangeException(nameof(config), "Plain text visitor only supports accept or reject tracked changes.");
        }
    }

    public DxpPlainTextVisitor(DxpPlainTextVisitorConfig config, ILogger? logger = null, DxpFieldEval? fieldEval = null)
        : this(TextWriter.Null, config, logger, fieldEval)
    {
    }

    public void SetOutput(TextWriter writer)
    {
        ReleaseOwnedWriter();
        _sinkWriter = writer ?? throw new ArgumentNullException(nameof(writer));
        _state = new DxpPlainTextVisitorState();
        _state.WriterStack.Push(_sinkWriter);
    }

    public override void SetOutput(Stream stream)
    {
        ReleaseOwnedWriter();
        _ownedStreamWriter = new StreamWriter(stream, Encoding.UTF8, bufferSize: 1024, leaveOpen: true) {
            AutoFlush = true
        };
        var writer = _ownedStreamWriter;
        SetOutput(writer);
    }

    private TextWriter CurrentWriter => _state.WriterStack.Peek();

    private void ReleaseOwnedWriter()
    {
        if (_ownedStreamWriter == null)
            return;

        _ownedStreamWriter.Flush();
        _ownedStreamWriter.Dispose();
        _ownedStreamWriter = null;
    }

    public void Dispose() => ReleaseOwnedWriter();

    private bool ShouldEmit(DxpIDocumentContext d)
    {
        return _config.TrackedChangeMode switch {
            DxpPlainTextTrackedChangeMode.AcceptChanges => d.KeepAccept,
            DxpPlainTextTrackedChangeMode.RejectChanges => d.KeepReject,
            _ => false
        };
    }

    private void Write(string text, DxpIDocumentContext d)
    {
        if (_state.SuppressDepth > 0 || _state.SuppressFieldDepth > 0)
            return;

        if (!ShouldEmit(d))
            return;

        if (_state.AllCaps)
        {
            text = text.ToUpper(CultureInfo.InvariantCulture);
        }

        if (_state.CurrentParagraph != null)
            _state.CurrentParagraph.Builder.Append(text);
        else
            CurrentWriter.Write(text);
    }

    private void WriteDirect(string text)
    {
        if (_state.SuppressDepth > 0)
            return;
        CurrentWriter.Write(text);
    }

    private void WriteDirectLine(string text)
    {
        WriteDirect(text);
        WriteDirect("\n");
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        Write(t.Text, d);
    }

    public override void VisitDeletedText(DeletedText dt, DxpIDocumentContext d)
    {
        Write(dt.Text, d);
    }

    public override void StyleAllCapsBegin(DxpIDocumentContext d) => _state.AllCaps = true;
    public override void StyleAllCapsEnd(DxpIDocumentContext d) => _state.AllCaps = false;
    public override void StyleSmallCapsBegin(DxpIDocumentContext d) => _state.AllCaps = true;
    public override void StyleSmallCapsEnd(DxpIDocumentContext d) => _state.AllCaps = false;

    public override void StyleBoldBegin(DxpIDocumentContext d) { }
    public override void StyleBoldEnd(DxpIDocumentContext d) { }
    public override void StyleItalicBegin(DxpIDocumentContext d) { }
    public override void StyleItalicEnd(DxpIDocumentContext d) { }
    public override void StyleUnderlineBegin(DxpIDocumentContext d) { }
    public override void StyleUnderlineEnd(DxpIDocumentContext d) { }
    public override void StyleStrikeBegin(DxpIDocumentContext d) { }
    public override void StyleStrikeEnd(DxpIDocumentContext d) { }
    public override void StyleDoubleStrikeBegin(DxpIDocumentContext d) { }
    public override void StyleDoubleStrikeEnd(DxpIDocumentContext d) { }
    public override void StyleSuperscriptBegin(DxpIDocumentContext d) { }
    public override void StyleSuperscriptEnd(DxpIDocumentContext d) { }
    public override void StyleSubscriptBegin(DxpIDocumentContext d) { }
    public override void StyleSubscriptEnd(DxpIDocumentContext d) { }
    public override void StyleFontBegin(DxpFont font, DxpIDocumentContext d) { }
    public override void StyleFontEnd(DxpIDocumentContext d) { }

    public override void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d) => Write("-", d);
    public override void VisitBreak(Break br, DxpIDocumentContext d) => Write("\n", d);
    public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d) => Write("\n", d);
    public override void VisitTab(TabChar tab, DxpIDocumentContext d) => Write("\t", d);

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
    {
        if (target?.uri is string href && !string.IsNullOrEmpty(href))
            return DxpDisposable.Create(() => Write($" ({href})", d));
        return DxpDisposable.Empty;
    }

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext d)
    {
        if (!_config.EmitDocumentProperties || _state.SuppressDepth > 0)
            return DxpDisposable.Empty;

        var lines = new List<string>();

        void Add(string label, string? value)
        {
            if (!string.IsNullOrWhiteSpace(value))
                lines.Add($"{label}: {value}");
        }

        IPackageProperties? core = d.DocumentProperties.PackageProperties;
        if (core != null)
        {
            Add("Title", core.Title);
            Add("Subject", core.Subject);
            Add("Author", core.Creator);
            Add("Description", core.Description);
            Add("Category", core.Category);
            Add("Keywords", core.Keywords);
            Add("LastModifiedBy", core.LastModifiedBy);
            Add("Revision", core.Revision);
            Add("Created", FormatDateUtc(core.Created));
            Add("Modified", FormatDateUtc(core.Modified));
        }

        IReadOnlyList<CustomFileProperty>? custom = d.DocumentProperties.CustomFileProperties;
        if (custom != null && _config.EmitCustomProperties)
        {
            foreach (var prop in custom)
            {
                if (prop.Value != null)
                    lines.Add($"{prop.Name}: {prop.Value}");
            }
        }

        foreach (var line in lines)
            WriteDirectLine(line);

        if (lines.Count > 0)
            WriteDirectLine(string.Empty);

        return DxpDisposable.Empty;
    }

    static string? FormatDateUtc(DateTime? value)
    {
        if (value == null)
            return null;
        return value.Value.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss'Z'");
    }

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        bool emit = ShouldEmit(d);
        var indent = paragraph.Indent;

        var marker = _config.TrackedChangeMode == DxpPlainTextTrackedChangeMode.AcceptChanges
            ? paragraph.MarkerAccept
            : paragraph.MarkerReject;

        string markerText = marker?.marker != null
            ? NormalizeMarker(marker.marker)
            : string.Empty;

        if (!string.IsNullOrEmpty(markerText))
            markerText += " ";

        double adjustedMargin = indent.Left.HasValue ? AdjustMarginLeft(DxpTwipValue.ToPoints(indent.Left.Value), d) : 0.0;
        int indentSpaces = adjustedMargin > 0.0 ? Math.Min(40, (int)Math.Round(adjustedMargin / 4.0)) : 0;
        string indentPrefix = indentSpaces > 0 ? new string(' ', indentSpaces) : string.Empty;

        var styleChain = d.Styles.GetParagraphStyleChain(p);
        var justification = p.ParagraphProperties?.Justification?.Val?.Value;
        var headingLevel = d.Styles.GetHeadingLevel(p);
        bool hasNumbering = marker?.numId != null;
        bool isHeading = headingLevel != null && !hasNumbering;
        if (!isHeading && !hasNumbering && styleChain.Any(sc => string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleSubtitle, StringComparison.OrdinalIgnoreCase)))
        {
            headingLevel = 2;
            isHeading = true;
        }

        bool isBlockQuote = styleChain.Any(sc =>
            string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleQuote, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleIntenseQuote, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleBlockQuotation, StringComparison.OrdinalIgnoreCase));

        bool isCode = styleChain.Any(sc =>
            string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleHtmlPre, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleHtmlCode, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(sc.StyleId, "Code", StringComparison.OrdinalIgnoreCase));

        bool inTable = _state.TableStack.Count > 0;

        string prefix = indentPrefix + markerText;
        if (isBlockQuote)
            prefix = "> " + prefix;
        if (isCode)
            prefix += "    ";
        if (isHeading)
            prefix = indentPrefix; // headings ignore markers

        var paragraphState = new ParagraphBuffer(
            emit: emit,
            prefix: prefix,
            headingLevel: isHeading ? headingLevel : null,
            separator: inTable ? "\n" : "\n\n");

        _state.CurrentParagraph = paragraphState;

        return DxpDisposable.Create(() => FlushParagraph(paragraphState));
    }

    private void FlushParagraph(ParagraphBuffer paragraph)
    {
        _state.CurrentParagraph = null;

        if (_state.SuppressDepth > 0)
            return;
        if (!paragraph.Emit)
            return;

        string content = paragraph.Builder.ToString().Replace("\r\n", "\n");
        content = content.TrimEnd('\n');

        if (string.IsNullOrEmpty(content))
        {
            WriteDirect(paragraph.Separator);
            return;
        }

        content = PrefixLines(content, paragraph.Prefix);

        if (paragraph.HeadingLevel != null)
            content = FormatHeading(content.Trim(), paragraph.HeadingLevel.Value);

        WriteDirect(content);
        WriteDirect(paragraph.Separator);
    }

    private static string PrefixLines(string text, string prefix)
    {
        if (string.IsNullOrEmpty(prefix))
            return text;

        var lines = text.Split('\n');
        var sb = new StringBuilder();
        for (int i = 0; i < lines.Length; i++)
        {
            if (i > 0)
                sb.Append('\n');
            sb.Append(prefix);
            sb.Append(lines[i]);
        }
        return sb.ToString();
    }

    private static string FormatHeading(string text, int level)
    {
        char underlineChar = level == 1 ? '=' : '-';
        string underline = new string(underlineChar, Math.Max(text.Length, 3));
        return $"{text}\n{underline}";
    }

    public override void VisitFootnoteReference(FootnoteReference fr, DxpIFootnoteContext footnote, DxpIDocumentContext d)
    {
        if (footnote.Index != null)
            Write($"[{footnote.Index}]", d);
    }

    public override IDisposable VisitFootnoteBegin(Footnote fn, DxpIFootnoteContext footnote, DxpIDocumentContext d)
    {
        WriteDirect($"\n[FOOTNOTE {footnote.Index ?? footnote.Id}]\n");
        return DxpDisposable.Create(() => WriteDirect("\n"));
    }

    public override IDisposable VisitSectionHeaderBegin(Header hdr, object kind, DxpIDocumentContext d)
    {
        _state.SuppressDepth++;
        return DxpDisposable.Create(() => _state.SuppressDepth--);
    }

    public override IDisposable VisitSectionFooterBegin(Footer ftr, object kind, DxpIDocumentContext d)
    {
        _state.SuppressDepth++;
        return DxpDisposable.Create(() => _state.SuppressDepth--);
    }

    public override void VisitPageNumber(PageNumber pn, DxpIDocumentContext d)
    {
        // Ignored in plain text output.
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        // Keep plain text output focused on visible text; skip field instructions.
    }

    public override IDisposable VisitComplexFieldResultBegin(DxpIDocumentContext d) => DxpDisposable.Empty;
    public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
    {
    }
    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        Write(text, d);
    }
    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        return DxpDisposable.Empty;
    }

    public override IDisposable VisitTableBegin(Table t, DxpTableModel model, DxpIDocumentContext d, DxpITableContext table)
    {
        var tableState = new TableState {
            Emit = ShouldEmit(d) && _state.SuppressDepth == 0
        };
        _state.TableStack.Push(tableState);
        return DxpDisposable.Create(() => FlushTable());
    }

    public override IDisposable VisitTableRowBegin(TableRow tr, DxpITableRowContext row, DxpIDocumentContext d)
    {
        var current = _state.TableStack.Peek();
        var rowBuffer = new TableRowBuffer(row.IsHeader);
        current.Rows.Add(rowBuffer);
        current.CurrentRow = rowBuffer;
        return DxpDisposable.Create(() => current.CurrentRow = null);
    }

    public override IDisposable VisitTableCellBegin(TableCell tc, DxpITableCellContext cell, DxpIDocumentContext d)
    {
        var tableState = _state.TableStack.Peek();
        var cellWriter = new DxpBufferedTextWriter();
        _state.WriterStack.Push(cellWriter);

        return DxpDisposable.Create(() => {
            _state.WriterStack.Pop();
            string raw = cellWriter.Drain();
            string normalized = NormalizeCell(raw);
            tableState.CurrentRow?.Cells.Add(normalized);
        });
    }

    private void FlushTable()
    {
        var tableState = _state.TableStack.Pop();
        if (!tableState.Emit || _state.SuppressDepth > 0)
            return;

        string rendered = RenderTable(tableState);
        if (string.IsNullOrWhiteSpace(rendered))
            return;

        WriteDirect(rendered);
        WriteDirect("\n\n");
    }

    private static string NormalizeCell(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return string.Empty;

        var parts = raw.Replace("\r\n", "\n").Split('\n');
        for (int i = 0; i < parts.Length; i++)
            parts[i] = parts[i].Trim();
        return string.Join(" ", parts.Where(p => p.Length > 0));
    }

    private static string RenderTable(TableState table)
    {
        if (table.Rows.Count == 0)
            return string.Empty;

        int columnCount = table.Rows.Max(r => r.Cells.Count);
        if (columnCount == 0)
            return string.Empty;

        foreach (var row in table.Rows)
        {
            while (row.Cells.Count < columnCount)
                row.Cells.Add(string.Empty);
        }

        var widths = new int[columnCount];
        foreach (var row in table.Rows)
        {
            for (int i = 0; i < columnCount; i++)
                widths[i] = Math.Max(widths[i], row.Cells[i].Length);
        }

        var sb = new StringBuilder();
        for (int r = 0; r < table.Rows.Count; r++)
        {
            var row = table.Rows[r];
            var padded = new List<string>(columnCount);
            for (int c = 0; c < columnCount; c++)
            {
                var cell = row.Cells[c];
                padded.Add(cell.PadRight(widths[c]));
            }

            sb.Append(string.Join(" | ", padded));
            if (r < table.Rows.Count - 1 || row.IsHeader)
                sb.Append('\n');

            if (row.IsHeader)
            {
                var sep = widths.Select(w => new string('-', Math.Max(w, 3)));
                sb.Append(string.Join("-+-", sep));
                if (r < table.Rows.Count - 1)
                    sb.Append('\n');
            }
        }

        return sb.ToString().TrimEnd('\n');
    }

    public override void VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
    {
        string alt = info?.AltText ?? "image";
        Write($"{_config.ImagePlaceholder}: {alt}", d);
    }

    IDisposable DxpIVisitor.VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
    {
        VisitDrawingBegin(drw, info, d);
        return DxpDisposable.Empty;
    }

    public new void VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
    {
        Write($"{_config.ImagePlaceholder}: image", d);
    }

    IDisposable DxpIVisitor.VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
    {
        VisitLegacyPictureBegin(pict, d);
        return DxpDisposable.Empty;
    }

    public override IDisposable VisitCommentThreadBegin(string anchorId, DxpCommentThread thread, DxpIDocumentContext d)
    {
        if (thread.Comments == null || thread.Comments.Count == 0)
            return DxpDisposable.Empty;

        WriteDirectLine($"[COMMENTS for {anchorId}]");
        return DxpDisposable.Create(() => WriteDirectLine(string.Empty));
    }

    public override IDisposable VisitCommentBegin(DxpCommentInfo c, DxpCommentThread thread, DxpIDocumentContext d)
    {
        var who = !string.IsNullOrEmpty(c.Author)
            ? c.Author!
            : !string.IsNullOrEmpty(c.Initials) ? c.Initials! : "Unknown";
        var when = c.DateUtc?.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss'Z'");
        if (!string.IsNullOrEmpty(when))
            WriteDirectLine($"  {who} on {when}:");
        else
            WriteDirectLine($"  {who}:");
        return DxpDisposable.Create(() => WriteDirectLine(string.Empty));
    }

    public override IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d) => DxpDisposable.Empty;
    public override IDisposable VisitDocumentBodyBegin(Body body, DxpIDocumentContext d) => DxpDisposable.Empty;

    private double AdjustMarginLeft(double marginPt, DxpIDocumentContext d)
    {
        var marginLeftPoints = d.CurrentSection.Layout?.MarginLeft?.Inches is double inches
            ? inches * 72.0
            : (double?)null;

        if (marginLeftPoints == null)
            return marginPt;
        var adjusted = marginPt - marginLeftPoints.Value;
        if (adjusted < 0)
            adjusted = 0;
        return adjusted;
    }

    private string NormalizeMarker(string marker)
    {
        if (string.IsNullOrEmpty(marker))
            return marker;

        if (marker.IndexOf("<span", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            string inner = StripTags(marker).Trim();
            var font = ExtractFontFamily(marker);
            var translatedSpan = TryTranslateSymbolFont(inner, font);
            if (!string.IsNullOrEmpty(translatedSpan))
                return translatedSpan!;
            return inner;
        }

        string trimmed = marker.Trim();
        if (trimmed.Length == 1)
        {
            char c = trimmed[0];
            if (c == '\u2022' || c == '•' || c == '·' || c == '')
                return "•";
        }

        var translated = TryTranslateSymbolFont(marker);
        if (!string.IsNullOrEmpty(translated))
            return translated!;

        return trimmed;
    }

    private static string StripTags(string input)
    {
        return Regex.Replace(input, "<.*?>", string.Empty);
    }

    private static string? ExtractFontFamily(string marker)
    {
        var m = Regex.Match(marker, "font-family\\s*:\\s*([^;\">]+)", RegexOptions.IgnoreCase);
        if (!m.Success)
            return null;
        var font = m.Groups[1].Value.Trim();
        return font.Trim('"', '\'');
    }

    private static string? TryTranslateSymbolFont(string marker, string? fontFamily = null)
    {
        if (string.IsNullOrEmpty(marker) || marker.Length != 1)
            return null;

        var ch = marker[0];
        var converter = DxpFontSymbols.GetSymbolConverter(fontFamily);
        if (converter != null)
        {
            var translated = converter.Substitute(ch, null);
            if (!string.IsNullOrEmpty(translated) && !string.Equals(translated, marker, StringComparison.Ordinal))
                return translated;
        }

        return null;
    }

    private sealed class ParagraphBuffer
    {
        public ParagraphBuffer(bool emit, string prefix, int? headingLevel, string separator)
        {
            Emit = emit;
            Prefix = prefix;
            HeadingLevel = headingLevel;
            Separator = separator;
        }

        public bool Emit { get; }
        public string Prefix { get; }
        public int? HeadingLevel { get; }
        public string Separator { get; }
        public StringBuilder Builder { get; } = new();
    }

    private sealed class TableRowBuffer
    {
        public TableRowBuffer(bool isHeader)
        {
            IsHeader = isHeader;
        }
        public bool IsHeader { get; }
        public List<string> Cells { get; } = new();
    }

    private sealed class TableState
    {
        public bool Emit { get; set; }
        public List<TableRowBuffer> Rows { get; } = new();
        public TableRowBuffer? CurrentRow { get; set; }
    }

    private sealed class DxpPlainTextVisitorState
    {
        public bool AllCaps { get; set; }
        public ParagraphBuffer? CurrentParagraph { get; set; }
        public Stack<TextWriter> WriterStack { get; } = new();
        public Stack<TableState> TableStack { get; } = new();
        public int SuppressDepth { get; set; }
        public int SuppressFieldDepth { get; set; }
    }

}
