using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Semantic;

internal static class DxpSemanticFieldResultAdapter
{
    public static void Replay(
        DxpSemanticFieldResult result,
        DxpIVisitor visitor,
        DxpIDocumentContext context,
        Run? sourceRun = null,
        DxpFieldEval? eval = null,
        ILogger? logger = null)
    {
        if (result.Content.IsEmpty)
            return;
        BuildBuffer(result.Content, sourceRun, eval, logger).Replay(visitor, context);
    }

    internal static DxpFieldNodeBuffer BuildBuffer(
        DxpSemanticContent content,
        Run? sourceRun = null,
        DxpFieldEval? eval = null,
        ILogger? logger = null)
    {
        var root = new DxpFieldNodeBuffer();

        foreach (DxpSemanticNode node in content.Nodes)
        {
            if (node is DxpSemanticParagraph paragraph)
            {
                root.Append(DxpFieldNodeBuffer.FromBlock(BuildParagraph(
                    paragraph.Content, sourceRun, paragraph.Format)));
                continue;
            }
            if (node is DxpSemanticTable table)
            {
                root.Append(DxpFieldNodeBuffer.FromBlock(BuildTable(table, sourceRun)));
                continue;
            }
            if (node is DxpSemanticInclude include)
            {
                if (eval != null)
                {
                    root.AddIncludeTextExpansion(new DxpIncludeTextExpansion(
                        include.Path,
                        include.Identity,
                        include.Content,
                        include.Bookmark,
                        CachedResult: null,
                        eval,
                        logger));
                }
                continue;
            }

            Run run = NewRun(sourceRun, GetFormat(node));
            DxpFieldNodeBuffer target = root.BeginRun(run);
            AppendInlineNode(target, node);
        }
        return root;
    }

    private static Run NewRun(Run? sourceRun, DxpSemanticRunFormat? format)
    {
        Run run = sourceRun == null
            ? new Run()
            : DxpRunCloner.CloneRunWithParagraphAncestor(sourceRun);
        ApplyFormat(run, format);
        return run;
    }

    private static DxpSemanticRunFormat? GetFormat(DxpSemanticNode node) => node switch
    {
        DxpSemanticText text => text.Format,
        DxpSemanticBreak lineBreak => lineBreak.Format,
        DxpSemanticTab tab => tab.Format,
        _ => null
    };

    private static void AppendInlineNode(DxpFieldNodeBuffer buffer, DxpSemanticNode node)
    {
        switch (node)
        {
            case DxpSemanticText text:
                buffer.AddText(text.Text);
                break;
            case DxpSemanticBreak:
                buffer.AddBreak();
                break;
            case DxpSemanticTab:
                buffer.AddTab();
                break;
        }
    }

    private static Paragraph BuildParagraph(
        DxpSemanticContent content,
        Run? sourceRun,
        DxpSemanticParagraphFormat? format = null)
    {
        var paragraph = sourceRun?.Ancestors<Paragraph>().FirstOrDefault() is Paragraph sourceParagraph
            ? (Paragraph)sourceParagraph.CloneNode(false)
            : new Paragraph();
        ApplyFormat(paragraph, format);
        paragraph.Append(BuildInlineElements(content, sourceRun));
        return paragraph;
    }

    private static Table BuildTable(DxpSemanticTable semantic, Run? sourceRun)
    {
        int columnCount = semantic.Rows.Select(static row => row.Cells.Count).DefaultIfEmpty().Max();
        var properties = new TableProperties(
            new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto });
        if (semantic.AutoFit)
            properties.AppendChild(new TableLayout { Type = TableLayoutValues.Autofit });
        if (semantic.ShowBorders)
        {
            properties.AppendChild(new TableBorders(
                NewBorder<TopBorder>(), NewBorder<LeftBorder>(),
                NewBorder<BottomBorder>(), NewBorder<RightBorder>(),
                NewBorder<InsideHorizontalBorder>(), NewBorder<InsideVerticalBorder>()));
        }

        var table = new Table(properties, new TableGrid(
            Enumerable.Range(0, columnCount).Select(static _ => new GridColumn())));
        for (int rowIndex = 0; rowIndex < semantic.Rows.Count; rowIndex++)
        {
            DxpSemanticTableRow semanticRow = semantic.Rows[rowIndex];
            var cells = semanticRow.Cells.Select(cell =>
                new TableCell(BuildParagraph(cell.Content, sourceRun))).ToList();
            while (cells.Count < columnCount)
                cells.Add(new TableCell(new Paragraph(new Run(new Text()))));
            var row = new TableRow(cells);
            if (semantic.HasHeader && rowIndex == 0)
                row.TableRowProperties = new TableRowProperties(new TableHeader());
            table.AppendChild(row);
        }
        return table;
    }

    private static IEnumerable<OpenXmlElement> BuildInlineElements(
        DxpSemanticContent content,
        Run? sourceRun)
    {
        foreach (DxpSemanticNode node in content.Nodes)
        {
            var run = sourceRun == null ? new Run() : (Run)sourceRun.CloneNode(false);
            if (sourceRun?.RunProperties != null && run.RunProperties == null)
                run.RunProperties = (RunProperties)sourceRun.RunProperties.CloneNode(true);
            ApplyFormat(run, GetFormat(node));
            switch (node)
            {
                case DxpSemanticText text:
                    var value = new Text(text.Text);
                    if (text.Text.Length > 0 &&
                        (char.IsWhiteSpace(text.Text[0]) || char.IsWhiteSpace(text.Text[text.Text.Length - 1])))
                        value.Space = SpaceProcessingModeValues.Preserve;
                    run.AppendChild(value);
                    break;
                case DxpSemanticBreak:
                    run.AppendChild(new Break());
                    break;
                case DxpSemanticTab:
                    run.AppendChild(new TabChar());
                    break;
            }
            if (run.ChildElements.Count > 0)
                yield return run;
        }
    }

    private static void ApplyFormat(Run run, DxpSemanticRunFormat? format)
    {
        if (format == null)
            return;
        RunProperties properties = run.RunProperties ??= new RunProperties();
        if (format.Bold.HasValue)
            properties.Bold = new Bold { Val = format.Bold.Value };
        if (format.Italic.HasValue)
            properties.Italic = new Italic { Val = format.Italic.Value };
        if (format.Strike.HasValue)
            properties.Strike = new Strike { Val = format.Strike.Value };
        if (format.Underline != null)
            properties.Underline = new Underline { Val = ParseUnderline(format.Underline) };
        if (format.Color != null)
            properties.Color = new Color { Val = format.Color };
        if (format.FontSizeHalfPoints != null)
            properties.FontSize = new FontSize { Val = format.FontSizeHalfPoints };
        if (format.StyleId != null)
            properties.RunStyle = new RunStyle { Val = format.StyleId };
        if (format.Language != null)
            properties.Languages = new Languages { Val = format.Language };
    }

    private static void ApplyFormat(Paragraph paragraph, DxpSemanticParagraphFormat? format)
    {
        if (format == null)
            return;
        ParagraphProperties properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
        if (format.StyleId != null)
            properties.ParagraphStyleId = new ParagraphStyleId { Val = format.StyleId };
        if (format.Alignment != null)
            properties.Justification = new Justification { Val = ParseJustification(format.Alignment) };
        if (format.OutlineLevel.HasValue)
            properties.OutlineLevel = new OutlineLevel { Val = format.OutlineLevel.Value };
        if (format.NumberingId.HasValue || format.NumberingLevel.HasValue)
        {
            properties.NumberingProperties = new NumberingProperties();
            if (format.NumberingLevel.HasValue)
                properties.NumberingProperties.NumberingLevelReference =
                    new NumberingLevelReference { Val = format.NumberingLevel.Value };
            if (format.NumberingId.HasValue)
                properties.NumberingProperties.NumberingId =
                    new NumberingId { Val = format.NumberingId.Value };
        }
    }

    private static UnderlineValues ParseUnderline(string value)
    {
        switch (value.ToLowerInvariant())
        {
            case "none": return UnderlineValues.None;
            case "double": return UnderlineValues.Double;
            case "dotted": return UnderlineValues.Dotted;
            case "dash": return UnderlineValues.Dash;
            case "wave": return UnderlineValues.Wave;
            case "words": return UnderlineValues.Words;
            default: return UnderlineValues.Single;
        }
    }

    private static JustificationValues ParseJustification(string value)
    {
        switch (value.ToLowerInvariant())
        {
            case "right": return JustificationValues.Right;
            case "center": return JustificationValues.Center;
            case "both": return JustificationValues.Both;
            case "distribute": return JustificationValues.Distribute;
            default: return JustificationValues.Left;
        }
    }

    private static T NewBorder<T>() where T : BorderType, new()
        => new() { Val = BorderValues.Single, Size = 4U };
}
