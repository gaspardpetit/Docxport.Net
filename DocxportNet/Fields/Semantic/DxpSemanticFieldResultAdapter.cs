using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Walker;

namespace DocxportNet.Fields.Semantic;

internal static class DxpSemanticFieldResultAdapter
{
    public static void Replay(
        DxpSemanticFieldResult result,
        DxpIVisitor visitor,
        DxpIDocumentContext context,
        Run? sourceRun = null)
    {
        if (result.Content.IsEmpty)
            return;
        BuildBuffer(result.Content, sourceRun).Replay(visitor, context);
    }

    internal static DxpFieldNodeBuffer BuildBuffer(
        DxpSemanticContent content,
        Run? sourceRun = null)
    {
        var root = new DxpFieldNodeBuffer();
        DxpFieldNodeBuffer? inlineRoot = null;
        DxpFieldNodeBuffer? inlineTarget = null;

        void FlushInline()
        {
            if (inlineRoot == null || inlineTarget == null || inlineTarget.IsEmpty)
                return;
            root.Append(inlineRoot);
            inlineRoot = null;
            inlineTarget = null;
        }

        foreach (DxpSemanticNode node in content.Nodes)
        {
            if (node is DxpSemanticParagraph paragraph)
            {
                FlushInline();
                root.Append(DxpFieldNodeBuffer.FromBlock(BuildParagraph(paragraph.Content, sourceRun)));
                continue;
            }
            if (node is DxpSemanticTable table)
            {
                FlushInline();
                root.Append(DxpFieldNodeBuffer.FromBlock(BuildTable(table, sourceRun)));
                continue;
            }

            if (inlineRoot == null)
                (inlineRoot, inlineTarget) = NewInlineBuffer(sourceRun);
            AppendInlineNode(inlineTarget!, node);
        }

        FlushInline();
        return root;
    }

    private static (DxpFieldNodeBuffer Root, DxpFieldNodeBuffer Target) NewInlineBuffer(Run? sourceRun)
    {
        Run run = sourceRun == null
            ? new Run()
            : DxpRunCloner.CloneRunWithParagraphAncestor(sourceRun);
        var root = new DxpFieldNodeBuffer();
        return (root, root.BeginRun(run));
    }

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

    private static Paragraph BuildParagraph(DxpSemanticContent content, Run? sourceRun)
    {
        var paragraph = sourceRun?.Ancestors<Paragraph>().FirstOrDefault() is Paragraph sourceParagraph
            ? (Paragraph)sourceParagraph.CloneNode(false)
            : new Paragraph();
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
        var run = sourceRun == null ? new Run() : (Run)sourceRun.CloneNode(false);
        if (sourceRun?.RunProperties != null && run.RunProperties == null)
            run.RunProperties = (RunProperties)sourceRun.RunProperties.CloneNode(true);
        foreach (DxpSemanticNode node in content.Nodes)
        {
            switch (node)
            {
                case DxpSemanticText text:
                    var value = new Text(text.Text);
                    if (text.Text.Length > 0 &&
                        (char.IsWhiteSpace(text.Text[0]) || char.IsWhiteSpace(text.Text[^1])))
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
        }
        yield return run;
    }

    private static T NewBorder<T>() where T : BorderType, new()
        => new() { Val = BorderValues.Single, Size = 4U };
}
