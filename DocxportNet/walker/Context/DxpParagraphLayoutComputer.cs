using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker.Context;

internal static class DxpParagraphLayoutComputer
{
    public static DxpComputedParagraphLayout? ComputeLayout(Paragraph p, DxpDocumentContext d)
    {
        // Keep it cheap: only compute layout info for paragraphs that actually contain tabs.
        if (!p.Descendants<TabChar>().Any() && !p.Descendants<PositionalTab>().Any())
            return null;

        Tabs? tabs = p.ParagraphProperties?.Tabs;
        if (tabs == null && d.Styles is DxpStyleResolver resolver)
            tabs = resolver.GetParagraphTabs(p);
        if (tabs == null)
            return new DxpComputedParagraphLayout([]);

        var stops = new List<DxpComputedTabStop>();
        foreach (var stop in tabs.Elements<TabStop>())
        {
            var kind = stop.Val?.Value;
            var pos = stop.Position?.Value;
            if (pos == null || pos.Value <= 0)
                continue;

            var k = MapKind(kind);
            if (k == null)
                continue;

            stops.Add(new DxpComputedTabStop(k.Value, DxpTwipValue.ToPoints((int)pos.Value)));
        }

        stops.Sort((a, b) => a.PositionPt.CompareTo(b.PositionPt));
        return new DxpComputedParagraphLayout(stops);
    }

    private static DxpComputedTabStopKind? MapKind(TabStopValues? kind)
    {
        if (kind == TabStopValues.Right)
            return DxpComputedTabStopKind.Right;
        if (kind == TabStopValues.Center)
            return DxpComputedTabStopKind.Center;
        if (kind == TabStopValues.Decimal)
            return DxpComputedTabStopKind.Decimal;
        if (kind == TabStopValues.Left)
            return DxpComputedTabStopKind.Left;
        return null;
    }
}
