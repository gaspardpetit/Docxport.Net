
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;

public sealed record SectionSlice(SectionProperties Properties, IReadOnlyList<OpenXmlElement> Blocks);

public class DxpSectionContext : DxpISectionContext
{
    public DxpSectionContext(SectionProperties? sectionProperties, SectionLayout? layoutRaw, DxpSectionLayout? layout)
    {
        SectionProperties = sectionProperties;
        LayoutRaw = layoutRaw;
        Layout = layout;
    }

    public SectionProperties? SectionProperties { get; internal set; }
    public SectionLayout? LayoutRaw { get; internal set; }
    public DxpSectionLayout? Layout { get; internal set; }
    public bool IsLast { get; internal set; }

    public static DxpSectionContext INVALID => new DxpSectionContext(null!, null!, null!);
}

public class DxpSections
{
    public static SectionProperties? ExtractSectionProperties(OpenXmlElement block, out bool includeBlock)
    {
        includeBlock = true;

        if (block is SectionProperties sp)
        {
            includeBlock = false;
            return sp;
        }

        if (block is Paragraph p)
        {
            var pp = p.GetFirstChild<ParagraphProperties>();
            var paragraphSectPr = pp?.GetFirstChild<SectionProperties>();
            return paragraphSectPr;
        }

        return null;
    }

    public static List<SectionSlice> SplitDocumentBodyIntoSections(Body body)
    {
        var sections = new List<SectionSlice>();
        var sectPrs = body.Descendants<SectionProperties>().ToList();
        if (sectPrs.Count == 0)
        {
            // Some minimal/hand-crafted DOCX packages omit a final w:sectPr.
            // Word treats such documents as having an implicit default section.
            sections.Add(new SectionSlice(new SectionProperties(), body.ChildElements.ToList()));
            return sections;
        }

        int sectIndex = 0;
        var currentSectPr = sectPrs[sectIndex];
        var currentBlocks = new List<OpenXmlElement>();

        foreach (var child in body.ChildElements)
        {
            bool include = true;
            var sp = ExtractSectionProperties(child, out include);

            if (include)
                currentBlocks.Add(child);

            if (sp != null)
            {
                var props = currentSectPr ?? sp;
                sections.Add(new SectionSlice(props, currentBlocks.ToList()));
                currentBlocks.Clear();

                sectIndex++;
                currentSectPr = sectIndex < sectPrs.Count ? sectPrs[sectIndex] : null;
            }
        }

        if (currentBlocks.Count > 0 && currentSectPr != null)
            sections.Add(new SectionSlice(currentSectPr, currentBlocks));
        else if (currentBlocks.Count > 0)
            sections.Add(new SectionSlice(sectPrs.LastOrDefault() ?? new SectionProperties(), currentBlocks));

        return sections;
    }

    public static HeaderReference? FindFirstSectionHeaderReference(SectionProperties sp)
    {
        return PickReference(sp.Elements<HeaderReference>(), sp, true, r => r.Type?.Value);
    }

    public static FooterReference? FindLastSectionFooterReference(SectionProperties sp)
    {
        return PickReference(sp.Elements<FooterReference>(), sp, false, r => r.Type?.Value);
    }

    public static T? PickReference<T>(IEnumerable<T> refs, SectionProperties sp, bool preferFirst, Func<T, HeaderFooterValues?> typeSelector)
        where T : class
    {
        var list = refs?.ToList();
        if (list == null || list.Count == 0)
            return null;

        bool useFirst = preferFirst && sp.GetFirstChild<TitlePage>() != null;

        var ordered = useFirst
            ? new[] { HeaderFooterValues.First, HeaderFooterValues.Default, HeaderFooterValues.Even }
            : new[] { HeaderFooterValues.Default, HeaderFooterValues.First, HeaderFooterValues.Even };

        foreach (var target in ordered)
        {
            var match = list.FirstOrDefault(r => NormalizeHeaderFooterType(typeSelector(r)) == target);
            if (match != null)
                return match;
        }

        return list.FirstOrDefault();
    }

    private static HeaderFooterValues NormalizeHeaderFooterType(HeaderFooterValues? type)
    {
        return type ?? HeaderFooterValues.Default;
    }

    public static SectionLayout CreateSectionLayout(SectionProperties sp)
    {
        return new SectionLayout {
            PageSize = sp.GetFirstChild<PageSize>(),
            PageMargin = sp.GetFirstChild<PageMargin>(),
            Columns = sp.GetFirstChild<Columns>(),
            DocGrid = sp.GetFirstChild<DocGrid>(),
            PageBorders = sp.GetFirstChild<PageBorders>(),
            LineNumbers = sp.GetFirstChild<LineNumberType>(),
            TextDirection = sp.GetFirstChild<TextDirection>(),
            VerticalJustification = sp.GetFirstChild<VerticalTextAlignment>(),
            FootnoteProperties = sp.GetFirstChild<FootnoteProperties>(),
            EndnoteProperties = sp.GetFirstChild<EndnoteProperties>(),
        };
    }
}
