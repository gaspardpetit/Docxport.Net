using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker.Context;
using Xunit.Abstractions;

namespace DocxportNet.Tests;

public class StyleResolverDuplicateStylesTests : TestBase<StyleResolverDuplicateStylesTests>
{
    public StyleResolverDuplicateStylesTests(ITestOutputHelper output) : base(output)
    {
    }

    [Fact]
    public void RunStyle_InheritsAndOverridesVisualProperties()
    {
        using var doc = CreateDocument((body, styles) => {
            styles.Append(new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                new Vanish(), new Color { Val = "112233" }))));
            styles.Append(new Style {
                StyleId = "ParagraphStyle", Type = StyleValues.Paragraph,
                StyleRunProperties = new StyleRunProperties(new Color { Val = "445566" })
            });
            styles.Append(new Style {
                StyleId = "CharacterStyle", Type = StyleValues.Character,
                StyleRunProperties = new StyleRunProperties(new Vanish { Val = false })
            });
            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "ParagraphStyle" }),
                new Run(new Text("inherited")),
                new Run(new RunProperties(new RunStyle { Val = "CharacterStyle" },
                    new Color { Val = "AABBCC" }), new Text("override"))));
        });

        var resolver = new DxpStyleResolver(doc, Logger);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var runs = paragraph.Elements<Run>().ToArray();

        var inherited = resolver.ResolveRunStyle(paragraph, runs[0]);
        Assert.True(inherited.Hidden);
        Assert.Equal("#445566", inherited.ColorCss);

        var overridden = resolver.ResolveRunStyle(paragraph, runs[1]);
        Assert.False(overridden.Hidden);
        Assert.Equal("#AABBCC", overridden.ColorCss);
    }

    [Fact]
    public void RunStyle_ResolvesWord2010LigaturesAcrossStyles()
    {
        const string word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";
        using var doc = CreateDocument((body, styles) => {
            var styleProperties = new StyleRunProperties();
            var styleLigatures = new OpenXmlUnknownElement("w14", "ligatures", word2010Namespace);
            styleLigatures.SetAttribute(new OpenXmlAttribute("w14", "val", word2010Namespace, "standardContextual"));
            styleProperties.Append(styleLigatures);
            styles.Append(new Style {
                StyleId = "CharacterStyle", Type = StyleValues.Character,
                StyleRunProperties = styleProperties
            });
            var directProperties = new RunProperties(new RunStyle { Val = "CharacterStyle" });
            var directLigatures = new OpenXmlUnknownElement("w14", "ligatures", word2010Namespace);
            directLigatures.SetAttribute(new OpenXmlAttribute("w14", "val", word2010Namespace, "none"));
            directProperties.Append(directLigatures);
            body.Append(new Paragraph(
                new Run(new RunProperties(new RunStyle { Val = "CharacterStyle" }), new Text("inherited")),
                new Run(directProperties, new Text("override"))));
        });

        var resolver = new DxpStyleResolver(doc, Logger);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var runs = paragraph.Elements<Run>().ToArray();

        Assert.Equal("standardContextual", resolver.ResolveRunStyle(paragraph, runs[0]).Ligatures);
        Assert.Equal("none", resolver.ResolveRunStyle(paragraph, runs[1]).Ligatures);
    }

    [Fact]
    public void CrossTypeDuplicateStyleId_DoesNotCrash_AndResolvesByContext()
    {
        using var doc = CreateDocument((body, styles) => {
            styles.Append(
                new Style {
                    StyleId = "DefaultParagraphFont",
                    Type = StyleValues.Paragraph,
                    StyleName = new StyleName { Val = "Normal" }
                },
                new Style {
                    StyleId = "DefaultParagraphFont",
                    Type = StyleValues.Character,
                    StyleName = new StyleName { Val = "Default Paragraph Font" },
                    StyleRunProperties = new StyleRunProperties(new Bold())
                });

            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "DefaultParagraphFont" }),
                new Run(
                    new RunProperties(new RunStyle { Val = "DefaultParagraphFont" }),
                    new Text("x"))));
        });

        var resolver = new DxpStyleResolver(doc, Logger);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var run = paragraph.Elements<Run>().First();

        var chain = resolver.GetParagraphStyleChain(paragraph);
        var runStyle = resolver.ResolveRunStyle(paragraph, run);

        Assert.Single(chain);
        Assert.Equal(StyleValues.Paragraph, chain[0].Type);
        Assert.Equal("Normal", chain[0].Name);
        Assert.True(runStyle.Bold);
    }

    [Fact]
    public void SameTypeDuplicateStyle_LastWins()
    {
        using var doc = CreateDocument((body, styles) => {
            styles.Append(
                new Style {
                    StyleId = "MyP",
                    Type = StyleValues.Paragraph,
                    StyleName = new StyleName { Val = "MyP old" },
                    StyleParagraphProperties = new StyleParagraphProperties(new Justification { Val = JustificationValues.Left })
                },
                new Style {
                    StyleId = "MyP",
                    Type = StyleValues.Paragraph,
                    StyleName = new StyleName { Val = "MyP new" },
                    StyleParagraphProperties = new StyleParagraphProperties(new Justification { Val = JustificationValues.Center })
                });

            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "MyP" }),
                new Run(new Text("x"))));
        });

        var resolver = new DxpStyleResolver(doc, Logger);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();

        var jc = resolver.GetJustification(paragraph);

        Assert.Equal(JustificationValues.Center, jc);
    }

    [Fact]
    public void MissingPreferredType_WithSingleAlternateCandidate_UsesFallback()
    {
        using var doc = CreateDocument((body, styles) => {
            styles.Append(new Style {
                StyleId = "OnlyChar",
                Type = StyleValues.Character,
                StyleName = new StyleName { Val = "Only char" }
            });

            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "OnlyChar" }),
                new Run(new Text("x"))));
        });

        var resolver = new DxpStyleResolver(doc, Logger);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();

        var chain = resolver.GetParagraphStyleChain(paragraph);

        Assert.Single(chain);
        Assert.Equal("OnlyChar", chain[0].StyleId);
        Assert.Equal(StyleValues.Character, chain[0].Type);
    }

    [Fact]
    public void MissingPreferredType_WithMultipleCandidates_UsesDeterministicFallback()
    {
        using var doc = CreateDocument((body, styles) => {
            styles.Append(
                new Style {
                    StyleId = "Ambi",
                    Type = StyleValues.Character,
                    StyleName = new StyleName { Val = "Ambi char" }
                },
                new Style {
                    StyleId = "Ambi",
                    Type = StyleValues.Table,
                    StyleName = new StyleName { Val = "Ambi table" }
                });

            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Ambi" }),
                new Run(new Text("x"))));
        });

        var resolver = new DxpStyleResolver(doc, Logger);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();

        var chain = resolver.GetParagraphStyleChain(paragraph);

        Assert.Single(chain);
        Assert.Equal(StyleValues.Character, chain[0].Type);
    }

    [Fact]
    public void SyntheticDocumentWithDuplicateDefaultParagraphFont_DoesNotThrow()
    {
        using var doc = CreateDocument((body, styles) => {
            styles.Append(
                new Style {
                    StyleId = "DefaultParagraphFont",
                    Type = StyleValues.Paragraph,
                    StyleName = new StyleName { Val = "Normal" }
                },
                new Style {
                    StyleId = "DefaultParagraphFont",
                    Type = StyleValues.Character,
                    StyleName = new StyleName { Val = "Default Paragraph Font" }
                });

            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "DefaultParagraphFont" }),
                new Run(
                    new RunProperties(new RunStyle { Val = "DefaultParagraphFont" }),
                    new Text("x"))));
        });

        var ex = Record.Exception(() => _ = new DxpStyleResolver(doc, Logger));
        Assert.Null(ex);
    }

    private static WordprocessingDocument CreateDocument(Action<Body, Styles> configure)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();
            main.Document = new Document(body);

            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles();
            stylesPart.Styles = styles;

            configure(body, styles);

            styles.Save();
            main.Document.Save();
        }

        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }
}
