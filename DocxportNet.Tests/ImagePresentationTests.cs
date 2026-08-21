using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxportNet.Tests;

public class ImagePresentationTests
{
    [Fact]
    public void ParserReadsFrameCropRotationFlipAndMetadata()
    {
        Drawing drawing = CreateDrawing(
            new A.SourceRectangle { Left = 10000, Top = 20000, Right = 30000, Bottom = 10000 },
            new A.Transform2D(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 1270000L, Cy = 635000L }) {
                Rotation = 5400000,
                HorizontalFlip = true,
                VerticalFlip = true
            });

        var presentation = DxpDrawings.BuildImagePresentation(drawing);

        Assert.Equal(100, presentation.FrameWidthPoints);
        Assert.Equal(50, presentation.FrameHeightPoints);
        Assert.Equal(0.1, presentation.Crop!.Left, 6);
        Assert.Equal(0.2, presentation.Crop.Top, 6);
        Assert.Equal(0.3, presentation.Crop.Right, 6);
        Assert.Equal(0.1, presentation.Crop.Bottom, 6);
        Assert.Equal(90, presentation.RotationDegrees);
        Assert.True(presentation.FlipHorizontal);
        Assert.True(presentation.FlipVertical);
        Assert.Equal("Accessible description", presentation.AlternativeText);
        Assert.Equal("Image title", presentation.Title);
    }

    [Fact]
    public void ParserTreatsEmptyAndDegenerateCropAsNoCrop()
    {
        var empty = DxpDrawings.BuildImagePresentation(CreateDrawing(new A.SourceRectangle(), new A.Transform2D()));
        var degenerate = DxpDrawings.BuildImagePresentation(CreateDrawing(
            new A.SourceRectangle { Left = 60000, Right = 60000 },
            new A.Transform2D()));

        Assert.Null(empty.Crop);
        Assert.Null(degenerate.Crop);
    }

    [Fact]
    public void ParserClampsCropAndNormalizesNegativeRotation()
    {
        var presentation = DxpDrawings.BuildImagePresentation(CreateDrawing(
            new A.SourceRectangle { Left = -100, Top = 10000, Right = 200000, Bottom = -50 },
            new A.Transform2D { Rotation = -5400000 }));

        Assert.Null(presentation.Crop);
        Assert.Equal(270, presentation.RotationDegrees);
    }

    [Fact]
    public void ParserFallsBackToPictureMetadataAndRecognizesDecorativeImages()
    {
        Drawing drawing = CreateDrawing(new A.SourceRectangle(), new A.Transform2D());
        var docProperties = drawing.Descendants<WP.DocProperties>().Single();
        docProperties.Description = null;
        docProperties.Title = null;
        var decorative = new OpenXmlUnknownElement(
            "a16",
            "decorative",
            "http://schemas.microsoft.com/office/drawing/2014/main");
        decorative.SetAttribute(new OpenXmlAttribute("", "val", "", "1"));
        drawing.Append(decorative);

        var presentation = DxpDrawings.BuildImagePresentation(drawing);

        Assert.True(presentation.IsDecorative);
        Assert.Null(presentation.AlternativeText);
        Assert.Equal("Picture title", presentation.Title);
    }

    private static Drawing CreateDrawing(A.SourceRectangle sourceRectangle, A.Transform2D transform)
    {
        var picture = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties {
                    Id = 1U,
                    Name = "Picture",
                    Description = "Picture description",
                    Title = "Picture title"
                },
                new PIC.NonVisualPictureDrawingProperties()),
            new PIC.BlipFill(new A.Blip { Embed = "rIdImage1" }, sourceRectangle, new A.Stretch(new A.FillRectangle())),
            new PIC.ShapeProperties(transform, new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        return new Drawing(new WP.Inline(
            new WP.Extent { Cx = 1270000L, Cy = 635000L },
            new WP.DocProperties {
                Id = 1U,
                Name = "Picture",
                Description = "Accessible description",
                Title = "Image title"
            },
            new A.Graphic(new A.GraphicData(picture) {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
            })));
    }
}
