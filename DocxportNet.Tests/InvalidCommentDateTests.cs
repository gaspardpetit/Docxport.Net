using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker.Context;
using DocxportNet.Walker.Parts;

namespace DocxportNet.Tests;

public sealed class InvalidCommentDateTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MalformedLegacyCommentDateKeepsCommentAndUsesValidUtcDate(bool validUtcDate)
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(new Paragraph(new Run(new CommentReference { Id = "0" }))));
        var comment = new Comment(new Paragraph(new Run(new Text("Annotation")))) { Id = "0" };
        comment.SetAttribute(new OpenXmlAttribute("w", "date",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0000-00-00T00:00:00Z"));
        if (validUtcDate)
            comment.DateUtc = new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc);
        main.AddNewPart<WordprocessingCommentsPart>().Comments = new Comments(comment);

        var comments = new DxpComments();
        comments.Init(main);
        var info = Assert.Single(comments.GetThreadForAnchor("0")!.Comments);
        Assert.Equal("Annotation", info.Text);
        Assert.Equal(validUtcDate ? new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc) : null,
            info.DateUtc);
        var timelineComment = Assert.Single(DxpTimeline.BuildTimeline(document), item => item.Kind == "comment");
        Assert.Equal(info.DateUtc, timelineComment.DateUtc);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MalformedCommentDateDoesNotPreventLoadingComment(bool validFallbackDate)
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(new Paragraph(new Run(new Text("Decision")),
            new Run(new CommentReference { Id = "0" }))));
        main.Document.Save();

        var comment = new Comment(new Paragraph(new Run(new Text("Annotation"))))
        {
            Id = "0", Author = "Editor"
        };
        comment.SetAttribute(new OpenXmlAttribute("w15", "dateUtc",
            "http://schemas.microsoft.com/office/word/2012/wordml", "0-00-00T00:00:00Z"));
        if (validFallbackDate)
            comment.Date = new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc);
        var part = main.AddNewPart<WordprocessingCommentsPart>();
        part.Comments = new Comments(comment);
        part.Comments.Save();

        var comments = new DxpComments();
        comments.Init(main);
        var info = Assert.Single(comments.GetThreadForAnchor("0")!.Comments);
        Assert.Equal("Annotation", info.Text);
        Assert.Equal(validFallbackDate ? new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc) : null,
            info.DateUtc);
        var timelineComment = Assert.Single(DxpTimeline.BuildTimeline(document), item => item.Kind == "comment");
        Assert.Equal(info.DateUtc, timelineComment.DateUtc);
    }
}
