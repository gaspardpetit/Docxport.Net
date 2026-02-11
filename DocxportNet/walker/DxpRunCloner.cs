using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Walker;

internal static class DxpRunCloner
{
    public static Run CloneRunWithParagraphAncestor(Run run)
    {
        var clonedRun = new Run();
        if (run.RunProperties != null)
            clonedRun.RunProperties = (RunProperties)run.RunProperties.CloneNode(true);

        var paragraph = run.Ancestors<Paragraph>().FirstOrDefault();
        if (paragraph == null)
            return clonedRun;

        var paraClone = (Paragraph)paragraph.CloneNode(false);
        if (paragraph.ParagraphProperties != null)
            paraClone.ParagraphProperties = (ParagraphProperties)paragraph.ParagraphProperties.CloneNode(true);
        paraClone.AppendChild(clonedRun);
        return clonedRun;
    }
}
