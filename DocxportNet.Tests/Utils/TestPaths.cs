namespace DocxportNet.Tests.Utils;

public static class TestPaths
{
    private static string GetSampleResultPath(string docxPath, string suite, string suffix, bool generated)
    {
        string projectRoot = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(docxPath) ??
            throw new ArgumentException("DOCX path must have a directory.", nameof(docxPath)), "..", "..", ".."));
        string sampleName = Path.GetFileNameWithoutExtension(docxPath);
        string resultDir = Path.Combine(projectRoot, generated ? "output" : "truth", suite, sampleName);
        if (generated)
            Directory.CreateDirectory(resultDir);
        return Path.Combine(resultDir, $"{sampleName}{suffix}");
    }

    public static string GetSampleTruthPath(string docxPath, string suite, string suffix) =>
        GetSampleResultPath(docxPath, suite, suffix, generated: false);

    public static string GetSampleOutputPath(string docxPath, string suite, string suffix) =>
        GetSampleResultPath(docxPath, suite, suffix, generated: true);
}
