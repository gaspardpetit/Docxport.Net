namespace DocxportNet.Tests.Utils;

public static class TestPaths
{
	public static string GetSampleOutputDirectory(string docxPath)
	{
		string? samplesDir = Path.GetDirectoryName(docxPath);
		if (string.IsNullOrWhiteSpace(samplesDir))
			throw new ArgumentException("DOCX path must have a directory.", nameof(docxPath));

		string sampleName = Path.GetFileNameWithoutExtension(docxPath);
		string outputDir = Path.Combine(samplesDir, sampleName);
		Directory.CreateDirectory(outputDir);
		return outputDir;
	}

	public static string GetSampleOutputPath(string docxPath, string suffix)
	{
		string outputDir = GetSampleOutputDirectory(docxPath);
		string sampleName = Path.GetFileNameWithoutExtension(docxPath);
		return Path.Combine(outputDir, $"{sampleName}{suffix}");
	}
}
