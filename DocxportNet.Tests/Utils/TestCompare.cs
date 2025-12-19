namespace DocxportNet.Tests.Utils;

internal class TestCompare
{
	public static string Normalize(string text)
	{
		return text.Replace("\r\n", "\n").Replace("\r", "\n");
	}

	public static string DescribeDifference(string expected, string actual)
	{
		int mismatchIndex = FindFirstDifferenceIndex(expected, actual);
		if (mismatchIndex < Math.Min(expected.Length, actual.Length))
		{
			var (line, column) = GetLineAndColumn(expected, mismatchIndex);
			return $"first difference at line {line}, column {column}: expected {DescribeChar(expected[mismatchIndex])}, found {DescribeChar(actual[mismatchIndex])}";
		}

		if (expected.Length != actual.Length)
		{
			var (line, column) = GetLineAndColumn(expected, Math.Min(expected.Length, actual.Length));
			return expected.Length < actual.Length
				? $"actual is longer starting at line {line}, column {column} with {DescribeChar(actual[mismatchIndex])}"
				: $"actual is shorter; expected {DescribeChar(expected[mismatchIndex])} at line {line}, column {column}";
		}

		return "unknown difference";
	}

	public static int FindFirstDifferenceIndex(string expected, string actual)
	{
		int minLength = Math.Min(expected.Length, actual.Length);
		for (int i = 0; i < minLength; i++)
		{
			if (expected[i] != actual[i])
				return i;
		}

		return minLength;
	}

	public static (int line, int column) GetLineAndColumn(string text, int index)
	{
		int line = 1;
		int column = 1;

		for (int i = 0; i < index; i++)
		{
			if (text[i] == '\n')
			{
				line++;
				column = 1;
			}
			else
			{
				column++;
			}
		}

		return (line, column);
	}

	public static string DescribeChar(char c)
	{
		return c switch {
			'\n' => "'\\n'",
			'\r' => "'\\r'",
			'\t' => "'\\t'",
			_ => $"'{c}' (0x{(int)c:x2})"
		};
	}
}
