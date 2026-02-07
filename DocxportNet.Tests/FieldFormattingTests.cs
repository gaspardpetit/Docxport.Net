using System.Globalization;
using DocxportNet.Fields;
using DocxportNet.Fields.Formatting;
using Xunit;

namespace DocxportNet.Tests;

public class FieldFormattingTests
{
	[Fact]
	public void Parser_ExtractsFieldTypeAndSwitches()
	{
		var parser = new DxpFieldParser();
		var result = parser.Parse("DATE \\@ \"yyyy-MM-dd\" \\* Upper");

		Assert.True(result.Success);
		Assert.Equal("DATE", result.Ast.FieldType);
		Assert.Equal(2, result.Ast.FormatSpecs.Count);
		Assert.IsType<DxpDateTimeFormatSpec>(result.Ast.FormatSpecs[0]);
		Assert.IsType<DxpTextTransformFormatSpec>(result.Ast.FormatSpecs[1]);
	}

	[Fact]
	public void Formatter_AppliesDateTimeFormatThenTextFormat()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };
		var value = new DxpFieldValue(new DateTimeOffset(2026, 2, 6, 15, 4, 0, TimeSpan.Zero));
		var switches = new DocxportNet.Fields.Formatting.IDxpFieldFormatSpec[]
		{
			new DxpDateTimeFormatSpec("\\@ \"MMM d, yyyy\"", new[]
			{
				new DxpDateTimeToken(DxpDateTimeTokenKind.MonthShortName, "MMM"),
				new DxpDateTimeToken(DxpDateTimeTokenKind.Literal, " "),
				new DxpDateTimeToken(DxpDateTimeTokenKind.DayNumeric, "d"),
				new DxpDateTimeToken(DxpDateTimeTokenKind.Literal, ", "),
				new DxpDateTimeToken(DxpDateTimeTokenKind.Year4, "yyyy")
			}),
			new DxpTextTransformFormatSpec(DxpTextTransformKind.Upper, "\\* Upper", "Upper")
		};

		var text = formatter.Format(value, switches, context);

		Assert.Equal("FEB 6, 2026", text);
	}

	[Fact]
	public void Formatter_AlphabeticAndRomanTransforms()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext();

		var alpha = formatter.Format(
			new DxpFieldValue(27),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.Alphabetic, "\\* ALPHABETIC", "ALPHABETIC") },
			context);
		var roman = formatter.Format(
			new DxpFieldValue(9),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.Roman, "\\* roman", "roman") },
			context);

		Assert.Equal("AA", alpha);
		Assert.Equal("ix", roman);
	}

	[Fact]
	public void Formatter_OrdinalAndCardText()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var ordinal = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.Ordinal, "\\* Ordinal", "Ordinal") },
			context);
		var card = formatter.Format(
			new DxpFieldValue(42),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.CardText, "\\* CardText", "CardText") },
			context);

		Assert.Equal("21st", ordinal);
		Assert.Equal("forty-two", card);
	}

	[Fact]
	public void Formatter_NumericPictureFormats()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var spec1 = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"00.00\"").Ast.FormatSpecs[0];
		var spec2 = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# x##").Ast.FormatSpecs[0];
		var spec3 = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# $###").Ast.FormatSpecs[0];

		var r1 = formatter.Format(new DxpFieldValue(9), new IDxpFieldFormatSpec[] { spec1 }, context);
		var r2 = formatter.Format(new DxpFieldValue(222492), new IDxpFieldFormatSpec[] { spec2 }, context);
		var r3 = formatter.Format(new DxpFieldValue(15), new IDxpFieldFormatSpec[] { spec3 }, context);

		Assert.Equal("09.00", r1);
		Assert.Equal("492", r2);
		Assert.Equal("$ 15", r3);
	}

	[Fact]
	public void Formatter_NumericSectionsAndSigns()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var spec = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"$#,##0.00;-$#,##0.00;ZERO\"").Ast.FormatSpecs[0];
		var pos = formatter.Format(new DxpFieldValue(1234.5), new IDxpFieldFormatSpec[] { spec }, context);
		var neg = formatter.Format(new DxpFieldValue(-1234.5), new IDxpFieldFormatSpec[] { spec }, context);
		var zero = formatter.Format(new DxpFieldValue(0), new IDxpFieldFormatSpec[] { spec }, context);

		Assert.Equal("$1,234.50", pos);
		Assert.Equal("-$1,234.50", neg);
		Assert.Equal("ZERO", zero);
	}

	[Fact]
	public void Formatter_NumericGroupingPreservesExtraDigits()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var spec = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"#,##0\"").Ast.FormatSpecs[0];
		var result = formatter.Format(new DxpFieldValue(1234567), new IDxpFieldFormatSpec[] { spec }, context);

		Assert.Equal("1,234,567", result);
	}

	[Fact]
	public void Formatter_NumericGroupingPreservesExtraDigits_WithPrefix()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var spec = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"$#,##0\"").Ast.FormatSpecs[0];
		var result = formatter.Format(new DxpFieldValue(1234567), new IDxpFieldFormatSpec[] { spec }, context);

		Assert.Equal("$1,234,567", result);
	}

	[Fact]
	public void Formatter_NumericDropDigitsWithGrouping()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var spec = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"#,##x\"").Ast.FormatSpecs[0];
		var result = formatter.Format(new DxpFieldValue(1234567), new IDxpFieldFormatSpec[] { spec }, context);

		Assert.Equal("   7", result);
	}

	[Fact]
	public void Formatter_NumericDropDigits_PositionalX()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var spec = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"#x#\"").Ast.FormatSpecs[0];
		var result = formatter.Format(new DxpFieldValue(1234567), new IDxpFieldFormatSpec[] { spec }, context);

		Assert.Equal(" 67", result);
	}

	[Fact]
	public void Formatter_NumericFormat_NoPlaceholdersOmitsDigits()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var spec = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"'Tax:'\"").Ast.FormatSpecs[0];
		var result = formatter.Format(new DxpFieldValue(123), new IDxpFieldFormatSpec[] { spec }, context);

		Assert.Equal("Tax:", result);
	}

	[Fact]
	public void Formatter_NumericFormat_RightAlignsDigits()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var specZero = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"0000\"").Ast.FormatSpecs[0];
		var specOptional = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"####\"").Ast.FormatSpecs[0];

		var r1 = formatter.Format(new DxpFieldValue(12), new IDxpFieldFormatSpec[] { specZero }, context);
		var r2 = formatter.Format(new DxpFieldValue(12), new IDxpFieldFormatSpec[] { specOptional }, context);

		Assert.Equal("0012", r1);
		Assert.Equal("  12", r2);
	}

	[Fact]
	public void Formatter_NumericOptionalDigitsAndPlusSign()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var spec = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"+#; -#; 0\"").Ast.FormatSpecs[0];
		var pos = formatter.Format(new DxpFieldValue(7), new IDxpFieldFormatSpec[] { spec }, context);
		var neg = formatter.Format(new DxpFieldValue(-7), new IDxpFieldFormatSpec[] { spec }, context);
		var zero = formatter.Format(new DxpFieldValue(0), new IDxpFieldFormatSpec[] { spec }, context);

		Assert.Equal("+7", pos.Trim());
		Assert.Equal("-7", neg.Trim());
		Assert.Equal("0", zero.Trim());
	}

	[Fact]
	public void Formatter_NumericPercentAndLiterals()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var spec = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"'Tax:' 0.0%\"").Ast.FormatSpecs[0];
		var result = formatter.Format(new DxpFieldValue(12.3), new IDxpFieldFormatSpec[] { spec }, context);

		Assert.Equal("Tax: 12.3%", result);
	}

	[Fact]
	public void Formatter_NumericTrimsOptionalFraction()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var spec = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"0.##\"").Ast.FormatSpecs[0];
		var r1 = formatter.Format(new DxpFieldValue(1.2), new IDxpFieldFormatSpec[] { spec }, context);
		var r2 = formatter.Format(new DxpFieldValue(1.0), new IDxpFieldFormatSpec[] { spec }, context);

		Assert.Equal("1.2", r1);
		Assert.Equal("1", r2);
	}

	[Fact]
	public void Formatter_DateTimeTokensAndLiterals()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };
		var value = new DxpFieldValue(new DateTimeOffset(2026, 2, 7, 9, 5, 6, TimeSpan.Zero));

		var spec = (DxpDateTimeFormatSpec)new DxpFieldParser().Parse("DATE \\@ \"ddd, MMM d 'at' h:mm AM/PM\"").Ast.FormatSpecs[0];
		var text = formatter.Format(value, new IDxpFieldFormatSpec[] { spec }, context);

		Assert.Equal("Sat, Feb 7 at 9:05 AM", text);
	}

	[Fact]
	public void Formatter_RespectsCultureForCasing()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("tr-TR") };
		var value = new DxpFieldValue("istanbul");

		var upper = formatter.Format(
			value,
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.Upper, "\\* Upper", "Upper") },
			context);
		var firstCap = formatter.Format(
			value,
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.FirstCap, "\\* FirstCap", "FirstCap") },
			context);

		Assert.Equal("İSTANBUL", upper);
		Assert.Equal("İstanbul", firstCap);
	}

	[Fact]
	public void Formatter_AmPmPeriodsCanBePreservedOrStripped()
	{
		var formatter = new DxpFieldFormatter();
		var value = new DxpFieldValue(new DateTimeOffset(2026, 2, 7, 9, 0, 0, TimeSpan.Zero));
		var spec = (DxpDateTimeFormatSpec)new DxpFieldParser().Parse("DATE \\@ \"h AM/PM\"").Ast.FormatSpecs[0];

		var contextKeep = new DxpFieldEvalContext { Culture = new CultureInfo("en-US"), StripAmPmPeriods = false };
		var contextStrip = new DxpFieldEvalContext { Culture = new CultureInfo("en-US"), StripAmPmPeriods = true };

		var keep = formatter.Format(value, new IDxpFieldFormatSpec[] { spec }, contextKeep);
		var strip = formatter.Format(value, new IDxpFieldFormatSpec[] { spec }, contextStrip);

		Assert.Equal("9 AM", keep);
		Assert.Equal("9 AM", strip);
	}

	[Fact]
	public void Formatter_NumberedItemInNumericAndDateFormats()
	{
		var formatter = new DxpFieldFormatter();
		DxpFieldEvalContext context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };
		context.SetNumberedItem("table", "2");

		var numericSpec = (DxpNumericFormatSpec)new DxpFieldParser().Parse("X \\# \"0 `table`\"").Ast.FormatSpecs[0];
		var numericText = formatter.Format(new DxpFieldValue(5), new IDxpFieldFormatSpec[] { numericSpec }, context);

		var dateSpec = (DxpDateTimeFormatSpec)new DxpFieldParser().Parse("DATE \\@ \"yyyy `table`\"").Ast.FormatSpecs[0];
		var dateText = formatter.Format(
			new DxpFieldValue(new DateTimeOffset(2026, 2, 7, 0, 0, 0, TimeSpan.Zero)),
			new IDxpFieldFormatSpec[] { dateSpec },
			context);

		Assert.Equal("5 2", numericText);
		Assert.Equal("2026 2", dateText);
	}

	[Fact]
	public void Formatter_NumericParsingHonorsInvariantFallbackFlag()
	{
		var formatter = new DxpFieldFormatter();
		var spec = (DxpTextTransformFormatSpec)new DxpFieldParser().Parse("X \\* Hex").Ast.FormatSpecs[0];
		var value = new DxpFieldValue("1,234.5");

		var contextStrict = new DxpFieldEvalContext { Culture = new CultureInfo("fr-FR"), AllowInvariantNumericFallback = false };
		var contextLoose = new DxpFieldEvalContext { Culture = new CultureInfo("fr-FR"), AllowInvariantNumericFallback = true };

		var strict = formatter.Format(value, new IDxpFieldFormatSpec[] { spec }, contextStrict);
		var loose = formatter.Format(value, new IDxpFieldFormatSpec[] { spec }, contextLoose);

		Assert.Equal("1,234.5", strict);
		Assert.Equal("4D2", loose);
	}

	[Fact]
	public void Formatter_TextTransforms_DollarAndOrdTextAndArabicDash()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("en-US") };

		var dollar = formatter.Format(
			new DxpFieldValue(14.55),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.DollarText, "\\* DollarText", "DollarText") },
			context);
		var ordText = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.OrdText, "\\* OrdText", "OrdText") },
			context);
		var arabicDash = formatter.Format(
			new DxpFieldValue(31),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.ArabicDash, "\\* ArabicDash", "ArabicDash") },
			context);

		Assert.Equal("fourteen and 55/100", dollar);
		Assert.Equal("twenty-first", ordText);
		Assert.Equal("-31-", arabicDash);
	}

	[Fact]
	public void Formatter_TextTransforms_FrenchNumberWords()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("fr-FR") };

		var card = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.CardText, "\\* CardText", "CardText") },
			context);
		var ord = formatter.Format(
			new DxpFieldValue(2),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.OrdText, "\\* OrdText", "OrdText") },
			context);
		var dollar = formatter.Format(
			new DxpFieldValue(14.55),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.DollarText, "\\* DollarText", "DollarText") },
			context);

		Assert.Equal("vingt et un", card);
		Assert.Equal("deuxième", ord);
		Assert.Equal("quatorze et 55/100", dollar);
	}

	[Fact]
	public void Formatter_TextTransforms_JapaneseNumberWords()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("ja-JP") };

		var card = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.CardText, "\\* CardText", "CardText") },
			context);
		var ord = formatter.Format(
			new DxpFieldValue(3),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.OrdText, "\\* OrdText", "OrdText") },
			context);
		var dollar = formatter.Format(
			new DxpFieldValue(14.55),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.DollarText, "\\* DollarText", "DollarText") },
			context);

		Assert.Equal("二十一", card);
		Assert.Equal("第三", ord);
		Assert.Equal("十四 と 55/100", dollar);
	}

	[Fact]
	public void Formatter_TextTransforms_ThaiNumberWords()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("th-TH") };

		var card = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.CardText, "\\* CardText", "CardText") },
			context);
		var ord = formatter.Format(
			new DxpFieldValue(2),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.OrdText, "\\* OrdText", "OrdText") },
			context);
		var dollar = formatter.Format(
			new DxpFieldValue(14.55),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.DollarText, "\\* DollarText", "DollarText") },
			context);

		Assert.Equal("ยี่สิบเอ็ด", card);
		Assert.Equal("ที่สอง", ord);
		Assert.Equal("สิบสี่ และ 55/100", dollar);
	}

	[Fact]
	public void Formatter_TextTransforms_GermanNumberWords()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("de-DE") };

		var card = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.CardText, "\\* CardText", "CardText") },
			context);
		var ord = formatter.Format(
			new DxpFieldValue(7),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.OrdText, "\\* OrdText", "OrdText") },
			context);
		var dollar = formatter.Format(
			new DxpFieldValue(14.55),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.DollarText, "\\* DollarText", "DollarText") },
			context);

		Assert.Equal("einundzwanzig", card);
		Assert.Equal("siebte", ord);
		Assert.Equal("vierzehn und 55/100", dollar);
	}

	[Fact]
	public void Formatter_TextTransforms_ChineseSimplifiedNumberWords()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("zh-CN") };

		var card = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.CardText, "\\* CardText", "CardText") },
			context);
		var ord = formatter.Format(
			new DxpFieldValue(3),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.OrdText, "\\* OrdText", "OrdText") },
			context);
		var dollar = formatter.Format(
			new DxpFieldValue(14.55),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.DollarText, "\\* DollarText", "DollarText") },
			context);

		Assert.Equal("二十一", card);
		Assert.Equal("第三", ord);
		Assert.Equal("十四和55/100", dollar);
	}

	[Fact]
	public void Formatter_TextTransforms_SpanishNumberWords()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("es-ES") };

		var card = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.CardText, "\\* CardText", "CardText") },
			context);
		var ord = formatter.Format(
			new DxpFieldValue(7),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.OrdText, "\\* OrdText", "OrdText") },
			context);
		var dollar = formatter.Format(
			new DxpFieldValue(14.55),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.DollarText, "\\* DollarText", "DollarText") },
			context);

		Assert.Equal("veintiuno", card);
		Assert.Equal("séptimo", ord);
		Assert.Equal("catorce y 55/100", dollar);
	}

	[Fact]
	public void Formatter_TextTransforms_ItalianNumberWords()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("it-IT") };

		var card = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.CardText, "\\* CardText", "CardText") },
			context);
		var ord = formatter.Format(
			new DxpFieldValue(8),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.OrdText, "\\* OrdText", "OrdText") },
			context);
		var dollar = formatter.Format(
			new DxpFieldValue(14.55),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.DollarText, "\\* DollarText", "DollarText") },
			context);

		Assert.Equal("ventuno", card);
		Assert.Equal("ottavo", ord);
		Assert.Equal("quattordici e 55/100", dollar);
	}

	[Fact]
	public void Formatter_TextTransforms_PortugueseNumberWords()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("pt-PT") };

		var card = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.CardText, "\\* CardText", "CardText") },
			context);
		var ord = formatter.Format(
			new DxpFieldValue(7),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.OrdText, "\\* OrdText", "OrdText") },
			context);
		var dollar = formatter.Format(
			new DxpFieldValue(14.55),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.DollarText, "\\* DollarText", "DollarText") },
			context);

		Assert.Equal("vinte e um", card);
		Assert.Equal("sétimo", ord);
		Assert.Equal("catorze e 55/100", dollar);
	}

	[Fact]
	public void Formatter_TextTransforms_DanishNumberWords()
	{
		var formatter = new DxpFieldFormatter();
		var context = new DxpFieldEvalContext { Culture = new CultureInfo("da-DK") };

		var card = formatter.Format(
			new DxpFieldValue(21),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.CardText, "\\* CardText", "CardText") },
			context);
		var ord = formatter.Format(
			new DxpFieldValue(2),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.OrdText, "\\* OrdText", "OrdText") },
			context);
		var dollar = formatter.Format(
			new DxpFieldValue(14.55),
			new IDxpFieldFormatSpec[] { new DxpTextTransformFormatSpec(DxpTextTransformKind.DollarText, "\\* DollarText", "DollarText") },
			context);

		Assert.Equal("enogtyve", card);
		Assert.Equal("anden", ord);
		Assert.Equal("fjorten og 55/100", dollar);
	}
}
