using System.Globalization;
using System.Text;

namespace DocxportNet.Fields.Formatting;

public enum DxpDateTimeTokenKind
{
    MonthNumeric,
    MonthNumeric2,
    MonthShortName,
    MonthFullName,
    DayNumeric,
    DayNumeric2,
    WeekdayShortName,
    WeekdayFullName,
    Year2,
    Year4,
    Hour12,
    Hour12_2,
    Hour24,
    Hour24_2,
    Minute,
    Minute2,
    Second,
    Second2,
    AmPm,
    Literal,
    NumberedItem
}

public sealed record DxpDateTimeToken(DxpDateTimeTokenKind Kind, string Text);

public sealed class DxpDateTimeFormatSpec : IDxpFieldFormatSpec
{
    public string RawText { get; }
    public IReadOnlyList<DxpDateTimeToken> Tokens { get; }

    public DxpDateTimeFormatSpec(string rawText, IReadOnlyList<DxpDateTimeToken> tokens)
    {
        RawText = rawText;
        Tokens = tokens;
    }

    public string Apply(string text, DxpFieldValue value, DxpFieldEvalContext context)
    {
        if (value.Kind != DxpFieldValueKind.DateTime || value.DateTimeValue == null)
        {
            if (!value.TryConvertToKind(DxpFieldValueKind.DateTime, context, out value) || value.DateTimeValue == null)
                return text;
        }

        var culture = context.Culture ?? CultureInfo.CurrentCulture;
        var sb = new StringBuilder();
        DateTimeOffset dt = value.DateTimeValue.Value;
        foreach (var token in Tokens)
        {
            switch (token.Kind)
            {
                case DxpDateTimeTokenKind.MonthNumeric:
                    sb.Append(dt.Month.ToString(culture));
                    break;
                case DxpDateTimeTokenKind.MonthNumeric2:
                    sb.Append(dt.Month.ToString("00", culture));
                    break;
                case DxpDateTimeTokenKind.MonthShortName:
                    sb.Append(culture.DateTimeFormat.GetAbbreviatedMonthName(dt.Month));
                    break;
                case DxpDateTimeTokenKind.MonthFullName:
                    sb.Append(culture.DateTimeFormat.GetMonthName(dt.Month));
                    break;
                case DxpDateTimeTokenKind.DayNumeric:
                    sb.Append(dt.Day.ToString(culture));
                    break;
                case DxpDateTimeTokenKind.DayNumeric2:
                    sb.Append(dt.Day.ToString("00", culture));
                    break;
                case DxpDateTimeTokenKind.WeekdayShortName:
                    sb.Append(culture.DateTimeFormat.GetAbbreviatedDayName(dt.DayOfWeek));
                    break;
                case DxpDateTimeTokenKind.WeekdayFullName:
                    sb.Append(culture.DateTimeFormat.GetDayName(dt.DayOfWeek));
                    break;
                case DxpDateTimeTokenKind.Year2:
                    sb.Append((dt.Year % 100).ToString("00", culture));
                    break;
                case DxpDateTimeTokenKind.Year4:
                    sb.Append(dt.Year.ToString("0000", culture));
                    break;
                case DxpDateTimeTokenKind.Hour12:
                    sb.Append(FormatHour12(dt, "0", culture));
                    break;
                case DxpDateTimeTokenKind.Hour12_2:
                    sb.Append(FormatHour12(dt, "00", culture));
                    break;
                case DxpDateTimeTokenKind.Hour24:
                    sb.Append(dt.Hour.ToString(culture));
                    break;
                case DxpDateTimeTokenKind.Hour24_2:
                    sb.Append(dt.Hour.ToString("00", culture));
                    break;
                case DxpDateTimeTokenKind.Minute:
                    sb.Append(dt.Minute.ToString(culture));
                    break;
                case DxpDateTimeTokenKind.Minute2:
                    sb.Append(dt.Minute.ToString("00", culture));
                    break;
                case DxpDateTimeTokenKind.Second:
                    sb.Append(dt.Second.ToString(culture));
                    break;
                case DxpDateTimeTokenKind.Second2:
                    sb.Append(dt.Second.ToString("00", culture));
                    break;
                case DxpDateTimeTokenKind.AmPm:
                {
                    string ampm = dt.ToString("tt", culture);
                    if (context.StripAmPmPeriods)
                        ampm = ampm.Replace(".", string.Empty);
                    sb.Append(ampm);
                    break;
                }
                case DxpDateTimeTokenKind.Literal:
                    sb.Append(token.Text);
                    break;
                case DxpDateTimeTokenKind.NumberedItem:
                    if (context.TryGetNumberedItem(token.Text, out var numbered))
                        sb.Append(numbered);
                    break;
            }
        }

        return sb.ToString();
    }

    private static string FormatHour12(DateTimeOffset dt, string format, CultureInfo culture)
    {
        int hour = dt.Hour % 12;
        if (hour == 0)
            hour = 12;
        return hour.ToString(format, culture);
    }
}
