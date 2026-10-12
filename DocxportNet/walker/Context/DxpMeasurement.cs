using System.Globalization;

namespace DocxportNet.Walker.Context;

/// <summary>A Word length, normalized to twips (twentieths of a point).</summary>
internal readonly record struct DxpMeasurement(int Twips)
{
    public double Points => Twips / 20.0;
    public double Inches => Twips / 1440.0;

    public static DxpMeasurement Parse(string? text) =>
        TryParse(text, out var value) ? value :
        throw new FormatException($"Invalid Word measurement: '{text}'.");

    public static bool TryParse(string? text, out DxpMeasurement value)
    {
        value = default;
        if (string.IsNullOrWhiteSpace(text))
            return false;

        text = text!.Trim();
        if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips))
        {
            value = new DxpMeasurement(twips);
            return true;
        }

        decimal factor = 0;
        int suffixLength = 2;
        if (text.EndsWith("in", StringComparison.OrdinalIgnoreCase)) factor = 1440m;
        else if (text.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) factor = 20m;
        else if (text.EndsWith("pc", StringComparison.OrdinalIgnoreCase) ||
                 text.EndsWith("pi", StringComparison.OrdinalIgnoreCase)) factor = 240m;
        else if (text.EndsWith("cm", StringComparison.OrdinalIgnoreCase)) factor = 1440m / 2.54m;
        else if (text.EndsWith("mm", StringComparison.OrdinalIgnoreCase)) factor = 1440m / 25.4m;
        else suffixLength = 0;

        if (suffixLength == 0 ||
            !decimal.TryParse(text.Substring(0, text.Length - suffixLength), NumberStyles.AllowLeadingSign |
                NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal amount))
            return false;

        decimal normalized;
        try
        {
            normalized = decimal.Round(amount * factor, 0, MidpointRounding.AwayFromZero);
        }
        catch (OverflowException)
        {
            return false;
        }
        if (normalized < int.MinValue || normalized > int.MaxValue)
            return false;

        value = new DxpMeasurement((int)normalized);
        return true;
    }
}
