using System.Globalization;

namespace DocxportNet.Walker.Context;

internal static class DxpThemeColor
{
    public static string? ApplyLuminance(string color, string modifier, bool tint)
    {
        if (modifier.Length != 2 || !byte.TryParse(modifier,
            NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var amount) ||
            !uint.TryParse(color, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
            return null;
        var red = ((rgb >> 16) & 0xFF) / 255d;
        var green = ((rgb >> 8) & 0xFF) / 255d;
        var blue = (rgb & 0xFF) / 255d;
        var maximum = Math.Max(red, Math.Max(green, blue));
        var minimum = Math.Min(red, Math.Min(green, blue));
        var delta = maximum - minimum;
        var luminance = (maximum + minimum) / 2;
        var saturation = delta == 0 ? 0 :
            delta / (1 - Math.Abs(2 * luminance - 1));
        var hue = delta == 0 ? 0 : maximum == red
            ? ((green - blue) / delta + 6) % 6
            : maximum == green ? (blue - red) / delta + 2
            : (red - green) / delta + 4;
        var factor = amount / 255d;
        luminance = tint ? luminance * factor + 1 - factor : luminance * factor;
        var chroma = (1 - Math.Abs(2 * luminance - 1)) * saturation;
        var x = chroma * (1 - Math.Abs(hue % 2 - 1));
        var middle = luminance - chroma / 2;
        var components = hue switch
        {
            < 1 => (chroma, x, 0d),
            < 2 => (x, chroma, 0d),
            < 3 => (0d, chroma, x),
            < 4 => (0d, x, chroma),
            < 5 => (x, 0d, chroma),
            _ => (chroma, 0d, x)
        };
        static int Channel(double value) => Math.Max(0, Math.Min(255,
            (int)Math.Round(value * 255, MidpointRounding.AwayFromZero)));
        return $"{Channel(components.Item1 + middle):X2}" +
            $"{Channel(components.Item2 + middle):X2}" +
            $"{Channel(components.Item3 + middle):X2}";
    }
}
