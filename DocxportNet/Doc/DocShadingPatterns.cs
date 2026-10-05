using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Doc;

internal static class DocShadingPatterns
{
    private static readonly ShadingPatternValues[] Patterns =
    [
        ShadingPatternValues.Clear, ShadingPatternValues.Solid,
        ShadingPatternValues.Percent5, ShadingPatternValues.Percent10,
        ShadingPatternValues.Percent20, ShadingPatternValues.Percent25,
        ShadingPatternValues.Percent30, ShadingPatternValues.Percent40,
        ShadingPatternValues.Percent50, ShadingPatternValues.Percent60,
        ShadingPatternValues.Percent70, ShadingPatternValues.Percent75,
        ShadingPatternValues.Percent80, ShadingPatternValues.Percent90,
        ShadingPatternValues.HorizontalStripe, ShadingPatternValues.VerticalStripe,
        ShadingPatternValues.ReverseDiagonalStripe, ShadingPatternValues.DiagonalStripe,
        ShadingPatternValues.HorizontalCross, ShadingPatternValues.DiagonalCross,
        ShadingPatternValues.ThinHorizontalStripe, ShadingPatternValues.ThinVerticalStripe,
        ShadingPatternValues.ThinReverseDiagonalStripe,
        ShadingPatternValues.ThinDiagonalStripe,
        ShadingPatternValues.ThinHorizontalCross,
        ShadingPatternValues.ThinDiagonalCross
    ];

    public static ushort? ToDoc(ShadingPatternValues value)
    {
        for (ushort i = 0; i < Patterns.Length; i++)
            if (Patterns[i] == value) return i;
        return value == ShadingPatternValues.Percent12 ? (ushort)0x25 :
            value == ShadingPatternValues.Percent15 ? (ushort)0x26 :
            value == ShadingPatternValues.Percent35 ? (ushort)0x2B :
            value == ShadingPatternValues.Percent37 ? (ushort)0x2C :
            value == ShadingPatternValues.Percent45 ? (ushort)0x2E :
            value == ShadingPatternValues.Percent55 ? (ushort)0x31 :
            value == ShadingPatternValues.Percent62 ? (ushort)0x33 :
            value == ShadingPatternValues.Percent65 ? (ushort)0x34 :
            value == ShadingPatternValues.Percent85 ? (ushort)0x39 :
            value == ShadingPatternValues.Percent87 ? (ushort)0x3A :
            value == ShadingPatternValues.Percent95 ? (ushort)0x3C : null;
    }

    public static ShadingPatternValues? ToOpenXml(ushort value) => value < Patterns.Length
        ? Patterns[value] : value switch
        {
            0x25 => ShadingPatternValues.Percent12,
            0x26 => ShadingPatternValues.Percent15,
            0x2B => ShadingPatternValues.Percent35,
            0x2C => ShadingPatternValues.Percent37,
            0x2E => ShadingPatternValues.Percent45,
            0x31 => ShadingPatternValues.Percent55,
            0x33 => ShadingPatternValues.Percent62,
            0x34 => ShadingPatternValues.Percent65,
            0x39 => ShadingPatternValues.Percent85,
            0x3A => ShadingPatternValues.Percent87,
            0x3C => ShadingPatternValues.Percent95,
            _ => null
        };
}
