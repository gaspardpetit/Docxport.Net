using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Doc;

internal static class DocTableStyleCondition
{
    internal static ushort FromOpenXml(TableStyleOverrideValues value)
    {
        if (value == TableStyleOverrideValues.FirstRow) return 0x0001;
        if (value == TableStyleOverrideValues.LastRow) return 0x0002;
        if (value == TableStyleOverrideValues.FirstColumn) return 0x0004;
        if (value == TableStyleOverrideValues.LastColumn) return 0x0008;
        if (value == TableStyleOverrideValues.Band1Vertical) return 0x0010;
        if (value == TableStyleOverrideValues.Band2Vertical) return 0x0020;
        if (value == TableStyleOverrideValues.Band1Horizontal) return 0x0040;
        if (value == TableStyleOverrideValues.Band2Horizontal) return 0x0080;
        if (value == TableStyleOverrideValues.NorthEastCell) return 0x0100;
        if (value == TableStyleOverrideValues.NorthWestCell) return 0x0200;
        if (value == TableStyleOverrideValues.SouthEastCell) return 0x0400;
        if (value == TableStyleOverrideValues.SouthWestCell) return 0x0800;
        return 0;
    }

    internal static TableStyleOverrideValues ToOpenXml(ushort value) => value switch
    {
        0x0001 => TableStyleOverrideValues.FirstRow,
        0x0002 => TableStyleOverrideValues.LastRow,
        0x0004 => TableStyleOverrideValues.FirstColumn,
        0x0008 => TableStyleOverrideValues.LastColumn,
        0x0010 => TableStyleOverrideValues.Band1Vertical,
        0x0020 => TableStyleOverrideValues.Band2Vertical,
        0x0040 => TableStyleOverrideValues.Band1Horizontal,
        0x0080 => TableStyleOverrideValues.Band2Horizontal,
        0x0100 => TableStyleOverrideValues.NorthEastCell,
        0x0200 => TableStyleOverrideValues.NorthWestCell,
        0x0400 => TableStyleOverrideValues.SouthEastCell,
        0x0800 => TableStyleOverrideValues.SouthWestCell,
        _ => throw new InvalidDataException("A DOC table style has an invalid condition.")
    };

    internal static bool IsValid(ushort value) => value is
        0x0001 or 0x0002 or 0x0004 or 0x0008 or 0x0010 or 0x0020 or
        0x0040 or 0x0080 or 0x0100 or 0x0200 or 0x0400 or 0x0800;
}
