namespace DocxportNet.Doc;

/// <summary>The minute-resolution DTTM used by binary DOC revision marks.</summary>
internal static class DocRevisionDate
{
    public static DateTime? Decode(uint value)
    {
        var minute = (int)(value & 0x3F);
        var hour = (int)((value >> 6) & 0x1F);
        var day = (int)((value >> 11) & 0x1F);
        var month = (int)((value >> 16) & 0x0F);
        var year = 1900 + (int)((value >> 20) & 0x1FF);
        if (day == 0 || month == 0) return null;
        if (minute > 59 || hour > 23 || month > 12 ||
            year > 9999 || day > DateTime.DaysInMonth(year, month))
            throw new InvalidDataException("A DOC revision timestamp is invalid.");
        return new DateTime(year, month, day, hour, minute, 0, DateTimeKind.Utc);
    }

    public static uint Encode(DateTime value)
    {
        var utc = value.Kind == DateTimeKind.Utc ? value : value.ToUniversalTime();
        if (utc.Year < 1900 || utc.Year > 2411)
            throw new InvalidDataException("A DOC revision timestamp is out of range.");
        return (uint)(utc.Minute | (utc.Hour << 6) | (utc.Day << 11) |
            (utc.Month << 16) | ((utc.Year - 1900) << 20) |
            ((int)utc.DayOfWeek << 29));
    }
}
