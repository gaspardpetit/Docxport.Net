using System.Collections.Generic;
using System.Globalization;

namespace DocxportNet.Fields.Resolution.Impl;

public sealed class ItalianMergeMacroProvider : SimpleMergeMacroProvider
{
    protected override string LanguageCode => "it";
    protected override string Salutation => "Gentile";
    protected override string DefaultGreeting => "Buongiorno,";
    protected override LocalityOrder LocalityLayout => LocalityOrder.PostalCityState;

    protected override string? ResolveAddressBlock(IDxpMergeRecordCursor cursor, CultureInfo culture)
    {
        var company = Normalize(GetFieldString(cursor, "Company", culture));
        var address1 = Normalize(GetFieldString(cursor, "Address1", culture));
        var address2 = Normalize(GetFieldString(cursor, "Address2", culture));
        var city = Normalize(GetFieldString(cursor, "City", culture));
        var province = Normalize(GetFieldString(cursor, "State", culture));
        var postalCode = Normalize(GetFieldString(cursor, "PostalCode", culture));
        var country = Normalize(GetFieldString(cursor, "Country", culture));

        var lines = new List<string>();
        if (company is { Length: > 0 })
            lines.Add(company);
        if (address1 is { Length: > 0 })
            lines.Add(address1);
        if (address2 is { Length: > 0 })
            lines.Add(address2);

        var locality = BuildLocality(postalCode, city, province);
        if (locality is { Length: > 0 })
            lines.Add(locality);

        if (country is { Length: > 0 })
            lines.Add(country);

        if (lines.Count == 0)
            return null;

        return string.Join("\n", lines);
    }

    private static string? BuildLocality(string? postalCode, string? city, string? province)
    {
        var baseLocality = JoinNonEmpty(" ", postalCode, city);
        if (baseLocality is not { Length: > 0 })
            return province;
        if (province is { Length: > 0 })
            return $"{baseLocality} ({province})";
        return baseLocality;
    }
}
