using System.Collections.Generic;
using System.Globalization;

namespace DocxportNet.Fields.Resolution.Impl;

public sealed class EnglishCanadaMergeMacroProvider : SimpleMergeMacroProvider
{
    protected override string LanguageCode => "en";
    protected override string Salutation => "Dear";
    protected override string DefaultGreeting => "Hello,";
    protected override string CityStateSeparator => " ";
    protected override string LocalityPostalSeparator => "  ";

    public override bool CanHandle(CultureInfo culture)
        => culture.Name.Equals("en-CA", StringComparison.OrdinalIgnoreCase);

    protected override string? ResolveAddressBlock(IDxpMergeRecordCursor cursor, CultureInfo culture)
    {
        var company = Normalize(GetFieldString(cursor, "Company", culture));
        var address1 = Normalize(GetFieldString(cursor, "Address1", culture));
        var address2 = Normalize(GetFieldString(cursor, "Address2", culture));
        var city = Normalize(GetFieldString(cursor, "City", culture));
        var state = Normalize(GetFieldString(cursor, "State", culture));
        var postalCode = Normalize(GetFieldString(cursor, "PostalCode", culture));
        var country = Normalize(GetFieldString(cursor, "Country", culture));

        if (city is { Length: > 0 })
            city = city.ToUpperInvariant();
        if (state is { Length: > 0 })
            state = state.ToUpperInvariant();
        if (postalCode is { Length: > 0 })
            postalCode = postalCode.ToUpperInvariant();
        if (country is { Length: > 0 })
            country = country.ToUpperInvariant();

        var lines = new List<string>();
        if (company is { Length: > 0 })
            lines.Add(company);
        if (address1 is { Length: > 0 })
            lines.Add(address1);
        if (address2 is { Length: > 0 })
            lines.Add(address2);

        var locality = JoinNonEmpty(" ", city, state);
        if (locality is { Length: > 0 } && postalCode is { Length: > 0 })
            locality = $"{locality}  {postalCode}";
        else if (postalCode is { Length: > 0 })
            locality = postalCode;
        if (locality is { Length: > 0 })
            lines.Add(locality);

        if (country is { Length: > 0 })
            lines.Add(country);

        if (lines.Count == 0)
            return null;

        return string.Join("\n", lines);
    }
}
