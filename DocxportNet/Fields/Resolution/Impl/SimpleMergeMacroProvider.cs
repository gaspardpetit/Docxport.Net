using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DocxportNet.Fields.Resolution.Impl;

public abstract class SimpleMergeMacroProvider : DxpMergeMacroProvider
{
    public enum GreetingStyle
    {
        Salutation,
        DualFormSalutation,
        NameOnly
    }

    public enum LocalityOrder
    {
        CityStatePostal,
        PostalCityState,
        StateCityPostal,
        PostalStateCity
    }

    protected abstract string LanguageCode { get; }
    protected abstract string Salutation { get; }
    protected abstract string DefaultGreeting { get; }
    protected virtual string? DualFormSalutation => null;
    protected virtual GreetingStyle Greeting => GreetingStyle.Salutation;
    protected virtual bool IncludeTitleInGreeting => true;
    protected virtual bool LastNameFirst => false;
    protected virtual string? NameSuffix => null;
    protected virtual string GreetingPunctuation => ",";
    protected virtual string CityStateSeparator => ", ";
    protected virtual string LocalityPostalSeparator => " ";
    protected virtual LocalityOrder LocalityLayout => LocalityOrder.CityStatePostal;

    public override bool CanHandle(CultureInfo culture)
        => culture.TwoLetterISOLanguageName.Equals(LanguageCode, StringComparison.OrdinalIgnoreCase);

    public override string? Resolve(string macroName, IDxpMergeRecordCursor cursor, CultureInfo culture)
    {
        switch (macroName.ToUpperInvariant())
        {
            case "GREETINGLINE":
                return ResolveGreetingLine(cursor, culture);
            case "ADDRESSBLOCK":
                return ResolveAddressBlock(cursor, culture);
            default:
                return null;
        }
    }

    protected virtual string ResolveGreetingLine(IDxpMergeRecordCursor cursor, CultureInfo culture)
    {
        var title = Normalize(GetFieldString(cursor, "Title", culture));
        var firstName = Normalize(GetFieldString(cursor, "FirstName", culture));
        var lastName = Normalize(GetFieldString(cursor, "LastName", culture));

        string? name = null;
        if (LastNameFirst)
        {
            name = JoinNonEmpty(" ", lastName, firstName);
        }
        else
        {
            if (lastName is { Length: > 0 })
                name = lastName;
            else if (firstName is { Length: > 0 })
                name = firstName;
        }

        if (name is { Length: > 0 })
        {
            var safeName = name;
            var composedName = IncludeTitleInGreeting ? JoinNonEmpty(" ", title, safeName) ?? safeName : safeName;
            return Greeting switch {
                GreetingStyle.DualFormSalutation => BuildGreeting(DualFormSalutation ?? Salutation, composedName),
                GreetingStyle.NameOnly => BuildNameOnlyGreeting(composedName),
                _ => BuildGreeting(Salutation, composedName)
            };
        }

        return DefaultGreeting;
    }

    protected virtual string? ResolveAddressBlock(IDxpMergeRecordCursor cursor, CultureInfo culture)
    {
        var company = Normalize(GetFieldString(cursor, "Company", culture));
        var address1 = Normalize(GetFieldString(cursor, "Address1", culture));
        var address2 = Normalize(GetFieldString(cursor, "Address2", culture));
        var city = Normalize(GetFieldString(cursor, "City", culture));
        var state = Normalize(GetFieldString(cursor, "State", culture));
        var postalCode = Normalize(GetFieldString(cursor, "PostalCode", culture));
        var country = Normalize(GetFieldString(cursor, "Country", culture));

        var lines = new List<string>();
        if (company is { Length: > 0 })
            lines.Add(company);
        if (address1 is { Length: > 0 })
            lines.Add(address1);
        if (address2 is { Length: > 0 })
            lines.Add(address2);

        var locality = ComposeLocality(city, state, postalCode);
        if (locality is { Length: > 0 })
            lines.Add(locality);

        if (country is { Length: > 0 })
            lines.Add(country);

        if (lines.Count == 0)
            return null;

        return string.Join("\n", lines);
    }

    private string? ComposeLocality(string? city, string? state, string? postalCode)
    {
        return LocalityLayout switch {
            LocalityOrder.PostalCityState => JoinNonEmpty(" ", postalCode, city, state),
            LocalityOrder.StateCityPostal => JoinNonEmpty(" ", state, city, postalCode),
            LocalityOrder.PostalStateCity => JoinNonEmpty(" ", postalCode, state, city),
            _ => ComposeCityStatePostal(city, state, postalCode)
        };
    }

    private string? ComposeCityStatePostal(string? city, string? state, string? postalCode)
    {
        string? locality = null;
        if (!string.IsNullOrEmpty(city) && !string.IsNullOrEmpty(state))
            locality = $"{city}{CityStateSeparator}{state}";
        else if (!string.IsNullOrEmpty(city))
            locality = city;
        else if (!string.IsNullOrEmpty(state))
            locality = state;

        if (!string.IsNullOrEmpty(postalCode))
        {
            if (!string.IsNullOrEmpty(locality))
                locality = $"{locality}{LocalityPostalSeparator}{postalCode}";
            else
                locality = postalCode;
        }

        return locality;
    }

    protected static string? JoinNonEmpty(string separator, params string?[] values)
    {
        var parts = values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value!.Trim());
        var joined = string.Join(separator, parts);
        return string.IsNullOrWhiteSpace(joined) ? null : joined;
    }

    private string BuildGreeting(string salutation, string name)
    {
        var greeting = $"{salutation} {name}".Trim();
        return $"{greeting}{GreetingPunctuation}";
    }

    private string BuildNameOnlyGreeting(string name)
    {
        var greeting = $"{name}{NameSuffix}".Trim();
        return $"{greeting}{GreetingPunctuation}";
    }

    protected static string? Normalize(string? value)
    {
        if (value == null)
            return null;
        var trimmed = value.Trim();
        return trimmed.Length == 0 ? null : trimmed;
    }
}
