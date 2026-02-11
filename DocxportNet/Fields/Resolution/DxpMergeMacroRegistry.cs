using DocxportNet.Fields.Resolution.Impl;
using System.Globalization;

namespace DocxportNet.Fields.Resolution;

public sealed class DxpMergeMacroRegistry
{
    private readonly List<DxpIMergeMacroProvider> _providers = new();

    public static DxpMergeMacroRegistry Default { get; } = CreateDefault();

    public void Register(DxpIMergeMacroProvider provider)
    {
        if (provider == null)
            throw new ArgumentNullException(nameof(provider));
        _providers.Insert(0, provider);
    }

    public DxpIMergeMacroProvider Resolve(CultureInfo culture)
    {
        foreach (var provider in _providers)
        {
            if (provider.CanHandle(culture))
                return provider;
        }

        foreach (var provider in _providers)
        {
            if (provider.CanHandle(new CultureInfo(culture.TwoLetterISOLanguageName)))
                return provider;
        }

        return _providers[0];
    }

    private static DxpMergeMacroRegistry CreateDefault()
    {
        var registry = new DxpMergeMacroRegistry();
        registry._providers.Add(new EnglishCanadaMergeMacroProvider());
        registry._providers.Add(new FrenchCanadaMergeMacroProvider());
        registry._providers.Add(new EnglishMergeMacroProvider());
        registry._providers.Add(new FrenchMergeMacroProvider());
        registry._providers.Add(new JapaneseMergeMacroProvider());
        registry._providers.Add(new ThaiMergeMacroProvider());
        registry._providers.Add(new GermanMergeMacroProvider());
        registry._providers.Add(new ChineseSimplifiedMergeMacroProvider());
        registry._providers.Add(new SpanishMergeMacroProvider());
        registry._providers.Add(new ItalianMergeMacroProvider());
        registry._providers.Add(new PortugueseMergeMacroProvider());
        registry._providers.Add(new DanishMergeMacroProvider());
        registry._providers.Add(new FinnishMergeMacroProvider());
        registry._providers.Add(new GreekMergeMacroProvider());
        return registry;
    }
}
