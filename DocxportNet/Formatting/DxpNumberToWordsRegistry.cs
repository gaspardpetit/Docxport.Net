using DocxportNet.Formatting.Impl;
using System.Globalization;

namespace DocxportNet.Formatting;

public sealed class DxpNumberToWordsRegistry
{
    private readonly List<IDxpNumberToWordsProvider> _providers = new();

    public static DxpNumberToWordsRegistry Default { get; } = CreateDefault();

    public void Register(IDxpNumberToWordsProvider provider)
    {
        if (provider == null)
            throw new ArgumentNullException(nameof(provider));
        _providers.Insert(0, provider);
    }

    public IDxpNumberToWordsProvider Resolve(CultureInfo culture)
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

    private static DxpNumberToWordsRegistry CreateDefault()
    {
        var registry = new DxpNumberToWordsRegistry();
        registry._providers.Add(new EnglishNumberToWordsProvider());
        registry._providers.Add(new FrenchNumberToWordsProvider());
        registry._providers.Add(new JapaneseNumberToWordsProvider());
        registry._providers.Add(new ThaiNumberToWordsProvider());
        registry._providers.Add(new GermanNumberToWordsProvider());
        registry._providers.Add(new ChineseSimplifiedNumberToWordsProvider());
        registry._providers.Add(new SpanishNumberToWordsProvider());
        registry._providers.Add(new ItalianNumberToWordsProvider());
        registry._providers.Add(new PortugueseNumberToWordsProvider());
        registry._providers.Add(new DanishNumberToWordsProvider());
        registry._providers.Add(new FinnishNumberToWordsProvider());
        registry._providers.Add(new GreekNumberToWordsProvider());
        return registry;
    }
}
