using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;

public class DxpRubyContext : DxpIRubyContext
{
    public Ruby Ruby { get; }
    public RubyProperties? Properties { get; }

    public DxpRubyContext(Ruby ruby, RubyProperties? properties)
    {
        Ruby = ruby;
        Properties = properties;
    }
}
