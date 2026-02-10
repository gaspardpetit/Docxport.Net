using DocxportNet.API;

namespace DocxportNet.Middleware;

public static class DxpVisitorMiddleware
{
    public static DxpIVisitor Chain(DxpIVisitor terminal, params Func<DxpIVisitor, DxpIVisitor>[] factories)
    {
        if (terminal == null)
            throw new ArgumentNullException(nameof(terminal));

        if (factories == null || factories.Length == 0)
            return terminal;

        DxpIVisitor current = terminal;
        for (int i = factories.Length - 1; i >= 0; i--)
            current = factories[i](current);

        return current;
    }
}
