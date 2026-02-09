namespace DocxportNet.Fields.Resolution;

public enum DxpFieldValueKindHint
{
    Any,
    DocVariable,
    DocumentProperty,
    Bookmark,
    MergeField
}

public interface IDxpFieldValueResolver
{
    Task<DxpFieldValue?> ResolveAsync(string name, DxpFieldValueKindHint kind, DxpFieldEvalContext context);
}

public sealed class DxpChainedFieldValueResolver : IDxpFieldValueResolver
{
    private readonly List<IDxpFieldValueResolver> _resolvers = new();

    public DxpChainedFieldValueResolver(params IDxpFieldValueResolver[] resolvers)
    {
        _resolvers.AddRange(resolvers);
    }

    public void Add(IDxpFieldValueResolver resolver)
    {
        _resolvers.Add(resolver);
    }

    public async Task<DxpFieldValue?> ResolveAsync(string name, DxpFieldValueKindHint kind, DxpFieldEvalContext context)
    {
        foreach (var resolver in _resolvers)
        {
            var value = await resolver.ResolveAsync(name, kind, context);
            if (value != null)
                return value;
        }
        return null;
    }
}

public sealed class DxpContextFieldValueResolver : IDxpFieldValueResolver
{
    public Task<DxpFieldValue?> ResolveAsync(string name, DxpFieldValueKindHint kind, DxpFieldEvalContext context)
    {
        if ((kind == DxpFieldValueKindHint.Any || kind == DxpFieldValueKindHint.Bookmark) &&
            context.TryGetBookmarkNodes(name, out var bmNodes))
        {
            var text = bmNodes.ToPlainText();
            return Task.FromResult<DxpFieldValue?>(new DxpFieldValue(text));
        }

        if ((kind == DxpFieldValueKindHint.Any || kind == DxpFieldValueKindHint.DocVariable) &&
            context.TryGetDocVariableNodes(name, out var dvNodes))
        {
            var text = dvNodes.ToPlainText();
            return Task.FromResult<DxpFieldValue?>(new DxpFieldValue(text));
        }

        if ((kind == DxpFieldValueKindHint.Any || kind == DxpFieldValueKindHint.DocVariable) &&
            context.TryGetDocVariable(name, out var dv) && dv != null)
            return Task.FromResult<DxpFieldValue?>(new DxpFieldValue(dv));

        if ((kind == DxpFieldValueKindHint.Any || kind == DxpFieldValueKindHint.DocumentProperty) &&
            context.TryGetDocumentPropertyValue(name, out var dpValue))
            return Task.FromResult<DxpFieldValue?>(dpValue);

        if ((kind == DxpFieldValueKindHint.Any || kind == DxpFieldValueKindHint.DocumentProperty) &&
            context.TryGetDocumentProperty(name, out var dp) && dp != null)
            return Task.FromResult<DxpFieldValue?>(new DxpFieldValue(dp));

        return Task.FromResult<DxpFieldValue?>(null);
    }
}

public sealed class DxpDelegateFieldValueResolver : IDxpFieldValueResolver
{
    private readonly DxpFieldEvalDelegates _delegates;

    public DxpDelegateFieldValueResolver(DxpFieldEvalDelegates delegates)
    {
        _delegates = delegates;
    }

    public async Task<DxpFieldValue?> ResolveAsync(string name, DxpFieldValueKindHint kind, DxpFieldEvalContext context)
    {
        if ((kind == DxpFieldValueKindHint.Any || kind == DxpFieldValueKindHint.DocVariable) &&
            _delegates.ResolveDocVariableAsync != null)
        {
            var value = await _delegates.ResolveDocVariableAsync(name, context);
            if (value != null)
                return value;
        }

        if ((kind == DxpFieldValueKindHint.Any || kind == DxpFieldValueKindHint.DocumentProperty) &&
            _delegates.ResolveDocumentPropertyAsync != null)
        {
            var value = await _delegates.ResolveDocumentPropertyAsync(name, context);
            if (value != null)
                return value;
        }

        if ((kind == DxpFieldValueKindHint.Any || kind == DxpFieldValueKindHint.MergeField) &&
            _delegates.ResolveMergeFieldAsync != null)
        {
            var value = await _delegates.ResolveMergeFieldAsync(name, context);
            if (value != null)
                return value;
        }

        return null;
    }
}
