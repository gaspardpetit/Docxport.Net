namespace DocxportNet.Omml;

internal sealed class OmmlDocument
{
    public OmmlDocument(bool isDisplay, IReadOnlyList<OmmlNode> children)
    {
        IsDisplay = isDisplay;
        Children = children;
    }

    public bool IsDisplay { get; }
    public IReadOnlyList<OmmlNode> Children { get; }
}

internal abstract class OmmlNode
{
    protected OmmlNode(string path) => Path = path;
    public string Path { get; }
}

internal sealed class OmmlSequence : OmmlNode
{
    public OmmlSequence(string path, IReadOnlyList<OmmlNode> children) : base(path) => Children = children;
    public IReadOnlyList<OmmlNode> Children { get; }
}

internal enum OmmlTokenKind { Identifier, Number, Operator, Text }
internal enum OmmlMathScript { Default, Roman, Script, Fraktur, DoubleStruck, SansSerif, Monospace }
internal enum OmmlMathStyle { Default, Plain, Bold, Italic, BoldItalic }

internal sealed class OmmlToken
{
    public OmmlToken(OmmlTokenKind kind, string value) { Kind = kind; Value = value; }
    public OmmlTokenKind Kind { get; }
    public string Value { get; }
}

internal sealed class OmmlRun : OmmlNode
{
    public OmmlRun(string path, IReadOnlyList<OmmlToken> tokens, OmmlMathScript script,
        OmmlMathStyle style, bool literal, bool normal, bool alignment, string? language, bool rightToLeft)
        : base(path)
    {
        Tokens = tokens; Script = script; Style = style; Literal = literal; Normal = normal;
        Alignment = alignment; Language = language; RightToLeft = rightToLeft;
    }
    public IReadOnlyList<OmmlToken> Tokens { get; }
    public OmmlMathScript Script { get; }
    public OmmlMathStyle Style { get; }
    public bool Literal { get; }
    public bool Normal { get; }
    public bool Alignment { get; }
    public string? Language { get; }
    public bool RightToLeft { get; }
}

internal enum OmmlFractionType { Bar, Skewed, Linear, NoBar }
internal enum OmmlScriptType { Subscript, Superscript, SubSup, PreSubSup }

internal sealed class OmmlFraction : OmmlNode
{
    public OmmlFraction(string path, OmmlFractionType type, OmmlSequence numerator,
        OmmlSequence denominator, bool hasControlProperties) : base(path)
    { Type = type; Numerator = numerator; Denominator = denominator; HasControlProperties = hasControlProperties; }
    public OmmlFractionType Type { get; }
    public OmmlSequence Numerator { get; }
    public OmmlSequence Denominator { get; }
    public bool HasControlProperties { get; }
}

internal sealed class OmmlRadical : OmmlNode
{
    public OmmlRadical(string path, OmmlSequence radicand, OmmlSequence degree,
        bool hasDegree, bool degreeHidden, bool hasControlProperties) : base(path)
    { Radicand = radicand; Degree = degree; HasDegree = hasDegree; DegreeHidden = degreeHidden; HasControlProperties = hasControlProperties; }
    public OmmlSequence Radicand { get; }
    public OmmlSequence Degree { get; }
    public bool HasDegree { get; }
    public bool DegreeHidden { get; }
    public bool HasControlProperties { get; }
}

internal sealed class OmmlScript : OmmlNode
{
    public OmmlScript(string path, OmmlScriptType type, OmmlSequence @base,
        OmmlSequence subscript, OmmlSequence superscript, bool alignScripts,
        bool hasControlProperties) : base(path)
    { Type = type; Base = @base; Subscript = subscript; Superscript = superscript; AlignScripts = alignScripts; HasControlProperties = hasControlProperties; }
    public OmmlScriptType Type { get; }
    public OmmlSequence Base { get; }
    public OmmlSequence Subscript { get; }
    public OmmlSequence Superscript { get; }
    public bool AlignScripts { get; }
    public bool HasControlProperties { get; }
}

internal enum OmmlDelimiterShape { Centered, Match }
internal enum OmmlVerticalPosition { Top, Bottom }
internal enum OmmlDecorationType { Accent, Bar, GroupCharacter }

internal sealed class OmmlDelimiter : OmmlNode
{
    public OmmlDelimiter(string path, string begin, string separator, string end,
        bool grow, OmmlDelimiterShape shape, IReadOnlyList<OmmlSequence> arguments,
        bool hasControlProperties) : base(path)
    { Begin = begin; Separator = separator; End = end; Grow = grow; Shape = shape; Arguments = arguments; HasControlProperties = hasControlProperties; }
    public string Begin { get; }
    public string Separator { get; }
    public string End { get; }
    public bool Grow { get; }
    public OmmlDelimiterShape Shape { get; }
    public IReadOnlyList<OmmlSequence> Arguments { get; }
    public bool HasControlProperties { get; }
}

internal sealed class OmmlDecoration : OmmlNode
{
    public OmmlDecoration(string path, OmmlDecorationType type, string character,
        OmmlVerticalPosition position, OmmlVerticalPosition verticalJustification,
        OmmlSequence argument, bool hasControlProperties) : base(path)
    { Type = type; Character = character; Position = position; VerticalJustification = verticalJustification; Argument = argument; HasControlProperties = hasControlProperties; }
    public OmmlDecorationType Type { get; }
    public string Character { get; }
    public OmmlVerticalPosition Position { get; }
    public OmmlVerticalPosition VerticalJustification { get; }
    public OmmlSequence Argument { get; }
    public bool HasControlProperties { get; }
}

internal enum OmmlLimitType { Lower, Upper }

internal sealed class OmmlFunction : OmmlNode
{
    public OmmlFunction(string path, OmmlSequence name, OmmlSequence argument,
        bool hasControlProperties) : base(path)
    { Name = name; Argument = argument; HasControlProperties = hasControlProperties; }
    public OmmlSequence Name { get; }
    public OmmlSequence Argument { get; }
    public bool HasControlProperties { get; }
}

internal sealed class OmmlLimit : OmmlNode
{
    public OmmlLimit(string path, OmmlLimitType type, OmmlSequence @base,
        OmmlSequence limit, bool hasControlProperties) : base(path)
    { Type = type; Base = @base; Limit = limit; HasControlProperties = hasControlProperties; }
    public OmmlLimitType Type { get; }
    public OmmlSequence Base { get; }
    public OmmlSequence Limit { get; }
    public bool HasControlProperties { get; }
}

internal sealed class OmmlNary : OmmlNode
{
    public OmmlNary(string path, string character, DxpOmmlLimitLocation? limitLocation,
        bool grow, bool hideSubscript, bool hideSuperscript, OmmlSequence subscript,
        OmmlSequence superscript, OmmlSequence argument, bool hasControlProperties) : base(path)
    {
        Character = character; LimitLocation = limitLocation; Grow = grow;
        HideSubscript = hideSubscript; HideSuperscript = hideSuperscript;
        Subscript = subscript; Superscript = superscript; Argument = argument;
        HasControlProperties = hasControlProperties;
    }
    public string Character { get; }
    public DxpOmmlLimitLocation? LimitLocation { get; }
    public bool Grow { get; }
    public bool HideSubscript { get; }
    public bool HideSuperscript { get; }
    public OmmlSequence Subscript { get; }
    public OmmlSequence Superscript { get; }
    public OmmlSequence Argument { get; }
    public bool HasControlProperties { get; }
}

internal enum OmmlHorizontalAlignment { Left, Center, Right }
internal enum OmmlVerticalAlignment { Top, Center, Bottom }

internal sealed class OmmlMatrixColumn
{
    public OmmlMatrixColumn(int count, OmmlHorizontalAlignment alignment)
    { Count = count; Alignment = alignment; }
    public int Count { get; }
    public OmmlHorizontalAlignment Alignment { get; }
}

internal sealed class OmmlMatrixRow
{
    public OmmlMatrixRow(IReadOnlyList<OmmlSequence> cells) => Cells = cells;
    public IReadOnlyList<OmmlSequence> Cells { get; }
}

internal sealed class OmmlMatrix : OmmlNode
{
    public OmmlMatrix(string path, IReadOnlyList<OmmlMatrixRow> rows,
        IReadOnlyList<OmmlMatrixColumn> columns, OmmlVerticalAlignment baseJustification,
        bool placeholdersHidden, uint rowSpacing, int rowSpacingRule, uint columnSpacing,
        uint columnGap, int columnGapRule, bool hasControlProperties) : base(path)
    {
        Rows = rows; Columns = columns; BaseJustification = baseJustification;
        PlaceholdersHidden = placeholdersHidden; RowSpacing = rowSpacing;
        RowSpacingRule = rowSpacingRule; ColumnSpacing = columnSpacing;
        ColumnGap = columnGap; ColumnGapRule = columnGapRule;
        HasControlProperties = hasControlProperties;
    }
    public IReadOnlyList<OmmlMatrixRow> Rows { get; }
    public IReadOnlyList<OmmlMatrixColumn> Columns { get; }
    public OmmlVerticalAlignment BaseJustification { get; }
    public bool PlaceholdersHidden { get; }
    public uint RowSpacing { get; }
    public int RowSpacingRule { get; }
    public uint ColumnSpacing { get; }
    public uint ColumnGap { get; }
    public int ColumnGapRule { get; }
    public bool HasControlProperties { get; }
}

internal sealed class OmmlEquationArray : OmmlNode
{
    public OmmlEquationArray(string path, IReadOnlyList<OmmlSequence> rows,
        OmmlVerticalAlignment baseJustification, bool maxDistribution,
        bool objectDistribution, uint rowSpacing, int rowSpacingRule,
        bool hasControlProperties) : base(path)
    {
        Rows = rows; BaseJustification = baseJustification;
        MaxDistribution = maxDistribution; ObjectDistribution = objectDistribution;
        RowSpacing = rowSpacing; RowSpacingRule = rowSpacingRule;
        HasControlProperties = hasControlProperties;
    }
    public IReadOnlyList<OmmlSequence> Rows { get; }
    public OmmlVerticalAlignment BaseJustification { get; }
    public bool MaxDistribution { get; }
    public bool ObjectDistribution { get; }
    public uint RowSpacing { get; }
    public int RowSpacingRule { get; }
    public bool HasControlProperties { get; }
}

internal sealed class OmmlUnsupported : OmmlNode
{
    public OmmlUnsupported(string path, string elementName, string visibleText) : base(path)
    {
        ElementName = elementName;
        VisibleText = visibleText;
    }

    public string ElementName { get; }
    public string VisibleText { get; }
}
