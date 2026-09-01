using System.Text;
using System.Xml.Linq;

namespace DocxportNet.Omml;

internal static class OmmlWriter
{
    private const string MathMlNamespace = "http://www.w3.org/1998/Math/MathML";

    public static DxpOmmlConversionResult Write(
        OmmlDocument document,
        DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions options)
    {
        bool isDisplay = options.Display ?? document.IsDisplay;
        List<DxpOmmlDiagnostic> diagnostics = new();
        string output = format switch
        {
            DxpOmmlOutputFormat.MathMl => WriteMathMl(document, isDisplay, options, diagnostics),
            DxpOmmlOutputFormat.Latex => WriteTextual(document, format, isDisplay, options, diagnostics, EscapeLatex),
            DxpOmmlOutputFormat.UnicodeMath => WriteTextual(document, format, isDisplay, options, diagnostics, static value => value),
            DxpOmmlOutputFormat.Text => WriteTextual(document, format, isDisplay, options, diagnostics, static value => value),
            _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unknown OMML output format."),
        };
        return new DxpOmmlConversionResult(output, format, isDisplay, diagnostics.AsReadOnly());
    }

    private static string WriteMathMl(
        OmmlDocument document,
        bool isDisplay,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics)
    {
        XNamespace math = MathMlNamespace;
        XElement row = new(math + "mrow");
        foreach (OmmlNode node in document.Children)
            AppendMathMl(row, node, options.SmallFractions && !isDisplay, isDisplay, options, diagnostics);

        XElement root = new(
            math + "math",
            new XAttribute("display", isDisplay ? "block" : "inline"),
            row);
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static void AppendMathMl(
        XElement parent,
        OmmlNode node,
        bool compactFractions,
        bool isDisplay,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics)
    {
        XNamespace math = MathMlNamespace;
        if (node is OmmlSequence sequence)
        {
            XElement row = new(math + "mrow");
            foreach (OmmlNode child in sequence.Children)
                AppendMathMl(row, child, compactFractions, isDisplay, options, diagnostics);
            parent.Add(row);
            return;
        }

        if (node is OmmlRun run)
        {
            XElement container = parent;
            string? variant = MathVariant(run);
            if (variant != null || run.Language != null || run.RightToLeft)
            {
                container = new XElement(math + "mstyle");
                if (variant != null) container.SetAttributeValue("mathvariant", variant);
                if (run.Language != null) container.SetAttributeValue(XNamespace.Xml + "lang", run.Language);
                if (run.RightToLeft) container.SetAttributeValue("dir", "rtl");
                parent.Add(container);
            }
            if (run.Alignment) container.Add(new XElement(math + "malignmark"));
            foreach (OmmlToken token in run.Tokens)
            {
                if (token.Value == "\u200B") container.Add(new XElement(math + "mspace", new XAttribute("width", "0")));
                else container.Add(new XElement(math + (token.Kind switch { OmmlTokenKind.Identifier => "mi", OmmlTokenKind.Number => "mn", OmmlTokenKind.Operator => "mo", _ => "mtext" }), token.Value));
            }
            return;
        }

        if (node is OmmlFraction fraction)
        {
            XElement rendered;
            if (fraction.Type == OmmlFractionType.Linear)
            {
                rendered = new XElement(math + "mrow");
                AppendMathMl(rendered, fraction.Numerator, compactFractions, isDisplay, options, diagnostics);
                rendered.Add(new XElement(math + "mo", "/"));
                AppendMathMl(rendered, fraction.Denominator, compactFractions, isDisplay, options, diagnostics);
            }
            else
            {
                rendered = new XElement(math + "mfrac");
                if (fraction.Type == OmmlFractionType.Skewed) rendered.SetAttributeValue("bevelled", "true");
                if (fraction.Type == OmmlFractionType.NoBar) rendered.SetAttributeValue("linethickness", "0");
                AppendMathMl(rendered, fraction.Numerator, compactFractions, isDisplay, options, diagnostics);
                AppendMathMl(rendered, fraction.Denominator, compactFractions, isDisplay, options, diagnostics);
            }
            if (compactFractions)
                parent.Add(new XElement(math + "mstyle", new XAttribute("displaystyle", "false"), new XAttribute("scriptlevel", "1"), rendered));
            else parent.Add(rendered);
            return;
        }

        if (node is OmmlRadical radical)
        {
            bool indexed = radical.HasDegree && !radical.DegreeHidden;
            XElement rendered = new(math + (indexed ? "mroot" : "msqrt"));
            AppendMathMl(rendered, radical.Radicand, compactFractions, isDisplay, options, diagnostics);
            if (indexed) AppendMathMl(rendered, radical.Degree, compactFractions, isDisplay, options, diagnostics);
            parent.Add(rendered);
            return;
        }

        if (node is OmmlScript script)
        {
            XElement rendered = new(math + (script.Type switch
            {
                OmmlScriptType.Subscript => "msub",
                OmmlScriptType.Superscript => "msup",
                OmmlScriptType.SubSup => "msubsup",
                _ => "mmultiscripts",
            }));
            AppendMathMl(rendered, script.Base, compactFractions, isDisplay, options, diagnostics);
            if (script.Type == OmmlScriptType.PreSubSup)
            {
                rendered.Add(new XElement(math + "mprescripts"));
                AppendMathMl(rendered, script.Subscript, compactFractions, isDisplay, options, diagnostics);
                AppendMathMl(rendered, script.Superscript, compactFractions, isDisplay, options, diagnostics);
            }
            else
            {
                if (script.Type != OmmlScriptType.Superscript) AppendMathMl(rendered, script.Subscript, compactFractions, isDisplay, options, diagnostics);
                if (script.Type != OmmlScriptType.Subscript) AppendMathMl(rendered, script.Superscript, compactFractions, isDisplay, options, diagnostics);
            }
            if (script.AlignScripts) rendered.SetAttributeValue("data-omml-align-scripts", "true");
            parent.Add(rendered);
            return;
        }

        if (node is OmmlDelimiter delimiter)
        {
            XElement row = new(math + "mrow", new XAttribute("data-omml-shape", delimiter.Shape == OmmlDelimiterShape.Match ? "match" : "centered"));
            AddFence(row, delimiter.Begin, delimiter.Grow, true);
            for (int i = 0; i < delimiter.Arguments.Count; i++)
            {
                if (i != 0 && delimiter.Separator.Length != 0)
                    row.Add(new XElement(math + "mo", new XAttribute("separator", "true"), delimiter.Separator));
                AppendMathMl(row, delimiter.Arguments[i], compactFractions, isDisplay, options, diagnostics);
            }
            AddFence(row, delimiter.End, delimiter.Grow, false);
            parent.Add(row);
            return;
        }

        if (node is OmmlDecoration decoration)
        {
            bool above = decoration.Position == OmmlVerticalPosition.Top;
            XElement rendered = new(math + (above ? "mover" : "munder"),
                new XAttribute("data-omml-vertical-justification",
                    decoration.VerticalJustification == OmmlVerticalPosition.Top ? "top" : "bot"));
            if (decoration.Type != OmmlDecorationType.GroupCharacter)
                rendered.SetAttributeValue(above ? "accent" : "accentunder",
                    decoration.Type == OmmlDecorationType.Accent ? "true" : "false");
            AppendMathMl(rendered, decoration.Argument, compactFractions, isDisplay, options, diagnostics);
            string character = decoration.Type == OmmlDecorationType.Accent
                ? MathMlAccentCharacter(decoration.Character)
                : decoration.Character;
            rendered.Add(new XElement(math + "mo", new XAttribute("stretchy", "true"), character));
            parent.Add(rendered);
            return;
        }

        if (node is OmmlFunction function)
        {
            XElement row = new(math + "mrow");
            AppendMathMl(row, function.Name, compactFractions, isDisplay, options, diagnostics);
            row.Add(new XElement(math + "mo", new XAttribute("form", "infix"), "⁡"));
            AppendMathMl(row, function.Argument, compactFractions, isDisplay, options, diagnostics);
            parent.Add(row);
            return;
        }

        if (node is OmmlLimit limit)
        {
            XElement rendered = new(math + (limit.Type == OmmlLimitType.Lower ? "munder" : "mover"));
            AppendMathMl(rendered, limit.Base, compactFractions, isDisplay, options, diagnostics);
            AppendMathMl(rendered, limit.Limit, compactFractions, isDisplay, options, diagnostics);
            parent.Add(rendered);
            return;
        }

        if (node is OmmlNary nary)
        {
            DxpOmmlLimitLocation location = NaryLimitLocation(nary, isDisplay, options);
            bool hasSubscript = !nary.HideSubscript;
            bool hasSuperscript = !nary.HideSuperscript;
            XElement op = new(math + "mo", new XAttribute("largeop", "true"),
                new XAttribute("stretchy", nary.Grow ? "true" : "false"), nary.Character);
            XElement rendered;
            if (hasSubscript && hasSuperscript)
            {
                rendered = new XElement(math + (location == DxpOmmlLimitLocation.UnderOver ? "munderover" : "msubsup"), op);
                AppendMathMl(rendered, nary.Subscript, compactFractions, isDisplay, options, diagnostics);
                AppendMathMl(rendered, nary.Superscript, compactFractions, isDisplay, options, diagnostics);
            }
            else if (hasSubscript)
            {
                rendered = new XElement(math + (location == DxpOmmlLimitLocation.UnderOver ? "munder" : "msub"), op);
                AppendMathMl(rendered, nary.Subscript, compactFractions, isDisplay, options, diagnostics);
            }
            else if (hasSuperscript)
            {
                rendered = new XElement(math + (location == DxpOmmlLimitLocation.UnderOver ? "mover" : "msup"), op);
                AppendMathMl(rendered, nary.Superscript, compactFractions, isDisplay, options, diagnostics);
            }
            else rendered = op;
            XElement row = new(math + "mrow", rendered);
            AppendMathMl(row, nary.Argument, compactFractions, isDisplay, options, diagnostics);
            parent.Add(row);
            return;
        }

        string fallback = ResolveFallback((OmmlUnsupported)node, options, diagnostics);
        if (fallback.Length != 0)
            parent.Add(new XElement(math + "mtext", fallback));
    }

    private static string WriteTextual(
        OmmlDocument document,
        DxpOmmlOutputFormat format,
        bool isDisplay,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics,
        Func<string, string> escape)
    {
        StringBuilder output = new();
        foreach (OmmlNode node in document.Children)
            AppendTextual(output, node, format, isDisplay, options, diagnostics, escape);
        return output.ToString();
    }

    private static void AppendTextual(
        StringBuilder output,
        OmmlNode node,
        DxpOmmlOutputFormat format,
        bool isDisplay,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics,
        Func<string, string> escape)
    {
        if (node is OmmlSequence sequence)
        {
            foreach (OmmlNode child in sequence.Children)
                AppendTextual(output, child, format, isDisplay, options, diagnostics, escape);
            return;
        }

        if (node is OmmlRun run)
        {
            string value = string.Concat(run.Tokens.Select(token => token.Value));
            if (format == DxpOmmlOutputFormat.Latex)
            {
                value = EscapeLatex(value);
                value = ApplyLatexStyle(run, value);
                if (run.Alignment) output.Append('&');
                output.Append(value.Replace("\u200B", "{}"));
            }
            else
            {
                if (run.Alignment && format == DxpOmmlOutputFormat.UnicodeMath) output.Append('&');
                if (format == DxpOmmlOutputFormat.UnicodeMath && UnicodeMathStyle(run) is string command)
                    output.Append($"\\{command}\"{value.Replace("\"", "\\\"")}\"");
                else output.Append(escape(format == DxpOmmlOutputFormat.Text ? value.Replace("\u200B", string.Empty) : value));
            }
            return;
        }

        if (node is OmmlFraction fraction)
        {
            string numerator = RenderTextual(fraction.Numerator, format, isDisplay, options, diagnostics, escape);
            string denominator = RenderTextual(fraction.Denominator, format, isDisplay, options, diagnostics, escape);
            output.Append(format switch
            {
                DxpOmmlOutputFormat.Latex when fraction.Type == OmmlFractionType.Bar => $"\\frac{{{numerator}}}{{{denominator}}}",
                DxpOmmlOutputFormat.Latex when fraction.Type == OmmlFractionType.NoBar => $"\\genfrac{{}}{{}}{{0pt}}{{}}{{{numerator}}}{{{denominator}}}",
                DxpOmmlOutputFormat.Latex => $"{{{numerator}}}/{{{denominator}}}",
                DxpOmmlOutputFormat.UnicodeMath => $"({numerator})/({denominator})",
                _ => $"({numerator})/({denominator})",
            });
            return;
        }

        if (node is OmmlRadical radical)
        {
            string radicand = RenderTextual(radical.Radicand, format, isDisplay, options, diagnostics, escape);
            string degree = RenderTextual(radical.Degree, format, isDisplay, options, diagnostics, escape);
            bool indexed = radical.HasDegree && !radical.DegreeHidden;
            output.Append(format switch
            {
                DxpOmmlOutputFormat.Latex when indexed => $"\\sqrt[{degree}]{{{radicand}}}",
                DxpOmmlOutputFormat.Latex => $"\\sqrt{{{radicand}}}",
                DxpOmmlOutputFormat.UnicodeMath when indexed => $"√({degree}&{radicand})",
                DxpOmmlOutputFormat.UnicodeMath => $"√({radicand})",
                _ when indexed => $"root({degree}, {radicand})",
                _ => $"sqrt({radicand})",
            });
            return;
        }

        if (node is OmmlScript script)
        {
            string @base = RenderTextual(script.Base, format, isDisplay, options, diagnostics, escape);
            string sub = RenderTextual(script.Subscript, format, isDisplay, options, diagnostics, escape);
            string sup = RenderTextual(script.Superscript, format, isDisplay, options, diagnostics, escape);
            if (format == DxpOmmlOutputFormat.Latex)
            {
                string latexBase = @base.Length == 0 ? "{}" : @base;
                output.Append(script.Type switch { OmmlScriptType.Subscript => $"{latexBase}_{{{sub}}}", OmmlScriptType.Superscript => $"{latexBase}^{{{sup}}}", OmmlScriptType.SubSup => $"{latexBase}_{{{sub}}}^{{{sup}}}", _ => $"{{}}_{{{sub}}}^{{{sup}}}{@base}" });
            }
            else if (format == DxpOmmlOutputFormat.UnicodeMath)
                output.Append(script.Type switch { OmmlScriptType.Subscript => $"{@base}_({sub})", OmmlScriptType.Superscript => $"{@base}^({sup})", OmmlScriptType.SubSup => $"{@base}_({sub})^({sup})", _ => $"_({sub})^({sup}) {@base}" });
            else
                output.Append(script.Type switch { OmmlScriptType.Subscript => $"{@base}_({sub})", OmmlScriptType.Superscript => $"{@base}^({sup})", OmmlScriptType.SubSup => $"{@base}_({sub})^({sup})", _ => $"[{sub},{sup}]{@base}" });
            return;
        }

        if (node is OmmlDelimiter delimiter)
        {
            string separator = format == DxpOmmlOutputFormat.Latex ? EscapeLatex(delimiter.Separator) : delimiter.Separator;
            string content = string.Join(separator,
                delimiter.Arguments.Select(argument => RenderTextual(argument, format, isDisplay, options, diagnostics, escape)));
            if (format == DxpOmmlOutputFormat.Latex)
            {
                string begin = LatexDelimiter(delimiter.Begin);
                string end = LatexDelimiter(delimiter.End);
                output.Append(delimiter.Grow
                    ? $"\\left{(begin.Length == 0 ? "." : begin)}{content}\\right{(end.Length == 0 ? "." : end)}"
                    : $"{begin}{content}{end}");
            }
            else output.Append(delimiter.Begin).Append(content).Append(delimiter.End);
            return;
        }

        if (node is OmmlDecoration decoration)
        {
            string argument = RenderTextual(decoration.Argument, format, isDisplay, options, diagnostics, escape);
            if (format == DxpOmmlOutputFormat.Latex)
                output.Append(LatexDecoration(decoration, argument));
            else if (format == DxpOmmlOutputFormat.UnicodeMath)
                output.Append(UnicodeMathDecoration(decoration, argument));
            else output.Append(decoration.Position == OmmlVerticalPosition.Top ? $"{argument} with {decoration.Character} above" : $"{argument} with {decoration.Character} below");
            return;
        }

        if (node is OmmlFunction function)
        {
            string name = RenderTextual(function.Name, format, isDisplay, options, diagnostics, escape);
            string argument = RenderTextual(function.Argument, format, isDisplay, options, diagnostics, escape);
            string? simpleName = SimpleText(function.Name);
            if (format == DxpOmmlOutputFormat.Latex)
            {
                string? command = simpleName == null ? null : LatexFunction(simpleName);
                string functionName = command != null ? "\\" + command
                    : simpleName != null ? $"\\operatorname{{{EscapeLatex(simpleName)}}}"
                    : $"\\mathop{{{name}}}";
                output.Append(functionName).Append('{').Append(argument).Append('}');
            }
            else if (format == DxpOmmlOutputFormat.UnicodeMath)
                output.Append(name).Append('⁡').Append(argument);
            else output.Append(name).Append('(').Append(argument).Append(')');
            return;
        }

        if (node is OmmlLimit limit)
        {
            string @base = RenderTextual(limit.Base, format, isDisplay, options, diagnostics, escape);
            string value = RenderTextual(limit.Limit, format, isDisplay, options, diagnostics, escape);
            string? simpleBase = SimpleText(limit.Base);
            if (format == DxpOmmlOutputFormat.Latex)
            {
                string? command = simpleBase == null ? null : LatexLimitOperator(simpleBase);
                string renderedBase = command == null ? @base : "\\" + command;
                output.Append('{').Append(renderedBase).Append('}')
                    .Append(limit.Type == OmmlLimitType.Lower ? "_{" : "^{").Append(value).Append('}');
            }
            else if (format == DxpOmmlOutputFormat.UnicodeMath)
                output.Append(@base).Append(limit.Type == OmmlLimitType.Lower ? "_(" : "^(").Append(value).Append(')');
            else output.Append(@base).Append(limit.Type == OmmlLimitType.Lower ? " with lower limit " : " with upper limit ").Append(value);
            return;
        }

        if (node is OmmlNary nary)
        {
            string subscript = RenderTextual(nary.Subscript, format, isDisplay, options, diagnostics, escape);
            string superscript = RenderTextual(nary.Superscript, format, isDisplay, options, diagnostics, escape);
            string argument = RenderTextual(nary.Argument, format, isDisplay, options, diagnostics, escape);
            if (format == DxpOmmlOutputFormat.Latex)
            {
                string op = LatexNaryOperator(nary.Character);
                string placement = NaryLimitLocation(nary, isDisplay, options) == DxpOmmlLimitLocation.UnderOver ? @"\limits" : @"\nolimits";
                output.Append(op).Append(placement);
                if (!nary.HideSubscript) output.Append("_{").Append(subscript).Append('}');
                if (!nary.HideSuperscript) output.Append("^{").Append(superscript).Append('}');
                output.Append(@"\,").Append(argument);
            }
            else if (format == DxpOmmlOutputFormat.UnicodeMath)
            {
                output.Append(nary.Character);
                if (!nary.HideSubscript) output.Append("_(").Append(subscript).Append(')');
                if (!nary.HideSuperscript) output.Append("^(").Append(superscript).Append(')');
                output.Append('▒').Append('〖').Append(argument).Append('〗');
            }
            else
            {
                output.Append(nary.Character);
                if (!nary.HideSubscript) output.Append(" from ").Append(subscript);
                if (!nary.HideSuperscript) output.Append(" to ").Append(superscript);
                output.Append(" of ").Append(argument);
            }
            return;
        }

        output.Append(escape(ResolveFallback((OmmlUnsupported)node, options, diagnostics)));
    }

    private static string RenderTextual(OmmlNode node, DxpOmmlOutputFormat format,
        bool isDisplay, DxpOmmlConversionOptions options, List<DxpOmmlDiagnostic> diagnostics, Func<string, string> escape)
    {
        StringBuilder result = new();
        AppendTextual(result, node, format, isDisplay, options, diagnostics, escape);
        return result.ToString();
    }

    private static void AddFence(XElement row, string value, bool grow, bool opening)
    {
        if (value.Length == 0) return;
        XNamespace math = MathMlNamespace;
        row.Add(new XElement(math + "mo", new XAttribute("fence", "true"),
            new XAttribute("stretchy", grow ? "true" : "false"),
            new XAttribute("form", opening ? "prefix" : "postfix"), value));
    }

    private static string LatexDelimiter(string value)
    {
        if (value.Length == 0) return string.Empty;
        return value switch
        {
            "{" => @"\{", "}" => @"\}", "⟨" => @"\langle", "⟩" => @"\rangle",
            "⌊" => @"\lfloor", "⌋" => @"\rfloor", "⌈" => @"\lceil", "⌉" => @"\rceil",
            "‖" => @"\Vert", "⟦" => @"\lbbrack", "⟧" => @"\rbbrack",
            _ => EscapeLatex(value),
        };
    }

    private static string MathMlAccentCharacter(string value) => value switch
    {
        "́" => "´", "̀" => "`", "̂" => "^", "̌" => "ˇ", "̃" => "~",
        "̄" => "¯", "̆" => "˘", "̇" => "˙", "̈" => "¨", "⃗" => "→",
        _ => value,
    };

    private static string LatexDecoration(OmmlDecoration decoration, string argument)
    {
        string? command = decoration.Type switch
        {
            OmmlDecorationType.Bar => decoration.Position == OmmlVerticalPosition.Top ? "overline" : "underline",
            OmmlDecorationType.GroupCharacter => (decoration.Position, decoration.Character) switch
            {
                (OmmlVerticalPosition.Top, "⏞") => "overbrace",
                (OmmlVerticalPosition.Bottom, "⏟") => "underbrace",
                (OmmlVerticalPosition.Top, "⏜") => "overparen",
                (OmmlVerticalPosition.Bottom, "⏝") => "underparen",
                _ => null,
            },
            _ => decoration.Character switch
            {
                "́" => "acute", "̀" => "grave", "̂" => "hat", "̌" => "check", "̃" => "tilde",
                "̄" => "bar", "̆" => "breve", "̇" => "dot", "̈" => "ddot", "⃗" => "vec",
                "⏞" => "overbrace", "⏜" => "overparen",
                _ => null,
            },
        };
        if (command != null) return $"\\{command}{{{argument}}}";
        string escaped = EscapeLatex(decoration.Character);
        return decoration.Position == OmmlVerticalPosition.Top
            ? $"\\overset{{\\text{{{escaped}}}}}{{{argument}}}"
            : $"\\underset{{\\text{{{escaped}}}}}{{{argument}}}";
    }

    private static string UnicodeMathDecoration(OmmlDecoration decoration, string argument)
    {
        string grouped = $"({argument})";
        if (decoration.Type == OmmlDecorationType.Accent)
            return decoration.Character is "⏞" or "⏜" ? decoration.Character + grouped : grouped + decoration.Character;
        if (decoration.Type == OmmlDecorationType.Bar)
            return grouped + (decoration.Position == OmmlVerticalPosition.Top ? "̅" : "̲");
        return decoration.Position == OmmlVerticalPosition.Top
            ? $"{grouped}┴({decoration.Character})"
            : $"{grouped}┬{decoration.Character}";
    }

    private static DxpOmmlLimitLocation NaryLimitLocation(OmmlNary nary, bool isDisplay,
        DxpOmmlConversionOptions options) => nary.LimitLocation ?? (!isDisplay
        ? DxpOmmlLimitLocation.SubscriptSuperscript
        : IsIntegral(nary.Character) ? options.IntegralLimitLocation : options.NaryLimitLocation);

    private static bool IsIntegral(string character) => character is "∫" or "∬" or "∭" or "∮" or "∯" or "∰";

    private static string LatexNaryOperator(string character) => character switch
    {
        "∫" => @"\int", "∬" => @"\iint", "∭" => @"\iiint", "∮" => @"\oint",
        "∯" => @"\oiint", "∰" => @"\oiiint", "∑" => @"\sum", "∏" => @"\prod",
        "∐" => @"\coprod", "⋂" => @"\bigcap", "⋃" => @"\bigcup",
        "⋀" => @"\bigwedge", "⋁" => @"\bigvee", "⨀" => @"\bigodot",
        "⨂" => @"\bigotimes", "⨁" => @"\bigoplus", "⨄" => @"\biguplus",
        _ => $"\\mathop{{\\text{{{EscapeLatex(character)}}}}}",
    };

    private static string? LatexFunction(string value) => value switch
    {
        "sin" or "cos" or "tan" or "cot" or "sec" or "csc" or
        "sinh" or "cosh" or "tanh" or "coth" or "log" or "ln" or "exp" or
        "arcsin" or "arccos" or "arctan" or "det" or "dim" or "gcd" or
        "hom" or "ker" or "max" or "min" or "Pr" or "sup" or "inf" or "lim" => value,
        _ => null,
    };

    private static string? LatexLimitOperator(string value) => value switch
    {
        "lim" or "liminf" or "limsup" or "max" or "min" or "sup" or "inf" or
        "det" or "dim" or "gcd" or "Pr" => value,
        _ => null,
    };

    private static string? SimpleText(OmmlSequence sequence)
    {
        if (sequence.Children.Any(node => node is not OmmlRun)) return null;
        IEnumerable<OmmlRun> runs = sequence.Children.Cast<OmmlRun>();
        if (runs.Any(run => run.Alignment || run.Script != OmmlMathScript.Default ||
                            run.Style != OmmlMathStyle.Default)) return null;
        return string.Concat(runs.SelectMany(run => run.Tokens).Select(token => token.Value));
    }

    private static string ResolveFallback(
        OmmlUnsupported unsupported,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics)
    {
        DxpOmmlDiagnostic diagnostic = new(
            "OMML001",
            DxpOmmlDiagnosticSeverity.Warning,
            $"OMML element '{unsupported.ElementName}' is not yet semantically supported; fallback policy '{options.FallbackPolicy}' was applied at {unsupported.Path}.",
            unsupported.Path,
            unsupported.ElementName);
        diagnostics.Add(diagnostic);

        return options.FallbackPolicy switch
        {
            DxpOmmlFallbackPolicy.Throw => throw new DxpOmmlUnsupportedException(diagnostic),
            DxpOmmlFallbackPolicy.ExtractText => unsupported.VisibleText,
            DxpOmmlFallbackPolicy.Placeholder => options.Placeholder ?? string.Empty,
            DxpOmmlFallbackPolicy.Omit => string.Empty,
            _ => throw new ArgumentOutOfRangeException(nameof(options.FallbackPolicy)),
        };
    }

    private static string EscapeLatex(string value)
    {
        StringBuilder escaped = new(value.Length);
        foreach (char character in value)
        {
            escaped.Append(character switch
            {
                '\\' => @"\textbackslash{}",
                '{' => @"\{",
                '}' => @"\}",
                '$' => @"\$",
                '&' => @"\&",
                '#' => @"\#",
                '%' => @"\%",
                '_' => @"\_",
                '^' => @"\^{}",
                '~' => @"\~{}",
                _ => character.ToString(),
            });
        }
        return escaped.ToString();
    }

    private static string? MathVariant(OmmlRun run)
    {
        if (run.Normal || (run.Style == OmmlMathStyle.Plain &&
            run.Script is OmmlMathScript.Default or OmmlMathScript.Roman))
            return "normal";
        string weight = run.Style switch
        {
            OmmlMathStyle.Bold => "bold",
            OmmlMathStyle.Italic => "italic",
            OmmlMathStyle.BoldItalic => "bold-italic",
            _ => "",
        };
        return run.Script switch
        {
            OmmlMathScript.Roman => weight.Length == 0 ? "normal" : weight,
            OmmlMathScript.Script => weight.StartsWith("bold", StringComparison.Ordinal) ? "bold-script" : "script",
            OmmlMathScript.Fraktur => weight.StartsWith("bold", StringComparison.Ordinal) ? "bold-fraktur" : "fraktur",
            OmmlMathScript.DoubleStruck => "double-struck",
            OmmlMathScript.SansSerif => weight switch
            {
                "bold" => "bold-sans-serif",
                "italic" => "sans-serif-italic",
                "bold-italic" => "sans-serif-bold-italic",
                _ => "sans-serif",
            },
            OmmlMathScript.Monospace => "monospace",
            _ => weight.Length == 0 ? null : weight,
        };
    }

    private static string ApplyLatexStyle(OmmlRun run, string value)
    {
        string? script = run.Script switch { OmmlMathScript.Roman => "mathrm", OmmlMathScript.Script => "mathcal", OmmlMathScript.Fraktur => "mathfrak", OmmlMathScript.DoubleStruck => "mathbb", OmmlMathScript.SansSerif => "mathsf", OmmlMathScript.Monospace => "mathtt", _ => null };
        if (script != null) value = $"\\{script}{{{value}}}";
        string? weight = run.Style switch { OmmlMathStyle.Bold => "mathbf", OmmlMathStyle.Italic => "mathit", OmmlMathStyle.BoldItalic => "boldsymbol", OmmlMathStyle.Plain when script == null => "mathrm", _ => null };
        return weight == null ? value : $"\\{weight}{{{value}}}";
    }

    private static string? UnicodeMathStyle(OmmlRun run)
    {
        return run.Script switch { OmmlMathScript.Roman => run.Style switch { OmmlMathStyle.Bold => "mbf", OmmlMathStyle.Italic => "mit", OmmlMathStyle.BoldItalic => "mbfit", _ => "mup" }, OmmlMathScript.Script => run.Style is OmmlMathStyle.Bold or OmmlMathStyle.BoldItalic ? "mbfscr" : "mscr", OmmlMathScript.Fraktur => run.Style is OmmlMathStyle.Bold or OmmlMathStyle.BoldItalic ? "mbffrak" : "mfrak", OmmlMathScript.DoubleStruck => "Bbb", OmmlMathScript.SansSerif => run.Style switch { OmmlMathStyle.Bold => "mbfsans", OmmlMathStyle.Italic => "mitsans", OmmlMathStyle.BoldItalic => "mbfitsans", _ => "msans" }, OmmlMathScript.Monospace => "mtt", _ => run.Style switch { OmmlMathStyle.Bold => "mbf", OmmlMathStyle.Italic => "mit", OmmlMathStyle.BoldItalic => "mbfit", OmmlMathStyle.Plain => "mup", _ => null } };
    }
}
