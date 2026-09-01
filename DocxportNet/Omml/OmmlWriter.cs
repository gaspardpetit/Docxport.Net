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
        bool isDisplay = options.Display ?? (document.IsDisplay || options.DisplayDefaults);
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
        XElement body;
        if (document.Children.Any(node => node is OmmlBreak))
        {
            body = new XElement(math + "mtable");
            XElement cell = AddMathMlLine(body, math);
            foreach (OmmlNode node in document.Children)
            {
                if (node is OmmlBreak)
                    cell = AddMathMlLine(body, math);
                else
                    AppendMathMl(cell, node, options.SmallFractions && !isDisplay, isDisplay, options, diagnostics);
            }
        }
        else
        {
            body = new XElement(math + "mrow");
            foreach (OmmlNode node in document.Children)
                AppendMathMl(body, node, options.SmallFractions && !isDisplay, isDisplay, options, diagnostics);
        }

        XElement root = new(
            math + "math",
            new XAttribute("display", isDisplay ? "block" : "inline"),
            body);
        ApplyMathMlLayout(root, document, options, diagnostics);
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static void AppendMathMl(
        XElement parent,
        OmmlNode node,
        bool compactFractions,
        bool isDisplay,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics,
        bool applyControlPresentation = true)
    {
        XNamespace math = MathMlNamespace;
        if (applyControlPresentation && node.ControlRevision != null && !ControlRevisionSelected(node, options))
        {
            AppendMathMl(parent, node, compactFractions, isDisplay, options, diagnostics, false);
            AddApproximation(diagnostics, node, "m:ctrlPr",
                "The rejected OMML control-character revision was omitted while retaining the surrounding mathematical structure.");
            return;
        }
        if (applyControlPresentation && (node.ControlPresentation != null || node.ControlRevision != null))
        {
            XElement content = new(math + "mrow");
            AppendMathMl(content, node, compactFractions, isDisplay, options, diagnostics, false);
            XElement style = new(math + "mstyle", content.Nodes());
            if (node.ControlPresentation != null) ApplyMathMlPresentation(style, node.ControlPresentation);
            style.SetAttributeValue("data-omml-control-properties", "true");
            if (node.ControlRevision != null && options.RevisionMode == DxpOmmlRevisionMode.Preserve)
                style.SetAttributeValue("data-omml-revision", node.ControlRevision == OmmlRevisionKind.Inserted ? "inserted" : "deleted");
            parent.Add(style);
            AddApproximation(diagnostics, node, "m:ctrlPr",
                "MathML applies OMML control-character formatting to the rendered structure because MathML has no separate selectable control character.");
            return;
        }
        if (node is OmmlSequence sequence)
        {
            XElement row = new(math + "mrow");
            foreach (OmmlNode child in sequence.Children)
                AppendMathMl(row, child, compactFractions, isDisplay, options, diagnostics);
            if (sequence.ArgumentSize is int argumentSize && argumentSize != 0)
            {
                XElement style = new(math + "mstyle", row);
                style.SetAttributeValue("scriptlevel", argumentSize > 0 ? $"-{argumentSize}" : $"+{-argumentSize}");
                parent.Add(style);
            }
            else
            {
                parent.Add(row);
            }
            return;
        }

        if (node is OmmlBreak boundary)
        {
            AppendMathMlBreak(parent, boundary.AlignmentAt);
            return;
        }

        if (node is OmmlRun run)
        {
            XElement container = parent;
            string? variant = MathVariant(run);
            if (variant != null || run.Language != null || run.RightToLeft || run.Color != null ||
                run.FontSizePoints.HasValue || run.FontFamily != null ||
                run.VerticalAlignment != OmmlRunVerticalAlignment.Baseline)
            {
                container = new XElement(math + "mstyle");
                if (variant != null) container.SetAttributeValue("mathvariant", variant);
                if (run.Language != null) container.SetAttributeValue(XNamespace.Xml + "lang", run.Language);
                if (run.RightToLeft) container.SetAttributeValue("dir", "rtl");
                if (run.Color != null) container.SetAttributeValue("mathcolor", $"#{run.Color}");
                if (run.FontSizePoints.HasValue)
                    container.SetAttributeValue("mathsize", $"{run.FontSizePoints.Value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt");
                List<string> styles = new();
                if (run.FontFamily != null)
                    styles.Add($"font-family:{CssFontFamily(run.FontFamily)}");
                if (run.VerticalAlignment != OmmlRunVerticalAlignment.Baseline)
                    styles.Add($"vertical-align:{(run.VerticalAlignment == OmmlRunVerticalAlignment.Superscript ? "super" : "sub")}");
                if (styles.Count != 0) container.SetAttributeValue("style", string.Join(";", styles));
                parent.Add(container);
            }
            if (run.Alignment && !run.BreakAlignmentAt.HasValue) container.Add(new XElement(math + "malignmark"));
            foreach (RunPiece piece in RunPieces(run, options))
            {
                if (piece.IsBreak)
                {
                    AppendMathMlBreak(container, piece.AlignmentAt);
                    if (run.Alignment && piece.AlignmentAfter)
                        container.Add(new XElement(math + "malignmark"));
                }
                else if (piece.Token is OmmlToken token)
                {
                    if (token.Value == "\u200B") container.Add(new XElement(math + "mspace", new XAttribute("width", "0")));
                    else if (token.Value == "\t") container.Add(new XElement(math + "mspace",
                        new XAttribute("width", "2em"), new XAttribute("data-omml-tab", "true")));
                    else container.Add(new XElement(math + (token.Kind switch { OmmlTokenKind.Identifier => "mi", OmmlTokenKind.Number => "mn", OmmlTokenKind.Operator => "mo", _ => "mtext" }), token.Value));
                }
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

        if (node is OmmlMatrix matrix)
        {
            XElement table = new(math + "mtable",
                new XAttribute("align", VerticalAlignment(matrix.BaseJustification)),
                new XAttribute("columnalign", string.Join(" ", MatrixAlignments(matrix).Select(HorizontalAlignment))),
                new XAttribute("data-omml-placeholder-hidden", XmlBoolean(matrix.PlaceholdersHidden)),
                new XAttribute("data-omml-row-spacing", Invariant(matrix.RowSpacing)),
                new XAttribute("data-omml-row-spacing-rule", Invariant(matrix.RowSpacingRule)),
                new XAttribute("data-omml-column-spacing", Invariant(matrix.ColumnSpacing)),
                new XAttribute("data-omml-column-gap", Invariant(matrix.ColumnGap)),
                new XAttribute("data-omml-column-gap-rule", Invariant(matrix.ColumnGapRule)));
            foreach (OmmlMatrixRow matrixRow in matrix.Rows)
            {
                XElement row = new(math + "mtr");
                foreach (OmmlSequence cell in matrixRow.Cells)
                {
                    XElement entry = new(math + "mtd");
                    if (cell.Children.Count == 0 && !matrix.PlaceholdersHidden)
                        entry.Add(new XElement(math + "mspace", new XAttribute("width", "0"), new XAttribute("data-omml-placeholder", "true")));
                    else
                        AppendMathMl(entry, cell, compactFractions, isDisplay, options, diagnostics);
                    row.Add(entry);
                }
                table.Add(row);
            }
            parent.Add(table);
            return;
        }

        if (node is OmmlEquationArray equationArray)
        {
            XElement table = new(math + "mtable",
                new XAttribute("align", VerticalAlignment(equationArray.BaseJustification)),
                new XAttribute("data-omml-max-distribution", XmlBoolean(equationArray.MaxDistribution)),
                new XAttribute("data-omml-object-distribution", XmlBoolean(equationArray.ObjectDistribution)),
                new XAttribute("data-omml-row-spacing", Invariant(equationArray.RowSpacing)),
                new XAttribute("data-omml-row-spacing-rule", Invariant(equationArray.RowSpacingRule)));
            if (equationArray.MaxDistribution) table.SetAttributeValue("width", "100%");
            foreach (OmmlSequence equationRow in equationArray.Rows)
            {
                XElement entry = new(math + "mtd");
                AppendMathMl(entry, equationRow, compactFractions, isDisplay, options, diagnostics);
                table.Add(new XElement(math + "mtr", entry));
            }
            parent.Add(table);
            return;
        }

        if (node is OmmlBox box)
        {
            string? operatorText = box.OperatorEmulator ? SimpleText(box.Argument) : null;
            string? operatorAfterBreak = operatorText;
            if (box.BreakAlignmentAt.HasValue)
            {
                if (operatorText != null && IsBreakableBinaryOperator(operatorText) &&
                    options.BreakBinary is DxpOmmlBreakBinary.After or DxpOmmlBreakBinary.Repeat)
                {
                    (string before, string after) = BreakOperators(operatorText, options);
                    if (before.Length != 0) parent.Add(BoxMathMlOperator(box, before));
                    operatorAfterBreak = after;
                }
                AppendMathMlBreak(parent, box.BreakAlignmentAt);
            }
            if (box.Alignment) parent.Add(new XElement(math + "malignmark"));
            if (box.Differential) parent.Add(new XElement(math + "mspace", new XAttribute("width", "0.1667em")));

            XElement? rendered;
            if (operatorAfterBreak != null)
                rendered = operatorAfterBreak.Length == 0 ? null : BoxMathMlOperator(box, operatorAfterBreak);
            else
            {
                rendered = new XElement(math + "mrow");
                AppendMathMl(rendered, box.Argument, compactFractions, isDisplay, options, diagnostics);
                ApplyBoxMathMlAttributes(rendered, box);
            }
            if (rendered != null) parent.Add(rendered);
            if (box.NoBreak || box.Differential || (box.OperatorEmulator && operatorText == null))
                AddApproximation(diagnostics, box, "m:box",
                    "MathML has no exact equivalent for every OMML box spacing/no-break behavior; semantic flags were retained as data attributes.");
            return;
        }

        if (node is OmmlBorderBox borderBox)
        {
            IReadOnlyList<string> notations = BorderNotations(borderBox);
            XElement rendered = notations.Count == 0
                ? new XElement(math + "mrow", new XAttribute("data-omml-border-box", "true"),
                    new XAttribute("data-omml-notation", "none"))
                : new XElement(math + "menclose", new XAttribute("notation", string.Join(" ", notations)));
            AppendMathMl(rendered, borderBox.Argument, compactFractions, isDisplay, options, diagnostics);
            parent.Add(rendered);
            return;
        }

        if (node is OmmlPhantom phantom)
        {
            XElement content = new(math + "mrow");
            AppendMathMl(content, phantom.Argument, compactFractions, isDisplay, options, diagnostics);
            XElement rendered = phantom.Show ? content : new XElement(math + "mphantom", content);
            if (phantom.ZeroWidth || phantom.ZeroAscent || phantom.ZeroDescent)
            {
                XElement padded = new(math + "mpadded", rendered);
                if (phantom.ZeroWidth) padded.SetAttributeValue("width", "0");
                if (phantom.ZeroAscent) padded.SetAttributeValue("height", "0");
                if (phantom.ZeroDescent) padded.SetAttributeValue("depth", "0");
                rendered = padded;
            }
            rendered.SetAttributeValue("data-omml-show", XmlBoolean(phantom.Show));
            rendered.SetAttributeValue("data-omml-transparent", XmlBoolean(phantom.Transparent));
            parent.Add(rendered);
            if (phantom.Transparent)
                AddApproximation(diagnostics, phantom, "m:phant",
                    "MathML cannot reproduce OMML phantom spacing-class transparency exactly; the flag was retained as metadata.");
            return;
        }

        string fallback = ResolveFallback((OmmlUnsupported)node, DxpOmmlOutputFormat.MathMl, options, diagnostics);
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
        bool multilineLatex = format == DxpOmmlOutputFormat.Latex && document.Children.Any(NeedsLatexEnvironment);
        if (multilineLatex) output.Append(@"\begin{aligned}");
        foreach (OmmlNode node in document.Children)
            AppendTextual(output, node, format, isDisplay, options, diagnostics, escape);
        if (multilineLatex) output.Append(@"\end{aligned}");
        AddTextualLayoutDiagnostics(document, format, options, diagnostics);
        return output.ToString();
    }

    private static void AppendTextual(
        StringBuilder output,
        OmmlNode node,
        DxpOmmlOutputFormat format,
        bool isDisplay,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics,
        Func<string, string> escape,
        bool applyControlPresentation = true)
    {
        if (applyControlPresentation && node.ControlRevision != null && !ControlRevisionSelected(node, options))
        {
            AppendTextual(output, node, format, isDisplay, options, diagnostics, escape, false);
            AddApproximation(diagnostics, node, "m:ctrlPr",
                "The rejected OMML control-character revision was omitted while retaining the surrounding mathematical structure.");
            return;
        }
        if (applyControlPresentation && (node.ControlPresentation != null || node.ControlRevision != null))
        {
            StringBuilder content = new();
            AppendTextual(content, node, format, isDisplay, options, diagnostics, escape, false);
            string rendered = node.ControlPresentation != null
                ? ApplyControlPresentation(node, content.ToString(), format, diagnostics)
                : content.ToString();
            if (node.ControlRevision != null && options.RevisionMode == DxpOmmlRevisionMode.Preserve)
                rendered = node.ControlRevision == OmmlRevisionKind.Inserted
                    ? $"[inserted:{rendered}]" : $"[deleted:{rendered}]";
            output.Append(rendered);
            return;
        }

        if (node is OmmlSequence sequence)
        {
            bool multilineLatex = format == DxpOmmlOutputFormat.Latex && sequence.Children.Any(NeedsLatexEnvironment);
            string? argumentSizePrefix = format == DxpOmmlOutputFormat.Latex ? sequence.ArgumentSize switch
            {
                -2 => @"{\scriptscriptstyle ",
                -1 => @"{\scriptstyle ",
                1 or 2 => @"{\displaystyle ",
                _ => null,
            } : null;
            if (argumentSizePrefix != null) output.Append(argumentSizePrefix);
            if (multilineLatex) output.Append(@"\begin{aligned}");
            foreach (OmmlNode child in sequence.Children)
                AppendTextual(output, child, format, isDisplay, options, diagnostics, escape);
            if (multilineLatex) output.Append(@"\end{aligned}");
            if (argumentSizePrefix != null) output.Append('}');
            if (sequence.ArgumentSize is int argumentSize && argumentSize != 0)
            {
                string message = format == DxpOmmlOutputFormat.Latex
                    ? "LaTeX uses the nearest standard math style for the relative OMML argument size."
                    : $"{format} preserves argument content but cannot represent its relative OMML argument size.";
                AddApproximation(diagnostics, sequence, "m:argSz", message);
            }
            return;
        }

        if (node is OmmlBreak boundary)
        {
            AppendTextualBreak(output, format, boundary.AlignmentAt);
            return;
        }

        if (node is OmmlRun run)
        {
            IReadOnlyList<RunPiece> pieces = RunPieces(run, options);
            StringBuilder renderedRun = new();
            if (format == DxpOmmlOutputFormat.Latex)
            {
                if (run.Alignment && !run.BreakAlignmentAt.HasValue) renderedRun.Append('&');
                AppendStyledLatexRun(renderedRun, run, pieces);
            }
            else
            {
                if (run.Alignment && !run.BreakAlignmentAt.HasValue && format == DxpOmmlOutputFormat.UnicodeMath) renderedRun.Append('&');
                AppendUnicodeOrTextRun(renderedRun, run, pieces, format, escape);
            }
            output.Append(ApplyRunPresentation(run, renderedRun.ToString(), format, diagnostics));
            if (run.BreakAlignmentAt > 0)
                AddApproximation(diagnostics, run, "m:brk",
                    $"{format} preserves the manual line boundary but cannot reproduce OMML's numeric operator alignment index exactly.");
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

        if (node is OmmlMatrix matrix)
        {
            string[][] rows = matrix.Rows.Select(row => row.Cells
                .Select(cell => RenderTextual(cell, format, isDisplay, options, diagnostics, escape)).ToArray()).ToArray();
            if (format == DxpOmmlOutputFormat.Latex)
            {
                string columns = string.Concat(MatrixAlignments(matrix).Select(LatexAlignment));
                output.Append("\\begin{array}[").Append(LatexVerticalAlignment(matrix.BaseJustification))
                    .Append("]{").Append(columns).Append('}');
                output.Append(string.Join(@" \\ ", rows.Select(row => string.Join(" & ", row))));
                output.Append(@"\end{array}");
            }
            else if (format == DxpOmmlOutputFormat.UnicodeMath)
                output.Append('■').Append('(').Append(string.Join("@", rows.Select(row => string.Join("&", row)))).Append(')');
            else
                output.Append('[').Append(string.Join("; ", rows.Select(row => "[" + string.Join(", ", row) + "]"))).Append(']');
            return;
        }

        if (node is OmmlEquationArray equationArray)
        {
            string[] rows = equationArray.Rows.Select(row =>
                RenderTextual(row, format, isDisplay, options, diagnostics, escape)).ToArray();
            if (format == DxpOmmlOutputFormat.Latex)
            {
                bool aligned = equationArray.Rows.Any(ContainsAlignmentMarker);
                string environment = aligned ? "aligned" : "gathered";
                output.Append("\\begin{").Append(environment).Append('}');
                if (equationArray.BaseJustification != OmmlVerticalAlignment.Center)
                    output.Append('[').Append(LatexVerticalAlignment(equationArray.BaseJustification)).Append(']');
                output
                    .Append(string.Join(@" \\ ", rows)).Append("\\end{").Append(environment).Append('}');
            }
            else if (format == DxpOmmlOutputFormat.UnicodeMath)
                output.Append('█').Append('(').Append(string.Join("@", rows)).Append(')');
            else
                output.Append(string.Join("; ", rows));
            return;
        }

        if (node is OmmlBox box)
        {
            string argument = RenderTextual(box.Argument, format, isDisplay, options, diagnostics, escape);
            string? operatorText = box.OperatorEmulator ? SimpleText(box.Argument) : null;
            string beforeBreak = string.Empty;
            string afterBreak = argument;
            if (box.BreakAlignmentAt.HasValue && operatorText != null && IsBreakableBinaryOperator(operatorText) &&
                options.BreakBinary is DxpOmmlBreakBinary.After or DxpOmmlBreakBinary.Repeat)
            {
                (string before, string after) = BreakOperators(operatorText, options);
                beforeBreak = FormatBoxOperator(before, format, box.NoBreak);
                afterBreak = FormatBoxOperator(after, format, box.NoBreak);
            }
            else if (format == DxpOmmlOutputFormat.Latex)
            {
                if (box.OperatorEmulator) afterBreak = $"\\mathop{{{afterBreak}}}";
                if (box.NoBreak) afterBreak = $"\\nobreak{{{afterBreak}}}\\nobreak";
            }
            output.Append(beforeBreak);
            if (box.BreakAlignmentAt.HasValue)
                AppendTextualBreak(output, format, box.BreakAlignmentAt);
            if (box.Alignment && format is DxpOmmlOutputFormat.Latex or DxpOmmlOutputFormat.UnicodeMath)
                output.Append('&');
            if (box.Differential && format != DxpOmmlOutputFormat.Text)
                output.Append(format == DxpOmmlOutputFormat.Latex ? @"\," : " ");
            output.Append(afterBreak);
            bool approximated = format switch
            {
                DxpOmmlOutputFormat.Latex => box.OperatorEmulator || box.NoBreak || box.Differential || box.BreakAlignmentAt > 0,
                DxpOmmlOutputFormat.UnicodeMath => box.OperatorEmulator || box.NoBreak || box.Differential || box.BreakAlignmentAt > 0,
                DxpOmmlOutputFormat.Text => box.OperatorEmulator || box.NoBreak || box.Differential || box.Alignment || box.BreakAlignmentAt > 0,
                _ => false,
            };
            if (approximated)
                AddApproximation(diagnostics, box, "m:box",
                    $"{format} cannot exactly reproduce every OMML box spacing/no-break behavior; a deterministic approximation was emitted.");
            return;
        }

        if (node is OmmlBorderBox borderBox)
        {
            string argument = RenderTextual(borderBox.Argument, format, isDisplay, options, diagnostics, escape);
            IReadOnlyList<string> notations = BorderNotations(borderBox);
            if (format == DxpOmmlOutputFormat.Latex)
            {
                if (IsPlainFourSidedBorder(borderBox)) output.Append("\\boxed{").Append(argument).Append('}');
                else
                {
                    output.Append("\\enclose{").Append(notations.Count == 0 ? "none" : string.Join(" ", notations))
                        .Append("}{").Append(argument).Append('}');
                    AddApproximation(diagnostics, borderBox, "m:borderBox",
                        "This border combination uses the MathJax/KaTeX \\enclose extension because core LaTeX has no equivalent.");
                }
            }
            else if (format == DxpOmmlOutputFormat.UnicodeMath)
            {
                if (IsPlainFourSidedBorder(borderBox)) output.Append('▭').Append('(').Append(argument).Append(')');
                else output.Append('▭').Append('(').Append(BorderMask(borderBox)).Append('&').Append(argument).Append(')');
            }
            else
            {
                output.Append("enclose[").Append(notations.Count == 0 ? "none" : string.Join(" ", notations))
                    .Append("](").Append(argument).Append(')');
                AddApproximation(diagnostics, borderBox, "m:borderBox",
                    "Readable text describes the border-box notation instead of reproducing its visual lines.");
            }
            return;
        }

        if (node is OmmlPhantom phantom)
        {
            string argument = RenderTextual(phantom.Argument, format, isDisplay, options, diagnostics, escape);
            if (format == DxpOmmlOutputFormat.Latex)
                output.Append(LatexPhantom(phantom, argument));
            else if (format == DxpOmmlOutputFormat.UnicodeMath)
                output.Append(UnicodeMathPhantom(phantom, argument));
            else if (phantom.Show)
                output.Append(argument);

            bool approximated = format switch
            {
                DxpOmmlOutputFormat.Latex => phantom.Transparent || (phantom.Show && phantom.ZeroWidth),
                DxpOmmlOutputFormat.UnicodeMath => phantom.Transparent,
                DxpOmmlOutputFormat.Text => !phantom.Show || phantom.ZeroWidth || phantom.ZeroAscent || phantom.ZeroDescent || phantom.Transparent,
                _ => false,
            };
            if (approximated)
                AddApproximation(diagnostics, phantom, "m:phant",
                    $"{format} cannot exactly reproduce every OMML phantom layout property; the documented deterministic policy was applied.");
            return;
        }

        output.Append(escape(ResolveFallback((OmmlUnsupported)node, format, options, diagnostics)));
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

    private static IReadOnlyList<OmmlHorizontalAlignment> MatrixAlignments(OmmlMatrix matrix)
    {
        int actualColumns = matrix.Rows.Count == 0 ? 0 : matrix.Rows.Max(row => row.Cells.Count);
        List<OmmlHorizontalAlignment> result = matrix.Columns
            .SelectMany(column => Enumerable.Repeat(column.Alignment, column.Count)).ToList();
        while (result.Count < actualColumns) result.Add(OmmlHorizontalAlignment.Center);
        if (result.Count == 0) result.Add(OmmlHorizontalAlignment.Center);
        return result;
    }

    private static void AppendMathMlBreak(XElement parent, int? alignmentAt)
    {
        XNamespace math = MathMlNamespace;
        XElement lineBreak = new(math + "mspace", new XAttribute("linebreak", "newline"));
        if (alignmentAt.HasValue)
            lineBreak.SetAttributeValue("data-omml-align-at", Invariant(alignmentAt.Value));
        parent.Add(lineBreak);
    }

    private static XElement AddMathMlLine(XElement table, XNamespace math)
    {
        XElement cell = new(math + "mtd");
        table.Add(new XElement(math + "mtr", cell));
        return cell;
    }

    private static void AppendTextualBreak(StringBuilder output, DxpOmmlOutputFormat format, int? alignmentAt)
    {
        output.Append(format == DxpOmmlOutputFormat.Latex ? @"\\" : "\n");
    }

    private static void AppendStyledLatexRun(StringBuilder output, OmmlRun run, IReadOnlyList<RunPiece> pieces)
    {
        IReadOnlyList<IReadOnlyList<OmmlToken>> segments = SplitRunAtBreaks(pieces);
        for (int i = 0; i < segments.Count; i++)
        {
            if (i != 0) output.Append(@"\\");
            if (i == 1 && run.Alignment && run.BreakAlignmentAt.HasValue) output.Append('&');
            string value = string.Concat(segments[i].Select(token => EscapeLatex(token.Value)))
                .Replace("\u200B", "{}")
                .Replace("\t", "\\quad ");
            output.Append(ApplyLatexStyle(run, value));
        }
    }

    private static void AppendUnicodeOrTextRun(StringBuilder output, OmmlRun run, IReadOnlyList<RunPiece> pieces,
        DxpOmmlOutputFormat format, Func<string, string> escape)
    {
        IReadOnlyList<IReadOnlyList<OmmlToken>> segments = SplitRunAtBreaks(pieces);
        for (int i = 0; i < segments.Count; i++)
        {
            if (i != 0) output.Append('\n');
            if (i == 1 && run.Alignment && run.BreakAlignmentAt.HasValue &&
                format == DxpOmmlOutputFormat.UnicodeMath) output.Append('&');
            string value = string.Concat(segments[i].Select(token => token.Value));
            if (format == DxpOmmlOutputFormat.UnicodeMath && UnicodeMathStyle(run) is string command)
                output.Append($"\\{command}\"{value.Replace("\"", "\\\"")}\"");
            else output.Append(escape(format == DxpOmmlOutputFormat.Text
                ? value.Replace("\u200B", string.Empty) : value));
        }
    }

    private static IReadOnlyList<IReadOnlyList<OmmlToken>> SplitRunAtBreaks(IReadOnlyList<RunPiece> pieces)
    {
        List<IReadOnlyList<OmmlToken>> result = new();
        List<OmmlToken> segment = new();
        foreach (RunPiece piece in pieces)
        {
            if (piece.IsBreak)
            {
                result.Add(segment.ToArray());
                segment.Clear();
            }
            else if (piece.Token != null) segment.Add(piece.Token);
        }
        result.Add(segment.ToArray());
        return result;
    }

    private static IReadOnlyList<RunPiece> RunPieces(OmmlRun run, DxpOmmlConversionOptions options)
    {
        List<RunPiece> source = new();
        if (run.BreakAlignmentAt.HasValue)
            source.Add(RunPiece.Break(run.BreakAlignmentAt, alignmentAfter: true));
        foreach (OmmlToken token in run.Tokens)
            source.Add(token.Kind == OmmlTokenKind.LineBreak ? RunPiece.Break(null) : RunPiece.Content(token));

        if (options.BreakBinary is null or DxpOmmlBreakBinary.Before) return source;
        List<RunPiece> result = new();
        for (int i = 0; i < source.Count; i++)
        {
            RunPiece piece = source[i];
            if (!piece.IsBreak || i + 1 >= source.Count || source[i + 1].Token is not OmmlToken token ||
                !IsBreakableBinaryOperator(token.Value))
            {
                result.Add(piece);
                continue;
            }

            (string before, string after) = BreakOperators(token.Value, options);
            if (before.Length != 0) result.Add(RunPiece.Content(new OmmlToken(OmmlTokenKind.Operator, before)));
            result.Add(piece);
            if (after.Length != 0) result.Add(RunPiece.Content(new OmmlToken(OmmlTokenKind.Operator, after)));
            i++;
        }
        return result;
    }

    private static bool IsBreakableBinaryOperator(string value) =>
        value is "+" or "-" or "−" or "=" or "==" or "≠" or "<" or ">" or "≤" or "≥" or
            "×" or "÷" or "/" or "*" or "±" or "∓";

    private static (string Before, string After) BreakOperators(string value, DxpOmmlConversionOptions options)
    {
        if (options.BreakBinary == DxpOmmlBreakBinary.After) return (value, string.Empty);
        if (value is not ("-" or "−")) return (value, value);
        string minus = value;
        return options.BreakBinarySubtraction switch
        {
            DxpOmmlBreakBinarySubtraction.MinusPlus => (minus, "+"),
            DxpOmmlBreakBinarySubtraction.PlusMinus => ("+", minus),
            _ => (minus, minus),
        };
    }

    private sealed class RunPiece
    {
        private RunPiece(OmmlToken? token, bool isBreak, int? alignmentAt, bool alignmentAfter)
        { Token = token; IsBreak = isBreak; AlignmentAt = alignmentAt; AlignmentAfter = alignmentAfter; }
        public OmmlToken? Token { get; }
        public bool IsBreak { get; }
        public int? AlignmentAt { get; }
        public bool AlignmentAfter { get; }
        public static RunPiece Content(OmmlToken token) => new(token, false, null, false);
        public static RunPiece Break(int? alignmentAt, bool alignmentAfter = false) =>
            new(null, true, alignmentAt, alignmentAfter);
    }

    private static bool NeedsLatexEnvironment(OmmlNode node) => node switch
    {
        OmmlBreak => true,
        OmmlRun run => run.BreakAlignmentAt.HasValue || run.Tokens.Any(token => token.Kind == OmmlTokenKind.LineBreak),
        OmmlBox box => box.BreakAlignmentAt.HasValue,
        _ => false,
    };

    private static XElement BoxMathMlOperator(OmmlBox box, string value)
    {
        XNamespace math = MathMlNamespace;
        XElement result = new(math + "mo", new XAttribute("form", "infix"), value);
        ApplyBoxMathMlAttributes(result, box);
        return result;
    }

    private static void ApplyBoxMathMlAttributes(XElement element, OmmlBox box)
    {
        element.SetAttributeValue("data-omml-operator-emulator", XmlBoolean(box.OperatorEmulator));
        element.SetAttributeValue("data-omml-no-break", XmlBoolean(box.NoBreak));
        element.SetAttributeValue("data-omml-differential", XmlBoolean(box.Differential));
    }

    private static string FormatBoxOperator(string value, DxpOmmlOutputFormat format, bool noBreak)
    {
        if (value.Length == 0) return string.Empty;
        if (format != DxpOmmlOutputFormat.Latex) return value;
        string result = $"\\mathop{{{EscapeLatex(value)}}}";
        return noBreak ? $"\\nobreak{{{result}}}\\nobreak" : result;
    }

    private static void ApplyMathMlLayout(XElement root, OmmlDocument document,
        DxpOmmlConversionOptions options, List<DxpOmmlDiagnostic> diagnostics)
    {
        DxpOmmlJustification? justification = document.Justification ?? options.DefaultJustification;
        if (justification.HasValue)
        {
            root.SetAttributeValue("data-omml-justification", Justification(justification.Value));
            root.SetAttributeValue("style", $"text-align: {CssJustification(justification.Value)}");
            XElement? table = root.Element(XName.Get("mtable", MathMlNamespace));
            if (table != null) table.SetAttributeValue("columnalign", CssJustification(justification.Value));
            if (justification == DxpOmmlJustification.CenterGroup)
                AddJustificationApproximation(document, diagnostics,
                    "MathML centers the paragraph but cannot distinguish OMML centerGroup from center exactly.");
        }
        SetData(root, "math-font", options.MathFont);
        SetData(root, "break-binary", options.BreakBinary?.ToString().ToLowerInvariant());
        SetData(root, "break-binary-subtraction", BreakBinarySubtraction(options.BreakBinarySubtraction));
        SetData(root, "left-margin-twips", options.LeftMarginTwips);
        SetData(root, "right-margin-twips", options.RightMarginTwips);
        SetData(root, "pre-spacing-twips", options.PreSpacingTwips);
        SetData(root, "post-spacing-twips", options.PostSpacingTwips);
        SetData(root, "inter-spacing-twips", options.InterSpacingTwips);
        SetData(root, "intra-spacing-twips", options.IntraSpacingTwips);
        if (!options.WrapRight) SetData(root, "wrap-indent-twips", options.WrapIndentTwips);
        if (options.WrapRight) root.SetAttributeValue("data-omml-wrap-right", "true");
        if (HasApproximateDocumentSettings(options))
            AddDocumentApproximation(diagnostics,
                "OMML document math font, break, spacing, and wrapping settings were retained as MathML metadata where no exact portable equivalent exists.");
    }

    private static void AddTextualLayoutDiagnostics(OmmlDocument document, DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions options, List<DxpOmmlDiagnostic> diagnostics)
    {
        if (document.Justification.HasValue || options.DefaultJustification.HasValue)
            AddJustificationApproximation(document, diagnostics,
                $"{format} output does not control the surrounding paragraph's OMML justification.");
        if (HasApproximateDocumentSettings(options))
            AddDocumentApproximation(diagnostics,
                $"{format} output cannot portably reproduce OMML document math font, break, spacing, margin, and wrapping settings.");
    }

    private static bool HasApproximateDocumentSettings(DxpOmmlConversionOptions options) =>
        !string.IsNullOrEmpty(options.MathFont) || options.BreakBinary.HasValue ||
        options.BreakBinarySubtraction.HasValue || options.LeftMarginTwips.HasValue ||
        options.RightMarginTwips.HasValue || options.PreSpacingTwips.HasValue ||
        options.PostSpacingTwips.HasValue || options.InterSpacingTwips.HasValue ||
        options.IntraSpacingTwips.HasValue || options.WrapIndentTwips.HasValue || options.WrapRight;

    private static void AddDocumentApproximation(List<DxpOmmlDiagnostic> diagnostics, string message) =>
        diagnostics.Add(new DxpOmmlDiagnostic("OMML002", DxpOmmlDiagnosticSeverity.Warning,
            message, "/m:mathPr[1]", "m:mathPr"));

    private static void AddJustificationApproximation(OmmlDocument document,
        List<DxpOmmlDiagnostic> diagnostics, string message) => diagnostics.Add(new DxpOmmlDiagnostic(
            "OMML002", DxpOmmlDiagnosticSeverity.Warning, message,
            document.Justification.HasValue ? "/m:oMathPara[1]/m:oMathParaPr[1]/m:jc[1]" : "/m:mathPr[1]/m:defJc[1]",
            document.Justification.HasValue ? "m:jc" : "m:defJc"));

    private static void SetData(XElement root, string name, string? value)
    { if (!string.IsNullOrEmpty(value)) root.SetAttributeValue("data-omml-" + name, value); }
    private static void SetData(XElement root, string name, uint? value)
    { if (value.HasValue) root.SetAttributeValue("data-omml-" + name, Invariant(value.Value)); }
    private static string Justification(DxpOmmlJustification value) => value switch
    { DxpOmmlJustification.Left => "left", DxpOmmlJustification.Right => "right", DxpOmmlJustification.Center => "center", _ => "centerGroup" };
    private static string CssJustification(DxpOmmlJustification value) => value switch
    { DxpOmmlJustification.Left => "left", DxpOmmlJustification.Right => "right", _ => "center" };
    private static string? BreakBinarySubtraction(DxpOmmlBreakBinarySubtraction? value) => value switch
    { DxpOmmlBreakBinarySubtraction.MinusMinus => "--", DxpOmmlBreakBinarySubtraction.MinusPlus => "-+", DxpOmmlBreakBinarySubtraction.PlusMinus => "+-", _ => null };

    private static bool ContainsAlignmentMarker(OmmlNode node) => node switch
    {
        OmmlRun run => run.Alignment,
        OmmlSequence sequence => sequence.Children.Any(ContainsAlignmentMarker),
        OmmlFraction fraction => ContainsAlignmentMarker(fraction.Numerator) || ContainsAlignmentMarker(fraction.Denominator),
        OmmlRadical radical => ContainsAlignmentMarker(radical.Radicand) || ContainsAlignmentMarker(radical.Degree),
        OmmlScript script => ContainsAlignmentMarker(script.Base) || ContainsAlignmentMarker(script.Subscript) || ContainsAlignmentMarker(script.Superscript),
        OmmlDelimiter delimiter => delimiter.Arguments.Any(ContainsAlignmentMarker),
        OmmlDecoration decoration => ContainsAlignmentMarker(decoration.Argument),
        OmmlFunction function => ContainsAlignmentMarker(function.Name) || ContainsAlignmentMarker(function.Argument),
        OmmlLimit limit => ContainsAlignmentMarker(limit.Base) || ContainsAlignmentMarker(limit.Limit),
        OmmlNary nary => ContainsAlignmentMarker(nary.Subscript) || ContainsAlignmentMarker(nary.Superscript) || ContainsAlignmentMarker(nary.Argument),
        OmmlMatrix matrix => matrix.Rows.Any(row => row.Cells.Any(ContainsAlignmentMarker)),
        OmmlEquationArray equationArray => equationArray.Rows.Any(ContainsAlignmentMarker),
        OmmlBox box => box.Alignment || ContainsAlignmentMarker(box.Argument),
        OmmlBorderBox borderBox => ContainsAlignmentMarker(borderBox.Argument),
        OmmlPhantom phantom => ContainsAlignmentMarker(phantom.Argument),
        _ => false,
    };

    private static IReadOnlyList<string> BorderNotations(OmmlBorderBox borderBox)
    {
        List<string> result = new();
        if (!borderBox.HideTop) result.Add("top");
        if (!borderBox.HideBottom) result.Add("bottom");
        if (!borderBox.HideLeft) result.Add("left");
        if (!borderBox.HideRight) result.Add("right");
        if (borderBox.StrikeHorizontal) result.Add("horizontalstrike");
        if (borderBox.StrikeVertical) result.Add("verticalstrike");
        if (borderBox.StrikeBottomLeftToTopRight) result.Add("updiagonalstrike");
        if (borderBox.StrikeTopLeftToBottomRight) result.Add("downdiagonalstrike");
        return result;
    }

    private static bool IsPlainFourSidedBorder(OmmlBorderBox value) =>
        !value.HideTop && !value.HideBottom && !value.HideLeft && !value.HideRight &&
        !value.StrikeHorizontal && !value.StrikeVertical &&
        !value.StrikeBottomLeftToTopRight && !value.StrikeTopLeftToBottomRight;

    private static int BorderMask(OmmlBorderBox value)
    {
        int mask = 0;
        if (!value.HideTop) mask |= 1;
        if (!value.HideBottom) mask |= 2;
        if (!value.HideLeft) mask |= 4;
        if (!value.HideRight) mask |= 8;
        if (value.StrikeHorizontal) mask |= 16;
        if (value.StrikeVertical) mask |= 32;
        if (value.StrikeTopLeftToBottomRight) mask |= 64;
        if (value.StrikeBottomLeftToTopRight) mask |= 128;
        return mask;
    }

    private static string LatexPhantom(OmmlPhantom phantom, string argument)
    {
        string vertical = SmashVerticalLatex(phantom, argument);
        if (!phantom.Show)
        {
            if (phantom.ZeroWidth) return $"\\vphantom{{{vertical}}}";
            if (phantom.ZeroAscent && phantom.ZeroDescent) return $"\\hphantom{{{argument}}}";
            return $"\\phantom{{{vertical}}}";
        }
        return phantom.ZeroWidth ? $"\\mathrlap{{{vertical}}}" : vertical;
    }

    private static string SmashVerticalLatex(OmmlPhantom phantom, string argument)
    {
        if (phantom.ZeroAscent && phantom.ZeroDescent) return $"\\smash{{{argument}}}";
        if (phantom.ZeroAscent) return $"\\smash[t]{{{argument}}}";
        if (phantom.ZeroDescent) return $"\\smash[b]{{{argument}}}";
        return argument;
    }

    private static string UnicodeMathPhantom(OmmlPhantom phantom, string argument)
    {
        if (!phantom.Show && !phantom.ZeroWidth && !phantom.ZeroAscent && !phantom.ZeroDescent)
            return $"⟡({argument})";
        if (!phantom.Show && phantom.ZeroWidth && !phantom.ZeroAscent && !phantom.ZeroDescent)
            return $"⇳({argument})";
        if (!phantom.Show && !phantom.ZeroWidth && phantom.ZeroAscent && phantom.ZeroDescent)
            return $"⬄({argument})";

        string result = phantom.Show ? argument : $"⟡({argument})";
        if (phantom.ZeroAscent && phantom.ZeroDescent) result = $"⬍({result})";
        else if (phantom.ZeroAscent) result = $"⬆({result})";
        else if (phantom.ZeroDescent) result = $"⬇({result})";
        if (phantom.ZeroWidth) result = $"⬌({result})";
        return result;
    }

    private static string HorizontalAlignment(OmmlHorizontalAlignment alignment) => alignment switch
    { OmmlHorizontalAlignment.Left => "left", OmmlHorizontalAlignment.Right => "right", _ => "center" };
    private static char LatexAlignment(OmmlHorizontalAlignment alignment) => alignment switch
    { OmmlHorizontalAlignment.Left => 'l', OmmlHorizontalAlignment.Right => 'r', _ => 'c' };
    private static string VerticalAlignment(OmmlVerticalAlignment alignment) => alignment switch
    { OmmlVerticalAlignment.Top => "top", OmmlVerticalAlignment.Bottom => "bottom", _ => "center" };
    private static char LatexVerticalAlignment(OmmlVerticalAlignment alignment) => alignment switch
    { OmmlVerticalAlignment.Top => 't', OmmlVerticalAlignment.Bottom => 'b', _ => 'c' };
    private static string XmlBoolean(bool value) => value ? "true" : "false";
    private static string Invariant(uint value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
    private static string Invariant(int value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static void AddApproximation(List<DxpOmmlDiagnostic> diagnostics,
        OmmlNode node, string elementName, string message) => diagnostics.Add(new DxpOmmlDiagnostic(
            "OMML002", DxpOmmlDiagnosticSeverity.Warning, message, node.Path, elementName));

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
        DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics)
    {
        if (format == DxpOmmlOutputFormat.Latex && options.EmbeddedContentResolver != null)
        {
            string? resolved = options.EmbeddedContentResolver.Resolve(new DxpOmmlEmbeddedContentRequest(
                unsupported.XmlElements, unsupported.OpenXmlElements,
                unsupported.Path, unsupported.ElementName, format,
                options.RevisionMode, options.FieldMode, options.IncludeHyperlinkTargets));
            if (resolved != null)
                return resolved;
        }

        if (OmmlEmbeddedWordprocessing.TryResolve(unsupported, options, diagnostics, out string embedded))
            return embedded;

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

    private static string ApplyRunPresentation(OmmlRun run, string value, DxpOmmlOutputFormat format,
        List<DxpOmmlDiagnostic> diagnostics)
    {
        if (format == DxpOmmlOutputFormat.Latex)
        {
            if (run.Color != null) value = $"\\textcolor[HTML]{{{run.Color}}}{{{value}}}";
            if (run.FontSizePoints.HasValue)
            {
                string size = run.FontSizePoints.Value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                string leading = (run.FontSizePoints.Value * 1.2).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                value = $"{{\\fontsize{{{size}pt}}{{{leading}pt}}\\selectfont {value}}}";
            }
            if (run.VerticalAlignment == OmmlRunVerticalAlignment.Superscript)
                value = $"{{}}^{{\\scriptstyle {value}}}";
            else if (run.VerticalAlignment == OmmlRunVerticalAlignment.Subscript)
                value = $"{{}}_{{\\scriptstyle {value}}}";
        }
        else if (format == DxpOmmlOutputFormat.UnicodeMath)
        {
            if (run.VerticalAlignment == OmmlRunVerticalAlignment.Superscript) value = $"^({value})";
            else if (run.VerticalAlignment == OmmlRunVerticalAlignment.Subscript) value = $"_({value})";
        }

        if (format != DxpOmmlOutputFormat.MathMl && run.FontFamily != null)
            AddApproximation(diagnostics, run, "w:rFonts", $"{format} cannot select arbitrary Word font '{run.FontFamily}' portably.");
        if (format is DxpOmmlOutputFormat.UnicodeMath or DxpOmmlOutputFormat.Text &&
            (run.Color != null || run.FontSizePoints.HasValue))
            AddApproximation(diagnostics, run, "w:rPr", $"{format} preserves visible content but cannot encode Word color or font size.");
        return value;
    }

    private static void ApplyMathMlPresentation(XElement style, OmmlRunPresentation presentation)
    {
        if (presentation.Bold) style.SetAttributeValue("data-omml-control-bold", "true");
        if (presentation.Italic) style.SetAttributeValue("data-omml-control-italic", "true");
        if (presentation.Color != null) style.SetAttributeValue("data-omml-control-color", presentation.Color);
        if (presentation.FontSizePoints.HasValue)
            style.SetAttributeValue("data-omml-control-size-pt",
                presentation.FontSizePoints.Value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture));
        if (presentation.FontFamily != null) style.SetAttributeValue("data-omml-control-font", presentation.FontFamily);
        if (presentation.Language != null) style.SetAttributeValue("data-omml-control-language", presentation.Language);
        if (presentation.RightToLeft) style.SetAttributeValue("data-omml-control-rtl", "true");
        if (presentation.VerticalAlignment != OmmlRunVerticalAlignment.Baseline)
            style.SetAttributeValue("data-omml-control-vertical-align",
                presentation.VerticalAlignment == OmmlRunVerticalAlignment.Superscript ? "superscript" : "subscript");
    }

    private static string ApplyControlPresentation(OmmlNode node, string value, DxpOmmlOutputFormat format,
        List<DxpOmmlDiagnostic> diagnostics)
    {
        AddApproximation(diagnostics, node, "m:ctrlPr",
            $"{format} preserves the mathematical structure but cannot apply Word formatting solely to its non-selectable control character.");
        return value;
    }

    private static bool ControlRevisionSelected(OmmlNode node, DxpOmmlConversionOptions options) =>
        options.RevisionMode == DxpOmmlRevisionMode.Preserve ||
        (options.RevisionMode == DxpOmmlRevisionMode.Accept && node.ControlRevision == OmmlRevisionKind.Inserted) ||
        (options.RevisionMode == DxpOmmlRevisionMode.Reject && node.ControlRevision == OmmlRevisionKind.Deleted);

    private static string CssFontFamily(string value) =>
        $"'{value.Replace("'", "\\'")}'";

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
