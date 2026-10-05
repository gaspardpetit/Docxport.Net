using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes property modifiers without applying formatting to text.</summary>
internal static class DocPropertyNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode parent)
    {
        if (parent.Children.Count != 0 || parent.StreamName == null ||
            parent.Offset == null || parent.Length == null) return;
        var bytes = structure.ReadRange(parent.StreamName, parent.Offset.Value,
            checked((int)parent.Length.Value));
        var start = 0;
        var end = bytes.Length;
        var output = parent;
        switch (parent.Kind)
        {
            case "Chpx":
                start = 1;
                break;
            case "PapxInFkp":
                start = bytes[0] == 0 ? 2 : 1;
                if (end - start < 2) return;
                var paragraph = new DocStructureNode("GrpPrlAndIstd", "ParagraphProperties",
                    parent.StreamName, parent.Offset + start, end - start);
                paragraph.Attributes["styleIndex"] = BinaryPrimitives.ReadUInt16LittleEndian(
                    bytes.AsSpan(start, 2)).ToString(CultureInfo.InvariantCulture);
                parent.Children.Add(paragraph);
                output = paragraph;
                start += 2;
                break;
            case "PrcData":
            case "Sepx":
                start = 2;
                break;
            case "UpxPapx":
                if (end < 2) return;
                var style = new DocStructureNode("GrpPrlAndIstd", "StyleParagraphProperties",
                    parent.StreamName, parent.Offset, end);
                style.Attributes["styleIndex"] = BinaryPrimitives.ReadUInt16LittleEndian(bytes)
                    .ToString(CultureInfo.InvariantCulture);
                parent.Children.Add(style);
                output = style;
                start = 2;
                break;
            case "UpxChpx":
            case "UpxTapx":
                break;
            default:
                return;
        }

        for (var cursor = start; cursor < end;)
        {
            if (end - cursor < 3) break;
            var sprm = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor, 2));
            var spra = sprm >> 13;
            var operandStart = cursor + 2;
            var operandSize = spra switch
            {
                0 or 1 => 1,
                2 or 4 or 5 => 2,
                3 => 4,
                7 => 3,
                6 when sprm == 0xD608 && end - operandStart >= 2 =>
                    2 + BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(operandStart, 2)) - 1,
                6 when bytes[operandStart] != 255 => 1 + bytes[operandStart],
                _ => -1
            };
            if (operandSize < 0 || operandSize > end - operandStart) break;
            var prl = new DocStructureNode("Prl", $"Modifier{output.Children.Count}",
                parent.StreamName, parent.Offset + cursor, 2L + operandSize);
            var modifier = new DocStructureNode("Sprm", "ModifierCode", parent.StreamName,
                parent.Offset + cursor, 2);
            modifier.Attributes["code"] = $"0x{sprm:X4}";
            modifier.Attributes["group"] = ((sprm >> 10) & 7).ToString(CultureInfo.InvariantCulture);
            modifier.Attributes["operandSizeCode"] = spra.ToString(CultureInfo.InvariantCulture);
            prl.Children.Add(modifier);
            var operandKind = OperandKind(sprm, spra);
            var operand = new DocStructureNode(operandKind, "Operand", parent.StreamName,
                parent.Offset + operandStart, operandSize);
            if (operandKind == "ToggleOperand")
                operand.Attributes["value"] = bytes[operandStart].ToString(CultureInfo.InvariantCulture);
            if (sprm == 0x6A03 && operandSize == 4)
                operand.Attributes["dataOffset"] = BinaryPrimitives.ReadUInt32LittleEndian(
                    bytes.AsSpan(operandStart, 4)).ToString(CultureInfo.InvariantCulture);
            if (sprm == 0x646B && operandSize == 4)
            {
                var dataOffset = BinaryPrimitives.ReadUInt32LittleEndian(
                    bytes.AsSpan(operandStart, 4));
                var chain = parent.Attributes.TryGetValue("tablePropsChain", out var previous)
                    ? previous.Split(',') : Array.Empty<string>();
                var key = dataOffset.ToString(CultureInfo.InvariantCulture);
                if (chain.Length >= 32 || chain.Contains(key, StringComparer.Ordinal))
                    throw new InvalidDataException("A DOC table-property reference is cyclic or too deep.");
                var size = BinaryPrimitives.ReadUInt16LittleEndian(
                    structure.ReadRange("Data", dataOffset, 2));
                var tableProperties = new DocStructureNode("PrcData", "TableProperties",
                    "Data", dataOffset, checked(2 + size));
                tableProperties.Attributes["tablePropsChain"] =
                    previous == null ? key : previous + "," + key;
                operand.Children.Add(tableProperties);
            }
            prl.Children.Add(operand);
            output.Children.Add(prl);
            cursor = operandStart + operandSize;
        }
    }

    private static string OperandKind(ushort sprm, int spra) => sprm switch
    {
        0xD608 => "TDefTableOperand",
        0x6A09 => "CSymbolOperand",
        0xCA47 => "CMajorityOperand",
        0xCA76 => "CFitTextOperand",
        0xCA78 => "FarEastLayoutOperand",
        0x2A86 => "FFM",
        0x2A3E => "Kul",
        0xCA62 => "DispFldRmOperand",
        0xC81A => "MathPrOperand",
        0xCA31 => "SPPOperand",
        0xC601 => "SPPOperand",
        0xC60D => "PChgTabsPapxOperand",
        0x6412 => "LSPD",
        0x442B => "WHeightAbs",
        0x442C => "DCS",
        0x442D => "Shd80",
        0x4866 => "Shd80",
        0x443A => "FrameTextFlowOperand",
        0xC645 => "NumRMOperand",
        0xC66C => "PTIstdInfoOperand",
        0x4888 => "PbiGrfOperand",
        0xD662 => "TCellBrcTypeOperand",
        0x484E => "HresiOperand",
        0x2879 => "LBCOperand",
        0xCA57 => "PropRMarkOperand",
        0xC615 => "PChgTabsOperand",
        0xCA71 => "SHDOperand",
        0xCA72 => "BrcOperand",
        0x2A0C => "Ico",
        0x6870 or 0x6877 => "COLORREF",
        0x6805 or 0x6864 => "DTTM",
        0x0806 or 0x080A or 0x0856 => "Bool8",
        0x9601 or 0x9602 => "XAS",
        0x9407 => "YAS",
        0x940E => "XAS_plusOne",
        0x940F => "YAS_plusOne",
        0x9410 or 0x941E => "XAS_nonNeg",
        0x9411 or 0x941F => "YAS_nonNeg",
        0x3403 or 0x3404 or 0x3615 or 0x3619 or 0x347D or
        0x3005 or 0x3006 or 0x300A or 0x3011 or 0x3012 or 0x3019 or
        0x3228 or 0x322A or 0x3239 => "Bool8",
        0x560B => "Bool16",
        0xD605 => "TableBordersOperand80",
        0xD613 => "TableBordersOperand",
        0xD609 => "DefTableShd80Operand",
        0xD60C or 0xD612 or 0xD616 or 0xD670 or 0xD671 or 0xD672 => "DefTableShdOperand",
        0x740A => "TLP",
        0x360D => "PositionCodeOperand",
        0xF614 => "FtsWWidth_Table",
        0xF617 or 0xF618 => "FtsWWidth_TablePart",
        0xF661 => "FtsWWidth_Indent",
        0xD61A or 0xD61B or 0xD61C or 0xD61D => "BrcCvOperand",
        0xD620 => "TableBrc80Operand",
        0xD62F => "TableBrcOperand",
        0x7621 => "TInsertOperand",
        0x5622 or 0x5624 or 0x5625 => "ItcFirstLim",
        0x7623 => "TDxaColOperand",
        0x7629 => "CellRangeTextFlow",
        0xD62B => "VertMergeOperand",
        0xD62C => "CellRangeVertAlign",
        0xD62D or 0xD62E => "TableShadeOperand",
        0xD632 or 0xD633 or 0xD634 or 0xD63E => "CSSAOperand",
        0xD635 => "TableCellWidthOperand",
        0xF636 => "CellRangeFitText",
        0xD639 => "CellRangeNoWrap",
        0xD642 => "CellHideMarkOperand",
        0xD660 => "SHDOperand",
        0xD66A => "CNFOperand",
        0x347C => "VerticalAlign",
        0x3000 => "CNS",
        0xF203 => "SDxaColWidthOperand",
        0xF204 => "SDxaColSpacingOperand",
        0x5007 or 0x5008 => "SDmBinOperand",
        0x3009 => "SBkcOperand",
        0x900C or 0x9016 or 0xB021 or 0xB022 => "XAS_nonNeg",
        0x3013 => "SLncOperand",
        0xB017 or 0xB018 => "YAS_nonNeg",
        0x301A => "Vjc",
        0x301D => "SBOrientationOperand",
        0x9023 or 0x9024 or 0x9031 => "YAS",
        0x702B or 0x702C or 0x702D or 0x702E => "Brc80",
        0x522F => "SPgbPropOperand",
        0x5032 => "SClmOperand",
        0xD234 or 0xD235 or 0xD236 or 0xD237 => "BrcOperand",
        0x303B => "SFpcOperand",
        0x303C or 0x303E => "Rnc",
        0xD243 => "PropRMarkOperand",
        _ when spra == 0 && ((sprm >> 10) & 7) == 2 => "ToggleOperand",
        _ => "Operand"
    };
}
