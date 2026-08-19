using System;
using System.Globalization;
using System.Text;

namespace ExcelPowerTools.Core.Comparison.Adapters.OpenXml
{
    /// <summary>
    /// Conservative A1-reference translation for a shared-formula follower. This is deliberately not a general
    /// Excel formula parser: unsupported bracketed references fail closed so callers can emit coverage instead of
    /// manufacturing formula text.
    /// </summary>
    internal static class OpenXmlSharedFormulaTranslator
    {
        private const int MaxRows = 1048576;
        private const int MaxColumns = 16384;

        internal static bool TryTranslate(
            string masterFormula,
            string masterCellReference,
            string followerCellReference,
            out string translatedFormula,
            out string failureReason)
        {
            translatedFormula = null;
            failureReason = null;

            if (string.IsNullOrEmpty(masterFormula))
            {
                failureReason = "The shared-formula master has no stored formula text.";
                return false;
            }

            if (!TryParseCellReference(masterCellReference, out CellCoordinate master) ||
                !TryParseCellReference(followerCellReference, out CellCoordinate follower))
            {
                failureReason = "The shared-formula master or follower cell reference is invalid.";
                return false;
            }

            int rowDelta = follower.Row - master.Row;
            int columnDelta = follower.Column - master.Column;
            if (!TryValidateQualifiers(masterFormula, out failureReason)) return false;
            var result = new StringBuilder(masterFormula.Length + 16);

            int index = 0;
            while (index < masterFormula.Length)
            {
                char current = masterFormula[index];
                if (current == '"')
                {
                    if (!CopyDoubleQuotedString(masterFormula, ref index, result))
                    {
                        failureReason = "The shared formula contains an unterminated string literal.";
                        return false;
                    }
                    continue;
                }

                if (current == '\'')
                {
                    if (!CopySingleQuotedIdentifier(masterFormula, ref index, result))
                    {
                        failureReason = "The shared formula contains an unterminated quoted sheet identifier.";
                        return false;
                    }
                    continue;
                }

                // A bracket can denote a table structured reference or an external-workbook qualifier. The full
                // Excel grammar is required to distinguish their copy semantics, so this narrow translator stops.
                if (current == '[')
                {
                    failureReason =
                        "Bracketed structured or external references require an Excel formula parser.";
                    return false;
                }

                if (LooksLikeR1C1Reference(masterFormula, index))
                {
                    failureReason =
                        "R1C1-like reference text is outside this A1-only shared-formula translator.";
                    return false;
                }

                if (TryReadWholeColumnRange(masterFormula, index, out ReferenceToken columns) ||
                    TryReadWholeRowRange(masterFormula, index, out columns) ||
                    TryReadCellReferenceToken(masterFormula, index, out columns))
                {
                    if (!TryTranslateToken(columns, rowDelta, columnDelta, out string replacement))
                    {
                        failureReason =
                            "Translating a relative reference would move it outside the Excel worksheet grid.";
                        return false;
                    }

                    result.Append(replacement);
                    index += columns.Length;
                    continue;
                }

                result.Append(current);
                index++;
            }

            translatedFormula = result.ToString();
            return true;
        }

        internal static bool TryParseCellReference(string reference, out int row, out int column)
        {
            if (TryParseCellReference(reference, out CellCoordinate coordinate))
            {
                row = coordinate.Row;
                column = coordinate.Column;
                return true;
            }

            row = 0;
            column = 0;
            return false;
        }

        internal static bool TryParseRange(
            string reference,
            out int firstRow,
            out int firstColumn,
            out int lastRow,
            out int lastColumn)
        {
            firstRow = firstColumn = lastRow = lastColumn = 0;
            if (string.IsNullOrEmpty(reference)) return false;

            int separator = reference.IndexOf(':');
            if (separator < 0)
            {
                if (!TryParseCellReference(reference, out firstRow, out firstColumn)) return false;
                lastRow = firstRow;
                lastColumn = firstColumn;
                return true;
            }

            if (separator == 0 || separator != reference.LastIndexOf(':') ||
                separator == reference.Length - 1)
                return false;

            if (!TryParseCellReference(reference.Substring(0, separator), out firstRow, out firstColumn) ||
                !TryParseCellReference(reference.Substring(separator + 1), out lastRow, out lastColumn))
                return false;

            return firstRow <= lastRow && firstColumn <= lastColumn;
        }

        private static bool TryParseCellReference(string reference, out CellCoordinate coordinate)
        {
            coordinate = default(CellCoordinate);
            if (string.IsNullOrEmpty(reference)) return false;

            int index = 0;
            if (reference[index] == '$') index++;
            int columnStart = index;
            while (index < reference.Length && IsAsciiLetter(reference[index])) index++;
            int columnLength = index - columnStart;
            if (columnLength == 0 || columnLength > 3) return false;

            if (index < reference.Length && reference[index] == '$') index++;
            int rowStart = index;
            while (index < reference.Length && IsAsciiDigit(reference[index])) index++;
            if (rowStart == index || index != reference.Length || reference[rowStart] == '0') return false;

            int column = ParseColumn(reference, columnStart, columnLength);
            if (column < 1 || column > MaxColumns) return false;
            if (!int.TryParse(
                    reference.Substring(rowStart),
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out int row) ||
                row < 1 || row > MaxRows)
                return false;

            coordinate = new CellCoordinate(row, column);
            return true;
        }

        private static bool TryReadCellReferenceToken(
            string formula,
            int start,
            out ReferenceToken token)
        {
            token = default(ReferenceToken);
            if (!IsTokenStartBoundary(formula, start)) return false;

            int index = start;
            bool columnAbsolute = index < formula.Length && formula[index] == '$';
            if (columnAbsolute) index++;
            int columnStart = index;
            while (index < formula.Length && IsAsciiLetter(formula[index])) index++;
            int columnLength = index - columnStart;
            if (columnLength == 0 || columnLength > 3) return false;

            bool rowAbsolute = index < formula.Length && formula[index] == '$';
            if (rowAbsolute) index++;
            int rowStart = index;
            while (index < formula.Length && IsAsciiDigit(formula[index])) index++;
            if (rowStart == index || formula[rowStart] == '0') return false;

            if (!IsTokenEndBoundary(formula, index) ||
                (index < formula.Length && (formula[index] == '(' || formula[index] == '!')))
                return false;

            int column = ParseColumn(formula, columnStart, columnLength);
            if (column < 1 || column > MaxColumns) return false;
            if (!TryParsePositiveInt(formula, rowStart, index - rowStart, MaxRows, out int row)) return false;

            token = ReferenceToken.Cell(
                formula.Substring(start, index - start),
                column,
                row,
                columnAbsolute,
                rowAbsolute);
            return true;
        }

        private static bool TryReadWholeColumnRange(
            string formula,
            int start,
            out ReferenceToken token)
        {
            token = default(ReferenceToken);
            if (!IsTokenStartBoundary(formula, start)) return false;

            int index = start;
            bool firstAbsolute = index < formula.Length && formula[index] == '$';
            if (firstAbsolute) index++;
            int firstStart = index;
            while (index < formula.Length && IsAsciiLetter(formula[index])) index++;
            int firstLength = index - firstStart;
            if (firstLength == 0 || firstLength > 3 || index >= formula.Length || formula[index] != ':')
                return false;
            index++;

            bool secondAbsolute = index < formula.Length && formula[index] == '$';
            if (secondAbsolute) index++;
            int secondStart = index;
            while (index < formula.Length && IsAsciiLetter(formula[index])) index++;
            int secondLength = index - secondStart;
            if (secondLength == 0 || secondLength > 3 || !IsTokenEndBoundary(formula, index)) return false;

            int first = ParseColumn(formula, firstStart, firstLength);
            int second = ParseColumn(formula, secondStart, secondLength);
            if (first < 1 || first > MaxColumns || second < 1 || second > MaxColumns) return false;

            token = ReferenceToken.ColumnRange(
                formula.Substring(start, index - start),
                first,
                second,
                firstAbsolute,
                secondAbsolute);
            return true;
        }

        private static bool TryReadWholeRowRange(
            string formula,
            int start,
            out ReferenceToken token)
        {
            token = default(ReferenceToken);
            if (!IsTokenStartBoundary(formula, start)) return false;

            int index = start;
            bool firstAbsolute = index < formula.Length && formula[index] == '$';
            if (firstAbsolute) index++;
            int firstStart = index;
            while (index < formula.Length && IsAsciiDigit(formula[index])) index++;
            int firstLength = index - firstStart;
            if (firstLength == 0 || formula[firstStart] == '0' ||
                index >= formula.Length || formula[index] != ':')
                return false;
            index++;

            bool secondAbsolute = index < formula.Length && formula[index] == '$';
            if (secondAbsolute) index++;
            int secondStart = index;
            while (index < formula.Length && IsAsciiDigit(formula[index])) index++;
            int secondLength = index - secondStart;
            if (secondLength == 0 || formula[secondStart] == '0' || !IsTokenEndBoundary(formula, index))
                return false;

            if (!TryParsePositiveInt(formula, firstStart, firstLength, MaxRows, out int first) ||
                !TryParsePositiveInt(formula, secondStart, secondLength, MaxRows, out int second))
                return false;

            token = ReferenceToken.RowRange(
                formula.Substring(start, index - start),
                first,
                second,
                firstAbsolute,
                secondAbsolute);
            return true;
        }

        private static bool TryTranslateToken(
            ReferenceToken token,
            int rowDelta,
            int columnDelta,
            out string replacement)
        {
            replacement = null;
            switch (token.Kind)
            {
                case ReferenceTokenKind.Cell:
                    int column = token.FirstAbsolute ? token.First : token.First + columnDelta;
                    int row = token.SecondAbsolute ? token.Second : token.Second + rowDelta;
                    if (!IsColumn(column) || !IsRow(row)) return false;
                    replacement = (token.FirstAbsolute ? "$" : "") + ColumnName(column) +
                                  (token.SecondAbsolute ? "$" : "") +
                                  row.ToString(CultureInfo.InvariantCulture);
                    return true;

                case ReferenceTokenKind.ColumnRange:
                    int firstColumn = token.FirstAbsolute ? token.First : token.First + columnDelta;
                    int secondColumn = token.SecondAbsolute ? token.Second : token.Second + columnDelta;
                    if (!IsColumn(firstColumn) || !IsColumn(secondColumn)) return false;
                    replacement = (token.FirstAbsolute ? "$" : "") + ColumnName(firstColumn) + ":" +
                                  (token.SecondAbsolute ? "$" : "") + ColumnName(secondColumn);
                    return true;

                case ReferenceTokenKind.RowRange:
                    int firstRow = token.FirstAbsolute ? token.First : token.First + rowDelta;
                    int secondRow = token.SecondAbsolute ? token.Second : token.Second + rowDelta;
                    if (!IsRow(firstRow) || !IsRow(secondRow)) return false;
                    replacement = (token.FirstAbsolute ? "$" : "") +
                                  firstRow.ToString(CultureInfo.InvariantCulture) + ":" +
                                  (token.SecondAbsolute ? "$" : "") +
                                  secondRow.ToString(CultureInfo.InvariantCulture);
                    return true;

                default:
                    return false;
            }
        }

        private static bool CopyDoubleQuotedString(string formula, ref int index, StringBuilder result)
        {
            result.Append(formula[index++]);
            while (index < formula.Length)
            {
                char current = formula[index++];
                result.Append(current);
                if (current != '"') continue;
                if (index < formula.Length && formula[index] == '"')
                {
                    result.Append(formula[index++]);
                    continue;
                }
                return true;
            }
            return false;
        }

        private static bool CopySingleQuotedIdentifier(string formula, ref int index, StringBuilder result)
        {
            result.Append(formula[index++]);
            while (index < formula.Length)
            {
                char current = formula[index++];
                result.Append(current);
                if (current != '\'') continue;
                if (index < formula.Length && formula[index] == '\'')
                {
                    result.Append(formula[index++]);
                    continue;
                }
                return true;
            }
            return false;
        }

        private static bool TryValidateQualifiers(string formula, out string failureReason)
        {
            failureReason = null;
            int index = 0;
            while (index < formula.Length)
            {
                if (formula[index] == '"')
                {
                    if (!SkipDoubleQuotedString(formula, ref index))
                    {
                        failureReason = "The shared formula contains an unterminated string literal.";
                        return false;
                    }
                    continue;
                }

                if (formula[index] == '\'')
                {
                    if (!SkipSingleQuotedIdentifier(
                            formula,
                            ref index,
                            out int contentStart,
                            out int contentEnd))
                    {
                        failureReason = "The shared formula contains an unterminated quoted sheet identifier.";
                        return false;
                    }

                    if (index < formula.Length && formula[index] == '!')
                    {
                        if (ContainsUnbracketedColon(formula, contentStart, contentEnd))
                        {
                            failureReason =
                                "A quoted 3D sheet qualifier is outside this narrow A1 translator.";
                            return false;
                        }
                        index++;
                    }
                    continue;
                }

                if (formula[index] == '!')
                {
                    if (!TryValidateUnquotedQualifier(formula, index, out failureReason)) return false;
                }
                index++;
            }
            return true;
        }

        private static bool SkipDoubleQuotedString(string formula, ref int index)
        {
            index++;
            while (index < formula.Length)
            {
                if (formula[index++] != '"') continue;
                if (index < formula.Length && formula[index] == '"')
                {
                    index++;
                    continue;
                }
                return true;
            }
            return false;
        }

        private static bool SkipSingleQuotedIdentifier(
            string formula,
            ref int index,
            out int contentStart,
            out int contentEnd)
        {
            contentStart = ++index;
            while (index < formula.Length)
            {
                if (formula[index++] != '\'') continue;
                if (index < formula.Length && formula[index] == '\'')
                {
                    index++;
                    continue;
                }
                contentEnd = index - 1;
                return true;
            }
            contentEnd = formula.Length;
            return false;
        }

        private static bool ContainsUnbracketedColon(string formula, int start, int end)
        {
            int bracketDepth = 0;
            for (int index = start; index < end; index++)
            {
                if (formula[index] == '[')
                {
                    bracketDepth++;
                    continue;
                }
                if (formula[index] == ']')
                {
                    if (bracketDepth > 0) bracketDepth--;
                    continue;
                }
                if (formula[index] == ':' && bracketDepth == 0) return true;
            }
            return false;
        }

        private static bool TryValidateUnquotedQualifier(
            string formula,
            int bang,
            out string failureReason)
        {
            failureReason = null;
            int qualifierStart = bang;
            while (qualifierStart > 0 && IsUnquotedQualifierPartAt(formula, qualifierStart - 1))
                qualifierStart--;
            if (qualifierStart == bang)
            {
                failureReason = "The shared formula contains a malformed unquoted sheet qualifier.";
                return false;
            }

            int colon = qualifierStart - 1;
            if (colon < 0 || formula[colon] != ':') return true;

            int leftStart = colon;
            while (leftStart > 0 && IsUnquotedQualifierPartAt(formula, leftStart - 1)) leftStart--;
            if (leftStart == colon || (leftStart > 0 && formula[leftStart - 1] == '!')) return true;

            failureReason = "An unquoted 3D sheet qualifier is outside this narrow A1 translator.";
            return false;
        }

        private static bool IsUnquotedQualifierPartAt(string formula, int index) =>
            IsIdentifierPartAt(formula, index) || formula[index] == '$';

        private static bool IsTokenStartBoundary(string formula, int start)
        {
            if (start < 0 || start >= formula.Length) return false;
            char current = formula[start];
            if (current != '$' && !IsAsciiLetter(current) && !IsAsciiDigit(current)) return false;
            if (start == 0) return true;

            return !IsIdentifierPartAt(formula, start - 1);
        }

        private static bool LooksLikeR1C1Reference(string formula, int start)
        {
            if (!IsTokenStartBoundary(formula, start) ||
                (formula[start] != 'R' && formula[start] != 'r'))
                return false;

            int index = start + 1;
            while (index < formula.Length && IsAsciiDigit(formula[index])) index++;
            if (index >= formula.Length ||
                (formula[index] != 'C' && formula[index] != 'c'))
                return false;

            index++;
            while (index < formula.Length && IsAsciiDigit(formula[index])) index++;
            return IsTokenEndBoundary(formula, index);
        }

        private static bool IsTokenEndBoundary(string formula, int end)
        {
            if (end >= formula.Length) return true;
            return !IsIdentifierPartAt(formula, end);
        }

        private static bool IsIdentifierPartAt(string text, int index)
        {
            if (index < 0 || index >= text.Length) return false;
            char value = text[index];
            if (IsAsciiLetter(value) || IsAsciiDigit(value) ||
                value == '_' || value == '.' || value == '\\' || value == '?')
                return true;
            if (value < 128) return false;

            // CharUnicodeInfo needs the high-surrogate position. Any malformed surrogate is treated as an
            // identifier constituent, which suppresses rewriting instead of guessing at a boundary.
            int scalarStart = index;
            if (char.IsLowSurrogate(value))
            {
                if (index == 0 || !char.IsHighSurrogate(text[index - 1])) return true;
                scalarStart = index - 1;
            }
            else if (char.IsHighSurrogate(value) &&
                     (index + 1 >= text.Length || !char.IsLowSurrogate(text[index + 1])))
            {
                return true;
            }

            UnicodeCategory category;
            try
            {
                category = CharUnicodeInfo.GetUnicodeCategory(text, scalarStart);
            }
            catch (ArgumentException)
            {
                return true;
            }

            switch (category)
            {
                case UnicodeCategory.UppercaseLetter:
                case UnicodeCategory.LowercaseLetter:
                case UnicodeCategory.TitlecaseLetter:
                case UnicodeCategory.ModifierLetter:
                case UnicodeCategory.OtherLetter:
                case UnicodeCategory.NonSpacingMark:
                case UnicodeCategory.SpacingCombiningMark:
                case UnicodeCategory.EnclosingMark:
                case UnicodeCategory.DecimalDigitNumber:
                case UnicodeCategory.LetterNumber:
                case UnicodeCategory.OtherNumber:
                case UnicodeCategory.ConnectorPunctuation:
                case UnicodeCategory.Format:
                    return true;
                default:
                    return false;
            }
        }

        private static int ParseColumn(string text, int start, int length)
        {
            int result = 0;
            for (int i = 0; i < length; i++)
            {
                char value = text[start + i];
                int digit = char.ToUpperInvariant(value) - 'A' + 1;
                result = result * 26 + digit;
            }
            return result;
        }

        private static bool TryParsePositiveInt(
            string text,
            int start,
            int length,
            int maximum,
            out int value)
        {
            value = 0;
            if (length <= 0) return false;
            for (int i = 0; i < length; i++)
            {
                int digit = text[start + i] - '0';
                if (digit < 0 || digit > 9 || value > (maximum - digit) / 10) return false;
                value = value * 10 + digit;
            }
            return value >= 1 && value <= maximum;
        }

        private static string ColumnName(int column)
        {
            var result = new StringBuilder(3);
            while (column > 0)
            {
                column--;
                result.Insert(0, (char)('A' + column % 26));
                column /= 26;
            }
            return result.ToString();
        }

        private static bool IsAsciiLetter(char value) =>
            (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z');

        private static bool IsAsciiDigit(char value) => value >= '0' && value <= '9';
        private static bool IsRow(int value) => value >= 1 && value <= MaxRows;
        private static bool IsColumn(int value) => value >= 1 && value <= MaxColumns;

        private struct CellCoordinate
        {
            internal CellCoordinate(int row, int column)
            {
                Row = row;
                Column = column;
            }

            internal int Row { get; }
            internal int Column { get; }
        }

        private enum ReferenceTokenKind
        {
            Cell,
            ColumnRange,
            RowRange,
        }

        private struct ReferenceToken
        {
            private ReferenceToken(
                ReferenceTokenKind kind,
                string original,
                int first,
                int second,
                bool firstAbsolute,
                bool secondAbsolute)
            {
                Kind = kind;
                Original = original;
                First = first;
                Second = second;
                FirstAbsolute = firstAbsolute;
                SecondAbsolute = secondAbsolute;
            }

            internal ReferenceTokenKind Kind { get; }
            internal string Original { get; }
            internal int First { get; }
            internal int Second { get; }
            internal bool FirstAbsolute { get; }
            internal bool SecondAbsolute { get; }
            internal int Length => Original == null ? 0 : Original.Length;

            internal static ReferenceToken Cell(
                string original,
                int column,
                int row,
                bool columnAbsolute,
                bool rowAbsolute) =>
                new ReferenceToken(
                    ReferenceTokenKind.Cell,
                    original,
                    column,
                    row,
                    columnAbsolute,
                    rowAbsolute);

            internal static ReferenceToken ColumnRange(
                string original,
                int first,
                int second,
                bool firstAbsolute,
                bool secondAbsolute) =>
                new ReferenceToken(
                    ReferenceTokenKind.ColumnRange,
                    original,
                    first,
                    second,
                    firstAbsolute,
                    secondAbsolute);

            internal static ReferenceToken RowRange(
                string original,
                int first,
                int second,
                bool firstAbsolute,
                bool secondAbsolute) =>
                new ReferenceToken(
                    ReferenceTokenKind.RowRange,
                    original,
                    first,
                    second,
                    firstAbsolute,
                    secondAbsolute);
        }
    }
}
