using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelPowerTools.Core.Csv
{
    /// <summary>Pure naming/filtering helpers for the Excel delimited-file workspace.</summary>
    public static class CsvWorkspaceLogic
    {
        private const string SheetAnchorPrefix = "__PT_CSV_SHEET_";
        private static readonly HashSet<string> SupportedExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".csv", ".tsv", ".txt" };

        public static bool IsSupportedPath(string path) =>
            !string.IsNullOrWhiteSpace(path) && SupportedExtensions.Contains(Path.GetExtension(path));

        public static string UniqueSheetName(string path, IEnumerable<string> existingNames)
        {
            string stem;
            try { stem = Path.GetFileNameWithoutExtension(path ?? ""); }
            catch { stem = ""; }
            stem = SanitizeSheetName(stem);
            if (stem.Length == 0) stem = "Data";

            var used = new HashSet<string>(existingNames ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            // Excel reserves this name for its change-history worksheet even when no such sheet is visible.
            used.Add("History");
            string candidate = TrimTo(stem, 31);
            if (!used.Contains(candidate)) return candidate;

            for (int n = 2; ; n++)
            {
                string suffix = " (" + n + ")";
                candidate = TrimTo(stem, 31 - suffix.Length) + suffix;
                if (!used.Contains(candidate)) return candidate;
            }
        }

        public static string EscapeSheetNameForSubAddress(string sheetName) =>
            "'" + (sheetName ?? "").Replace("'", "''") + "'!A1";

        /// <summary>
        /// Invariant whole-sheet reference used by a hidden workbook-scoped worksheet anchor. Excel keeps
        /// this reference attached to the worksheet across cell cuts and arbitrary row/column deletion.
        /// </summary>
        public static string SheetAnchorFormula(string sheetName) =>
            "='" + (sheetName ?? "").Replace("'", "''") + "'!$1:$1048576";

        /// <summary>Whole-column anchor for source metadata; insertion moves it, deletion breaks it closed.</summary>
        public static string SheetColumnAnchorFormula(string sheetName, int column)
        {
            if (column < 1 || column > 16384)
                throw new ArgumentOutOfRangeException(nameof(column));
            string name = ExcelColumnName(column);
            return "='" + (sheetName ?? "").Replace("'", "''") + "'!$" + name + ":$" + name;
        }

        public static string SheetAnchorName(string sheetId)
        {
            if (!IsWorkspaceId(sheetId)) throw new ArgumentException("Invalid worksheet identity.", nameof(sheetId));
            return SheetAnchorPrefix + sheetId.ToLowerInvariant();
        }

        /// <summary>
        /// Excel keeps the worksheet qualifier when deleting the one column referenced by an anchor, for
        /// example <c>='My Sheet'!#REF!</c>. This lets pre-delete cleanup identify the owning sheet even after
        /// both the marker and the live range reference have been destroyed.
        /// </summary>
        public static bool BrokenSheetAnchorTargetsSheet(string refersTo, string sheetName)
        {
            if (string.IsNullOrEmpty(refersTo) || string.IsNullOrEmpty(sheetName) ||
                refersTo.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) < 0) return false;

            string quoted = "='" + sheetName.Replace("'", "''") + "'!#REF!";
            string bare = "=" + sheetName + "!#REF!";
            return string.Equals(refersTo, quoted, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(refersTo, bare, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// A copied source sheet can carry the original sheet identity into another workbook. Use that
        /// identity as a close-safety fallback only when the runtime session is not still owned by a
        /// different open workbook.
        /// </summary>
        public static bool UseSheetIdentityForClosingWorkbook(
            bool sheetIdentityMatches,
            bool knownOwnerWorkbookIsOpen,
            bool knownOwnerWorkbookIsClosing)
        {
            return sheetIdentityMatches &&
                (!knownOwnerWorkbookIsOpen || knownOwnerWorkbookIsClosing);
        }

        public static bool KeepRuntimeSessionDuringCleanup(
            bool workspaceIdOwnerIsOpen, bool rememberedNameOwnerIsOpen) =>
            workspaceIdOwnerIsOpen || rememberedNameOwnerIsOpen;

        /// <summary>
        /// Extracts this add-in's worksheet identity from either a short workbook name or Excel's
        /// workbook/sheet-qualified <c>Name.Name</c> representation.
        /// </summary>
        public static string TryParseSheetAnchorId(string definedName)
        {
            if (string.IsNullOrEmpty(definedName)) return null;
            int bang = definedName.LastIndexOf('!');
            string local = bang >= 0 ? definedName.Substring(bang + 1) : definedName;
            if (!local.StartsWith(SheetAnchorPrefix, StringComparison.OrdinalIgnoreCase)) return null;

            string id = local.Substring(SheetAnchorPrefix.Length);
            return IsWorkspaceId(id) ? id.ToLowerInvariant() : null;
        }

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "";
            foreach (char c in new[] { '[', ']', ':', '*', '?', '/', '\\' }) name = name.Replace(c, '_');
            name = name.Trim();
            // Excel rejects a name whose first or last character is an apostrophe.
            name = name.Trim('\'');
            return name;
        }

        public static string CanonicalPath(string path)
        {
            try { return Path.GetFullPath(path ?? ""); }
            catch { return path ?? ""; }
        }

        public static bool SameSourcePath(string left, string right) =>
            string.Equals(CanonicalPath(left), CanonicalPath(right), StringComparison.OrdinalIgnoreCase);

        public static bool IsWorkspaceId(string value)
        {
            if (value == null || value.Length != 32) return false;
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')))
                    return false;
            }
            return true;
        }

        private static string TrimTo(string value, int max)
        {
            if (value.Length <= max) return value;
            max = Math.Max(0, max);
            // Excel's limit is measured in UTF-16 code units, but never return a broken surrogate pair.
            if (max > 0 && max < value.Length && char.IsHighSurrogate(value[max - 1]) && char.IsLowSurrogate(value[max]))
                max--;
            return value.Substring(0, max);
        }

        private static string ExcelColumnName(int column)
        {
            var chars = new char[3];
            int index = chars.Length;
            while (column > 0)
            {
                column--;
                chars[--index] = (char)('A' + column % 26);
                column /= 26;
            }
            return new string(chars, index, chars.Length - index);
        }
    }
}
