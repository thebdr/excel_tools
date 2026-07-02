using System;
using System.Text;

namespace ExcelPowerTools.Core.Csv
{
    /// <summary>
    /// Encoding detection and BOM handling. Detects UTF-8 (with/without BOM), UTF-16 LE/BE, and falls
    /// back to Windows-1252 for byte streams that aren't valid UTF-8 (common for German/French exports).
    /// </summary>
    public static class CsvEncoding
    {
        private const char Bom = (char)0xFEFF;   // U+FEFF (zero-width no-break space / BOM)

        public static (Encoding encoding, bool hasBom) Detect(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return (new UTF8Encoding(false), false);

            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return (new UTF8Encoding(false), true);
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return (Encoding.Unicode, true);            // UTF-16 LE
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return (Encoding.BigEndianUnicode, true);   // UTF-16 BE

            Encoding bomless16 = DetectBomlessUtf16(bytes);
            if (bomless16 != null)
                return (bomless16, false);

            if (IsValidUtf8(bytes))
                return (new UTF8Encoding(false), false);

            return (Windows1252(), false);
        }

        /// <summary>
        /// Heuristic for UTF-16 text that lacks a BOM (e.g. some Notepad/tool exports): ASCII-heavy UTF-16
        /// has a NUL as every other byte. A strong bias of NULs toward odd byte positions => LE, even => BE.
        /// </summary>
        private static Encoding DetectBomlessUtf16(byte[] bytes)
        {
            if (bytes.Length < 4 || (bytes.Length % 2) != 0) return null;

            int sample = Math.Min(bytes.Length, 4096);
            int zerosEven = 0, zerosOdd = 0, pairs = 0;
            for (int i = 0; i + 1 < sample; i += 2)
            {
                if (bytes[i] == 0) zerosEven++;
                if (bytes[i + 1] == 0) zerosOdd++;
                pairs++;
            }
            if (pairs == 0) return null;

            double evenFrac = (double)zerosEven / pairs;
            double oddFrac = (double)zerosOdd / pairs;

            if (oddFrac > 0.30 && oddFrac > evenFrac * 3) return Encoding.Unicode;            // UTF-16 LE
            if (evenFrac > 0.30 && evenFrac > oddFrac * 3) return Encoding.BigEndianUnicode;  // UTF-16 BE
            return null;
        }

        /// <summary>Decodes bytes with the detected encoding and strips a leading U+FEFF (BOM) if present.</summary>
        public static string DecodeText(byte[] bytes, Encoding encoding, bool hasBom)
        {
            if (bytes == null || bytes.Length == 0) return string.Empty;
            string s = encoding.GetString(bytes);
            if (s.Length > 0 && s[0] == Bom) s = s.Substring(1);
            return s;
        }

        /// <summary>The BOM bytes to prepend on save (empty when hasBom is false or the encoding has none).</summary>
        public static byte[] Preamble(Encoding encoding, bool hasBom)
        {
            if (!hasBom || encoding == null) return Array.Empty<byte>();
            switch (encoding.CodePage)
            {
                case 65001: return new byte[] { 0xEF, 0xBB, 0xBF };  // UTF-8
                case 1200: return new byte[] { 0xFF, 0xFE };         // UTF-16 LE
                case 1201: return new byte[] { 0xFE, 0xFF };         // UTF-16 BE
                default: return encoding.GetPreamble();
            }
        }

        /// <summary>
        /// Windows-1252. Available natively on .NET Framework; on modern .NET the caller must have
        /// registered CodePagesEncodingProvider (the test project does this in a module initializer).
        /// </summary>
        public static Encoding Windows1252() => Encoding.GetEncoding(1252);

        private static bool IsValidUtf8(byte[] bytes)
        {
            try
            {
                new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes);
                return true;
            }
            catch (DecoderFallbackException)
            {
                return false;
            }
        }
    }
}
