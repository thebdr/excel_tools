using System;
using System.Security.Cryptography;

namespace ExcelPowerTools.Core.Csv
{
    /// <summary>
    /// Content identity for a source file. Timestamps and lengths are useful display metadata, but only a
    /// byte-level digest can reliably detect an external edit (including same-length edits and restored dates).
    /// </summary>
    public sealed class CsvSourceFingerprint : IEquatable<CsvSourceFingerprint>
    {
        public long Length { get; }
        public string Sha256 { get; }

        private CsvSourceFingerprint(long length, string sha256)
        {
            Length = length;
            Sha256 = sha256 ?? "";
        }

        public static CsvSourceFingerprint FromBytes(byte[] bytes)
        {
            bytes = bytes ?? Array.Empty<byte>();
            using (var sha = SHA256.Create())
            {
                byte[] digest = sha.ComputeHash(bytes);
                return new CsvSourceFingerprint(bytes.LongLength, ToHex(digest));
            }
        }

        public bool Matches(byte[] bytes) => Equals(FromBytes(bytes));

        public bool Equals(CsvSourceFingerprint other) =>
            other != null && Length == other.Length &&
            string.Equals(Sha256, other.Sha256, StringComparison.Ordinal);

        public override bool Equals(object obj) => Equals(obj as CsvSourceFingerprint);

        public override int GetHashCode()
        {
            unchecked { return (Length.GetHashCode() * 397) ^ Sha256.GetHashCode(); }
        }

        private static string ToHex(byte[] bytes)
        {
            var chars = new char[bytes.Length * 2];
            const string hex = "0123456789ABCDEF";
            for (int i = 0; i < bytes.Length; i++)
            {
                chars[i * 2] = hex[bytes[i] >> 4];
                chars[i * 2 + 1] = hex[bytes[i] & 0x0F];
            }
            return new string(chars);
        }
    }
}
