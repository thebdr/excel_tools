using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ExcelPowerTools.Core.Comparison.Identifiers
{
    /// <summary>
    /// Deterministic structural identity based on SHA-256 and ordinal UTF-8 components.
    ///
    /// Encoding v1 writes a fixed namespace marker, algorithm version, domain and component count followed by
    /// every component. Strings are UTF-8 with a four-byte big-endian length prefix; null is encoded as -1 and
    /// is therefore distinct from empty. No Unicode or case normalization is performed.
    /// </summary>
    public static class StableId
    {
        public const string AlgorithmVersion = "sha256-lp-utf8-v1";
        private const string NamespaceMarker = "ExcelPowerTools.Comparison";
        private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(false, true);

        public static string Compute(string domain, params string[] components)
        {
            if (string.IsNullOrWhiteSpace(domain))
                throw new ArgumentException("A stable-ID domain is required.", nameof(domain));

            components = components ?? Array.Empty<string>();

            using (var payload = new MemoryStream())
            {
                WriteString(payload, NamespaceMarker);
                WriteString(payload, AlgorithmVersion);
                WriteString(payload, domain);
                WriteInt32(payload, components.Length);
                for (int i = 0; i < components.Length; i++)
                    WriteString(payload, components[i]);

                payload.Position = 0;
                using (var sha = SHA256.Create())
                    return ToHex(sha.ComputeHash(payload));
            }
        }

        /// <summary>Returns the content address specified by the design: SHA-256 of the exact bytes.</summary>
        public static string ForBytes(byte[] bytes)
        {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            using (var sha = SHA256.Create())
                return ToHex(sha.ComputeHash(bytes));
        }

        /// <summary>Compatibility name for <see cref="ForBytes"/>.</summary>
        public static string ComputeByteContentId(byte[] bytes) => ForBytes(bytes);

        public static bool IsSha256(string value)
        {
            if (value == null || value.Length != 64)
                return false;

            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                bool isHex = (c >= '0' && c <= '9') ||
                             (c >= 'A' && c <= 'F') ||
                             (c >= 'a' && c <= 'f');
                if (!isHex)
                    return false;
            }

            return true;
        }

        private static void WriteString(Stream stream, string value)
        {
            if (value == null)
            {
                WriteInt32(stream, -1);
                return;
            }

            byte[] bytes;
            try
            {
                bytes = StrictUtf8.GetBytes(value);
            }
            catch (EncoderFallbackException exception)
            {
                throw new ArgumentException(
                    "Stable ID components must be well-formed UTF-16 strings.",
                    nameof(value),
                    exception);
            }
            WriteInt32(stream, bytes.Length);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static void WriteInt32(Stream stream, int value)
        {
            unchecked
            {
                stream.WriteByte((byte)(value >> 24));
                stream.WriteByte((byte)(value >> 16));
                stream.WriteByte((byte)(value >> 8));
                stream.WriteByte((byte)value);
            }
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
