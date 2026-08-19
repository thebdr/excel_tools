using System.Text;
using ExcelPowerTools.Core.Csv;
using Xunit;

namespace ExcelPowerTools.Tests.Csv
{
    public class CsvSourceFingerprintTests
    {
        [Fact]
        public void Identical_bytes_have_the_same_fingerprint()
        {
            byte[] a = Encoding.UTF8.GetBytes("a;b\r\n1;2\r\n");
            byte[] b = Encoding.UTF8.GetBytes("a;b\r\n1;2\r\n");

            Assert.Equal(CsvSourceFingerprint.FromBytes(a), CsvSourceFingerprint.FromBytes(b));
            Assert.True(CsvSourceFingerprint.FromBytes(a).Matches(b));
        }

        [Fact]
        public void Same_length_external_edit_is_detected()
        {
            byte[] before = Encoding.UTF8.GetBytes("a;b\r\n1;2\r\n");
            byte[] after = Encoding.UTF8.GetBytes("a;b\r\n1;9\r\n");

            Assert.Equal(before.Length, after.Length);
            Assert.NotEqual(CsvSourceFingerprint.FromBytes(before), CsvSourceFingerprint.FromBytes(after));
        }

        [Fact]
        public void Null_and_empty_content_share_the_empty_fingerprint()
        {
            Assert.Equal(CsvSourceFingerprint.FromBytes(null), CsvSourceFingerprint.FromBytes(new byte[0]));
        }
    }
}
