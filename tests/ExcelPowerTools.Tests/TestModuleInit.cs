using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelPowerTools.Tests
{
    internal static class TestModuleInit
    {
        // Register Windows-1252 (and other code pages) so CsvEncoding.Windows1252() works under net8.
        [ModuleInitializer]
        internal static void Init()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
    }
}
