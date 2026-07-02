using ExcelDna.Integration;
using ExcelPowerTools.Commands;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPowerTools
{
    /// <summary>
    /// Add-in lifecycle: subscribes to the application-level WorkbookOpen event so a natively-opened
    /// .csv can be re-routed through the corruption-safe CSV pipeline.
    /// </summary>
    public sealed class AddIn : IExcelAddIn
    {
        // Held for the lifetime of the add-in so the COM event sink is not garbage-collected.
        private static Excel.Application _app;

        public void AutoOpen()
        {
            _app = (Excel.Application)ExcelDnaUtil.Application;
            _app.WorkbookOpen += CsvCommands.OnWorkbookOpen;

            // Show a readable message instead of a bare #VALUE! if a UDF throws.
            ExcelIntegration.RegisterUnhandledExceptionHandler(
                ex => "!ERR: " + ((ex as System.Exception)?.Message ?? ex?.ToString()));
        }

        public void AutoClose()
        {
            if (_app == null) return;
            _app.WorkbookOpen -= CsvCommands.OnWorkbookOpen;
            _app = null;
        }
    }
}
